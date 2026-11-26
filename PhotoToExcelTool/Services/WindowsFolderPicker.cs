// ==============================================================================
// 檔案名稱：Services/WindowsFolderPicker.cs
// 學習重點：
// 1. 原生 Shell COM 介面 (IFileOpenDialog + FOS_PICKFOLDERS)：直接呼叫 Windows 內建的現代版檔案總管對話框，
//    完全不依賴笨重的 WinForms/WPF 外部原生 DLL，保持單一獨立執行檔 (.exe) 輕量且零外部相依。
// 2. 突破背景程序焦點限制 (WS_EX_TOPMOST + AttachThreadInput)：
//    當使用者在瀏覽器點擊按鈕時，焦點在 Chrome/Edge 上，Windows 預設禁止背景 Console 程序搶奪焦點，
//    會導致對話框躲在瀏覽器視窗背後看似「按鈕沒反應」。
//    此處建立一個位於瀏覽器中央的 WS_EX_TOPMOST 置頂宿主視窗，並短暫透過 AttachThreadInput 切換前景焦點，
//    確保資料夾選擇視窗 100% 彈出在瀏覽器正前方。
// ==============================================================================

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace PhotoToExcelTool.Services;

/// <summary>
/// 提供本機 Windows 10/11 原生現代版資料夾選擇對話框服務（純 Win32 + Shell COM 實作，零外部 DLL 相依）。
/// </summary>
public static class WindowsFolderPicker
{
    private const uint FOS_NOCHANGEDIR = 0x00000008;
    private const uint FOS_PICKFOLDERS = 0x00000020;
    private const uint FOS_FORCEFILESYSTEM = 0x00000040;
    private const uint FOS_PATHMUSTEXIST = 0x00000800;

    private const uint SIGDN_FILESYSPATH = 0x80058000;
    private const int ERROR_CANCELLED_HR = unchecked((int)0x800704C7);

    private const uint WS_POPUP = 0x80000000;
    private const uint WS_VISIBLE = 0x10000000;
    private const uint WS_EX_TOPMOST = 0x00000008;

    private static readonly Guid ClsidFileOpenDialog = new("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7");
    private static readonly Guid IidIShellItem = new("43826d1e-e718-42ee-bc55-a1e261c37bfe");

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateWindowExW(
        uint dwExStyle,
        string lpClassName,
        string lpWindowName,
        uint dwStyle,
        int x,
        int y,
        int nWidth,
        int nHeight,
        IntPtr hWndParent,
        IntPtr hMenu,
        IntPtr hInstance,
        IntPtr lpParam);

    [DllImport("user32.dll")]
    private static extern bool DestroyWindow(IntPtr hWnd);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
    private static extern int SHCreateItemFromParsingName(
        string pszPath,
        IntPtr pbc,
        [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
        out IShellItem? ppv);

    /// <summary>
    /// 在專屬 STA 執行緒上開啟 Windows 原生現代資料夾選擇器，確保置頂顯示於瀏覽器前方。
    /// </summary>
    public static string? PickFolder(string? initialDirectory = null)
    {
        string? selectedPath = null;
        Exception? threadEx = null;

        var staThread = new Thread(() =>
        {
            IntPtr ownerHwnd = IntPtr.Zero;
            IFileOpenDialog? dialog = null;
            IShellItem? initItem = null;
            IShellItem? resultItem = null;

            try
            {
                // 1. 取得目前使用者正在操作的瀏覽器視窗位置，讓對話框彈在同一個螢幕正中央
                int spawnX = 360;
                int spawnY = 220;
                IntPtr fgHwnd = GetForegroundWindow();
                if (fgHwnd != IntPtr.Zero && GetWindowRect(fgHwnd, out RECT rect) && rect.Right > rect.Left)
                {
                    spawnX = rect.Left + (rect.Right - rect.Left) / 3;
                    spawnY = rect.Top + (rect.Bottom - rect.Top) / 4;
                }

                // 2. 建立 1x1 的 WS_EX_TOPMOST 置頂宿主視窗，讓子對話框繼承最上層 Z-Order，絕不會躲在瀏覽器背後
                ownerHwnd = CreateWindowExW(
                    WS_EX_TOPMOST,
                    "STATIC",
                    "SiCFolderPickerHost",
                    WS_POPUP | WS_VISIBLE,
                    spawnX,
                    spawnY,
                    1,
                    1,
                    IntPtr.Zero,
                    IntPtr.Zero,
                    IntPtr.Zero,
                    IntPtr.Zero);

                if (ownerHwnd != IntPtr.Zero)
                {
                    uint fgThreadId = fgHwnd != IntPtr.Zero ? GetWindowThreadProcessId(fgHwnd, out _) : 0;
                    uint curThreadId = GetCurrentThreadId();
                    bool attached = false;
                    if (fgThreadId != 0 && fgThreadId != curThreadId)
                    {
                        attached = AttachThreadInput(fgThreadId, curThreadId, true);
                    }
                    BringWindowToTop(ownerHwnd);
                    SetForegroundWindow(ownerHwnd);
                    if (attached)
                    {
                        AttachThreadInput(fgThreadId, curThreadId, false);
                    }
                }

                // 3. 實例化 Windows 現代版 IFileOpenDialog COM 物件
                Type? dialogType = Type.GetTypeFromCLSID(ClsidFileOpenDialog, throwOnError: true);
                dialog = (IFileOpenDialog)Activator.CreateInstance(dialogType!)!;

                dialog.GetOptions(out uint options);
                options |= FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_NOCHANGEDIR;
                dialog.SetOptions(options);
                dialog.SetTitle("請選擇工編資料夾");

                if (!string.IsNullOrWhiteSpace(initialDirectory))
                {
                    string cleanDir = initialDirectory.Trim().Trim('"');
                    if (Directory.Exists(cleanDir))
                    {
                        int initHr = SHCreateItemFromParsingName(cleanDir, IntPtr.Zero, IidIShellItem, out initItem);
                        if (initHr >= 0 && initItem != null)
                        {
                            dialog.SetFolder(initItem);
                        }
                    }
                }

                // 4. 顯示對話框（若使用者按取消，回傳 ERROR_CANCELLED_HR，正常返回 null）
                int hr = dialog.Show(ownerHwnd);
                if (hr == ERROR_CANCELLED_HR || hr < 0)
                {
                    return;
                }

                dialog.GetResult(out resultItem);
                if (resultItem != null)
                {
                    resultItem.GetDisplayName(SIGDN_FILESYSPATH, out string? path);
                    if (!string.IsNullOrWhiteSpace(path))
                    {
                        selectedPath = path;
                    }
                }
            }
            catch (Exception ex)
            {
                threadEx = ex;
            }
            finally
            {
                if (resultItem != null) Marshal.ReleaseComObject(resultItem);
                if (initItem != null) Marshal.ReleaseComObject(initItem);
                if (dialog != null) Marshal.ReleaseComObject(dialog);
                if (ownerHwnd != IntPtr.Zero) DestroyWindow(ownerHwnd);
            }
        });

        staThread.IsBackground = true;
        staThread.SetApartmentState(ApartmentState.STA);
        staThread.Start();
        staThread.Join();

        if (threadEx != null)
        {
            throw new InvalidOperationException($"開啟資料夾選擇器失敗: {threadEx.Message}", threadEx);
        }

        return selectedPath;
    }

    // 完整對應 Windows SDK ShObjIdl_core.h 中 IModalWindow + IFileDialog + IFileOpenDialog 的 VTable 順序
    [ComImport]
    [Guid("d57c7288-d4ad-4768-be02-9d969532d960")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog
    {
        [PreserveSig]
        int Show(IntPtr parent);

        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
        void SetClientGuid([MarshalAs(UnmanagedType.LPStruct)] Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
        void GetResults(out IntPtr ppenum);
        void GetSelectedItems(out IntPtr ppsai);
    }

    [ComImport]
    [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        void BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }
}
