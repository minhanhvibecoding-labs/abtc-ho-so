# Mở hộp thoại "Chọn thư mục" kiểu Windows hiện đại và ghi đường dẫn đã chọn vào file kết quả (UTF-8, rỗng nếu người dùng hủy).
# Do server.ps1 gọi trong một tiến trình riêng (-STA) để không phụ thuộc trạng thái của máy chủ web.
param([Parameter(Mandatory)][string]$KetQua, [string]$BatDau = '', [string]$TieuDe = 'Chọn folder hồ sơ ABTC')

$cs = @'
using System;
using System.Runtime.InteropServices;
using System.Threading;
namespace AbtcDlg {
  [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")] class FileOpenDialogRCW { }
  [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  interface IShellItem {
    void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
    void GetParent(out IShellItem ppsi);
    void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
    void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
    void Compare(IShellItem psi, uint hint, out int piOrder);
  }
  [ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  interface IFileOpenDialog {
    [PreserveSig] int Show(IntPtr parent);
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
    void SetClientGuid(ref Guid guid);
    void ClearClientData();
    void SetFilter(IntPtr pFilter);
    void GetResults(out IntPtr ppenum);
    void GetSelectedItems(out IntPtr ppsai);
  }
  public static class Chon {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    static extern IShellItem SHCreateItemFromParsingName(string pszPath, IntPtr pbc, ref Guid riid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);

    // Luồng nền: dò hộp thoại theo tiêu đề (tối đa ~8 giây) rồi đặt lên trên cùng.
    // Cần vì tiến trình chạy ẩn thường không được Windows cho nổi lên trước cửa sổ trình duyệt.
    static void Log(string s) { try { System.IO.File.AppendAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "abtc_chon_folder.log"), DateTime.Now.ToString("s") + " " + s + Environment.NewLine); } catch { } }
    static void DuaLenTren(string title) {
      var th = new Thread(() => {
        for (int i = 0; i < 80; i++) {
          IntPtr h = FindWindow(null, title);
          if (h != IntPtr.Zero) {
            SetWindowPos(h, new IntPtr(-1), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040);   // HWND_TOPMOST, giữ nguyên cỡ/vị trí
            bool fg = SetForegroundWindow(h); Log("tìm thấy hộp thoại sau " + (i * 100) + "ms, foreground=" + fg); return;
          }
          Thread.Sleep(100);
        }
        Log("KHÔNG tìm thấy cửa sổ tiêu đề: " + title);
      });
      th.IsBackground = true; th.Start();
    }

    public static string Folder(string title, string okLabel, string start, IntPtr owner) {
      var d = (IFileOpenDialog)new FileOpenDialogRCW();
      uint o; d.GetOptions(out o);
      d.SetOptions(o | 0x20 /*PICKFOLDERS*/ | 0x40 /*FORCEFILESYSTEM*/ | 0x8 /*NOCHANGEDIR*/);
      d.SetTitle(title); d.SetOkButtonLabel(okLabel);
      if (!string.IsNullOrEmpty(start)) { try { Guid g = typeof(IShellItem).GUID; d.SetFolder(SHCreateItemFromParsingName(start, IntPtr.Zero, ref g)); } catch { } }
      DuaLenTren(title);
      if (d.Show(owner) != 0) return null;
      IShellItem it; d.GetResult(out it);
      string p; it.GetDisplayName(0x80058000 /*SIGDN_FILESYSPATH*/, out p); return p;
    }
  }
}
'@

$duongDan = $null
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
# Cửa sổ chủ vô hình, luôn nổi trên cùng, được kích hoạt bằng AppActivate (API chính thống) để hộp thoại hiện lên trước cửa sổ trình duyệt
$chu = New-Object Windows.Forms.Form
$chu.Text = 'ABTC chọn folder'
$chu.TopMost = $true; $chu.ShowInTaskbar = $false; $chu.FormBorderStyle = 'None'; $chu.Opacity = 0
$chu.StartPosition = 'Manual'; $chu.Location = New-Object Drawing.Point(-9999, -9999); $chu.Size = New-Object Drawing.Size(1, 1)
$chu.Show(); $chu.Activate()
try { Add-Type -AssemblyName Microsoft.VisualBasic; [Microsoft.VisualBasic.Interaction]::AppActivate($PID) } catch { }
try {
  try {
    Add-Type -TypeDefinition $cs -ErrorAction Stop
    $duongDan = [AbtcDlg.Chon]::Folder($TieuDe, 'Chọn folder này', $BatDau, $chu.Handle)
  } catch {
    Add-Content -LiteralPath (Join-Path $env:TEMP 'abtc_chon_folder.log') -Value ("$(Get-Date -Format s) LỖI hộp thoại mới: " + $_.Exception.Message) -Encoding UTF8
    # Dự phòng: hộp thoại cây thư mục cổ điển
    $d = New-Object Windows.Forms.FolderBrowserDialog; $d.Description = $TieuDe; $d.ShowNewFolderButton = $false
    if ($BatDau -and (Test-Path -LiteralPath $BatDau)) { $d.SelectedPath = $BatDau }
    if ($d.ShowDialog($chu) -eq 'OK') { $duongDan = $d.SelectedPath }
  }
} finally { $chu.Close(); $chu.Dispose() }
[IO.File]::WriteAllText($KetQua, [string]$duongDan, (New-Object Text.UTF8Encoding($false)))
