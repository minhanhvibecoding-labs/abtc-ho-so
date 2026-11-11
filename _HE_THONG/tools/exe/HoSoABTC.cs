// Trình khởi chạy "Hồ sơ ABTC": một file .exe chứa sẵn phần mềm (app.zip nhúng trong exe).
//  - Lần đầu: giải nén vào Documents\CƠ SỞ DỮ LIỆU 2026, tạo lối tắt Desktop, mở ứng dụng.
//  - Lần sau: ứng dụng đang chạy → mở lại cửa sổ; chưa chạy → khởi động máy chủ cục bộ (server.ps1).
//  - Bản exe mới hơn bản đã cài (do exe cài) → hỏi rồi cập nhật phần mềm; KHÔNG đụng dữ liệu hồ sơ, Excel tổng hợp, cache.
//  - Thư mục đã có phần mềm nhưng không do exe cài (máy phát triển) → không ghi đè gì, chỉ mở ứng dụng.
// Tham số thử nghiệm: --goc "<thư mục>" (cài vào thư mục khác), --khong-mo (chỉ cài, không mở ứng dụng).
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

[assembly: AssemblyTitle("Hồ sơ ABTC")]
[assembly: AssemblyDescription("Phần mềm Hồ sơ ABTC – Phòng Lãnh sự (chạy trên máy, không gửi dữ liệu ra ngoài)")]
[assembly: AssemblyProduct("Hồ sơ ABTC")]
[assembly: AssemblyVersion("2026.11.11.0")]
[assembly: AssemblyFileVersion("2026.11.11.0")]

static class HoSoABTC
{
    const string TEN_GOC = "CƠ SỞ DỮ LIỆU 2026";
    const string PHIEN_BAN = "2026.11.11";
    const string TIEU_DE = "Hồ sơ ABTC";

    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            string goc = null; bool khongMo = false; bool thu = false;
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "--goc" && i + 1 < args.Length) { goc = args[++i]; }
                else if (args[i] == "--khong-mo") { khongMo = true; }
                else if (args[i] == "--thu") { thu = true; khongMo = true; }   // kiểm thử: không hộp thoại, không lối tắt, không mở ứng dụng; cập nhật luôn
            }
            string exeDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');
            if (goc == null)
            {
                // exe đặt ngay trong thư mục đã cài → dùng thư mục đó; còn lại → Documents\CƠ SỞ DỮ LIỆU 2026
                goc = File.Exists(Path.Combine(exeDir, @"_HE_THONG\ABTC_Tools.ps1"))
                    ? exeDir
                    : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), TEN_GOC);
            }
            string tools = Path.Combine(goc, @"_HE_THONG\ABTC_Tools.ps1");
            string verFile = Path.Combine(goc, @"_HE_THONG\phien_ban_app.txt");
            bool daCo = File.Exists(tools);
            bool doExeCai = File.Exists(verFile);
            string verCu = doExeCai ? File.ReadAllText(verFile).Trim() : "";

            if (!daCo)
            {
                GiaiNen(goc, false);
                File.WriteAllText(verFile, PHIEN_BAN, new UTF8Encoding(false));
                string exeCai = SaoChepExe(goc);
                if (!thu) { TaoLoiTat(exeCai); }
                if (!thu) MessageBox.Show(
                    "Đã cài phần mềm Hồ sơ ABTC (bản " + PHIEN_BAN + ") vào:\n" + goc +
                    "\n\nĐã tạo lối tắt \"Hồ sơ ABTC\" trên Desktop.\n\nỨng dụng sẽ mở ngay sau đây. Cửa sổ PowerShell thu nhỏ dưới taskbar là máy chủ của ứng dụng – giữ nó mở trong lúc dùng.",
                    TIEU_DE, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (doExeCai && string.CompareOrdinal(verCu, PHIEN_BAN) < 0)
            {
                DialogResult r = thu ? DialogResult.Yes : MessageBox.Show(
                    "Máy đang có bản " + verCu + ". Cập nhật phần mềm lên bản " + PHIEN_BAN + "?\n\nChỉ thay phần mềm; dữ liệu hồ sơ, kho, Excel tổng hợp giữ nguyên.\nNên tắt ứng dụng đang mở trước khi cập nhật.",
                    TIEU_DE, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (r == DialogResult.Yes)
                {
                    GiaiNen(goc, true);
                    File.WriteAllText(verFile, PHIEN_BAN, new UTF8Encoding(false));
                    SaoChepExe(goc);
                }
            }

            if (khongMo) { return 0; }

            string url = TimMayChu();
            if (url != null) { MoCuaSo(url); return 0; }

            string server = Path.Combine(goc, @"_HE_THONG\app\server.ps1");
            if (!File.Exists(server)) { throw new FileNotFoundException("Không thấy " + server); }
            ProcessStartInfo psi = new ProcessStartInfo("powershell.exe",
                "-NoProfile -ExecutionPolicy Bypass -Command \"& '" + server.Replace("'", "''") + "'; Write-Host ''; pause\"");
            psi.WorkingDirectory = goc;
            psi.WindowStyle = ProcessWindowStyle.Minimized;
            psi.UseShellExecute = true;
            Process.Start(psi);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.ToString());
            MessageBox.Show("Không mở được ứng dụng Hồ sơ ABTC:\n\n" + ex.Message, TIEU_DE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }

    // Giải nén app.zip (nhúng trong exe). capNhat = true: ghi đè phần mềm, nhưng KHÔNG ghi đè file Excel mẫu người dùng có thể đã chỉnh.
    static void GiaiNen(string goc, bool capNhat)
    {
        Directory.CreateDirectory(goc);
        Stream zs = Assembly.GetExecutingAssembly().GetManifestResourceStream("app.zip");
        if (zs == null) { throw new InvalidOperationException("Gói phần mềm trong exe bị hỏng (thiếu app.zip)."); }
        using (zs)
        using (ZipArchive za = new ZipArchive(zs, ZipArchiveMode.Read, false, Encoding.UTF8))
        {
            string gocDay = Path.GetFullPath(goc).TrimEnd('\\') + "\\";
            foreach (ZipArchiveEntry e in za.Entries)
            {
                string ten = e.FullName.Replace('/', '\\');
                if (ten.EndsWith("\\")) { Directory.CreateDirectory(Path.Combine(goc, ten)); continue; }
                string dich = Path.GetFullPath(Path.Combine(goc, ten));
                if (!dich.StartsWith(gocDay, StringComparison.OrdinalIgnoreCase)) { continue; }   // không giải nén ra ngoài thư mục cài
                if (capNhat && dich.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) && File.Exists(dich)) { continue; }
                Directory.CreateDirectory(Path.GetDirectoryName(dich));
                e.ExtractToFile(dich, true);
            }
        }
        Directory.CreateDirectory(Path.Combine(goc, @"KẾT QUẢ\CƠ SỞ DỮ LIỆU"));
        Directory.CreateDirectory(Path.Combine(goc, "Folder Hồ sơ ABTC Mẫu"));   // kho hồ sơ mặc định (trống) để dùng ngay
    }

    // Chép exe vào thư mục cài (để lối tắt trỏ tới bản cố định, không phụ thuộc file trong Downloads)
    static string SaoChepExe(string goc)
    {
        string nguon = Assembly.GetExecutingAssembly().Location;
        string dich = Path.Combine(goc, "Hồ sơ ABTC.exe");
        try
        {
            if (!string.Equals(Path.GetFullPath(nguon), Path.GetFullPath(dich), StringComparison.OrdinalIgnoreCase)) { File.Copy(nguon, dich, true); }
            return dich;
        }
        catch { return nguon; }
    }

    static void TaoLoiTat(string exe)
    {
        try
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            Type t = Type.GetTypeFromProgID("WScript.Shell");
            object sh = Activator.CreateInstance(t);
            object lnk = t.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, sh, new object[] { Path.Combine(desktop, "Hồ sơ ABTC.lnk") });
            Type lt = lnk.GetType();
            lt.InvokeMember("TargetPath", BindingFlags.SetProperty, null, lnk, new object[] { exe });
            lt.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, lnk, new object[] { Path.GetDirectoryName(exe) });
            lt.InvokeMember("IconLocation", BindingFlags.SetProperty, null, lnk, new object[] { exe + ",0" });
            lt.InvokeMember("Description", BindingFlags.SetProperty, null, lnk, new object[] { "Mở ứng dụng Hồ sơ ABTC" });
            lt.InvokeMember("Save", BindingFlags.InvokeMethod, null, lnk, null);
        }
        catch { }   // không tạo được lối tắt thì vẫn dùng exe bình thường
    }

    // Ứng dụng đang chạy? (máy chủ cục bộ cổng 8765–8775 trả lời /api/tong-quan)
    static string TimMayChu()
    {
        for (int p = 8765; p <= 8775; p++)
        {
            try
            {
                HttpWebRequest rq = (HttpWebRequest)WebRequest.Create("http://localhost:" + p + "/api/tong-quan");
                rq.Timeout = 800; rq.ReadWriteTimeout = 800; rq.Proxy = null;
                using (HttpWebResponse rs = (HttpWebResponse)rq.GetResponse())
                using (StreamReader sr = new StreamReader(rs.GetResponseStream(), Encoding.UTF8))
                {
                    if (sr.ReadToEnd().Contains("\"goc\"")) { return "http://localhost:" + p + "/"; }
                }
            }
            catch { }
        }
        return null;
    }

    static void MoCuaSo(string url)
    {
        try { Process.Start(new ProcessStartInfo("msedge.exe", "--app=" + url) { UseShellExecute = true }); }
        catch { Process.Start(url); }
    }
}
