// TouchokuLauncher.cs — 当直勤務表入力の起動係（ランチャー）。20260925
//
// 管理者 2026-09-25「もうちょっと起動速くなる？」:
//   electron-builder の portable exe は起動のたびに中身（約355MB）を一時フォルダへ展開していて、画面が出るまで
//   約19秒かかっていた（展開済みのアプリを直接起動すると約4秒）。
//   そこで、共有の月フォルダに置く「当直勤務表入力_R8年N月.exe」をこの小さな起動係に置き換え、
//   展開済みのアプリを各 PC の %LOCALAPPDATA%\Touchoku\app\<版>\ に1回だけ置いて、2回目からはそれを起動する。
//
// 動き:
//   1. 自分と同じフォルダの touchoku_app.ver（版の名前1行）と touchoku_app.zip（展開済みアプリの zip）を読む。
//   2. %LOCALAPPDATA%\Touchoku\app\<版>\ に .ok があればそのまま起動。無ければ「準備しています」を出して展開する
//      （同時に2つ起動しても壊れないよう、版ごとの Mutex で1つずつ）。古い版のフォルダは消す（使用中なら残す）。
//   3. 環境変数 PORTABLE_EXECUTABLE_DIR に自分のフォルダ（共有の月フォルダ）を入れて起動する。
//      アプリ（main.js）はこれを「自分の置き場」とみなすので、共有の予定表・設定・起動場所の確認は今までどおり効く。
//   4. 起動したらすぐ終わる（共有の exe を掴みっぱなしにしない＝アプリを開いたままでも更新できる）。
//
// ビルド: tools\build_launcher.bat（Windows に入っている .NET Framework の csc.exe でコンパイル）

using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyTitle("当直勤務表入力")]
[assembly: System.Reflection.AssemblyProduct("当直勤務表入力")]

static class TouchokuLauncher
{
    const string APP_EXE = "当直勤務表入力.exe";
    const string ZIP_NAME = "touchoku_app.zip";
    const string VER_NAME = "touchoku_app.ver";
    const string TITLE = "当直勤務表入力";

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    static extern bool SetProcessDPIAware();

    [STAThread]
    static int Main(string[] args)
    {
        try { SetProcessDPIAware(); } catch { }      // 「準備しています」の窓がぼやけないように（表示の拡大に対応）
        string here = Path.GetDirectoryName(Application.ExecutablePath);
        try
        {
            string verPath = Path.Combine(here, VER_NAME);
            string zipPath = Path.Combine(here, ZIP_NAME);
            if (!File.Exists(verPath) || !File.Exists(zipPath))
            {
                Fail("このフォルダにアプリの本体（" + ZIP_NAME + "）がありません。\n管理者に連絡してください。\n\n" + here);
                return 1;
            }
            string ver = File.ReadAllText(verPath).Trim();
            if (ver.Length == 0 || ver.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                Fail("版の名前（" + VER_NAME + "）が読めません。管理者に連絡してください。");
                return 1;
            }
            string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Touchoku", "app");
            string dir = Path.Combine(root, ver);
            string exe = Path.Combine(dir, APP_EXE);

            if (!Ready(dir, exe))
            {
                using (var mtx = new Mutex(false, "Touchoku_app_extract_" + ver))
                {
                    // 20260925: 待ちきれなかった（別の起動がまだ展開している）ときは、展開せずに終わる。
                    //   取れていないミューテックスは ReleaseMutex しない（例外になる）
                    bool got = false;
                    try { got = mtx.WaitOne(TimeSpan.FromMinutes(5)); }
                    catch (AbandonedMutexException) { got = true; }   // 前の持ち主が落ちた＝取れている
                    if (!got)
                    {
                        Fail("別の起動が準備中です。少し待ってからもう一度開いてください。");
                        return 1;
                    }
                    try
                    {
                        if (!Ready(dir, exe)) Extract(zipPath, root, dir);   // 待っている間に別の起動が置いていれば、しない
                    }
                    finally { mtx.ReleaseMutex(); }
                }
                CleanOld(root, ver);
            }
            if (!Ready(dir, exe))
            {
                Fail("アプリの準備に失敗しました。もう一度起動してください。\n何度も失敗するときは管理者に連絡してください。");
                return 1;
            }

            var psi = new ProcessStartInfo(exe);
            psi.UseShellExecute = false;
            psi.WorkingDirectory = dir;
            psi.EnvironmentVariables["PORTABLE_EXECUTABLE_DIR"] = here;
            psi.Arguments = string.Join(" ", Array.ConvertAll(args, a => "\"" + a.Replace("\"", "\\\"") + "\""));
            Process.Start(psi);
            return 0;
        }
        catch (Exception ex)
        {
            Fail("起動できませんでした。\n" + ex.Message);
            return 1;
        }
    }

    static bool Ready(string dir, string exe)
    {
        return File.Exists(Path.Combine(dir, ".ok")) && File.Exists(exe);
    }

    // 展開: 版のフォルダへ直接展開し、最後に .ok を書く（.ok が無い版は使わない＝途中で止まっても中途半端な版を起動しない）。
    //   20260925: 一時フォルダに展開してから名前を変える作りをやめた。展開したばかりのファイルをウイルス対策（EDR）が
    //   掴んでいると、フォルダの名前変更（Directory.Move）が失敗するため。失敗したら作りかけのフォルダを消す（消せなければ残す。
    //   .ok が無いので次の起動がやり直す）。「準備しています」を出す
    static void Extract(string zipPath, string root, string dir)
    {
        Directory.CreateDirectory(root);
        Exception err = null;
        var form = new Form();
        form.Text = TITLE;
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.StartPosition = FormStartPosition.CenterScreen;
        form.ControlBox = false;
        form.AutoScaleMode = AutoScaleMode.Dpi;
        form.Font = new Font("Yu Gothic UI", 12f);
        form.ClientSize = new Size(520, 130);
        var lbl = new Label();
        lbl.Text = "当直勤務表アプリを準備しています。\n初回と、アプリが新しくなったときだけ、30秒ほどかかります。";
        lbl.SetBounds(20, 16, 480, 56);
        var bar = new ProgressBar();
        bar.Style = ProgressBarStyle.Marquee;
        bar.SetBounds(20, 84, 480, 22);
        form.Controls.Add(lbl);
        form.Controls.Add(bar);
        form.Shown += (s, e) =>
        {
            var th = new Thread(() =>
            {
                try
                {
                    if (Directory.Exists(dir)) Directory.Delete(dir, true);   // 前の作りかけ（.ok が無い）
                    ZipFile.ExtractToDirectory(zipPath, dir);
                    if (!File.Exists(Path.Combine(dir, APP_EXE)))
                        throw new Exception("zip に " + APP_EXE + " がありません");
                    File.WriteAllText(Path.Combine(dir, ".ok"), DateTime.Now.ToString("s"));   // 必ず最後
                }
                catch (Exception ex) { err = ex; }
                try { form.BeginInvoke((Action)(() => form.Close())); } catch { }
            });
            th.IsBackground = true;
            th.Start();
        };
        Application.EnableVisualStyles();
        Application.Run(form);
        if (err != null)
        {
            try { if (Directory.Exists(dir)) Directory.Delete(dir, true); } catch { }
            throw new Exception("展開に失敗: " + err.Message);
        }
    }

    // 古い版のフォルダを消す。20260925: 途中まで消えた版を残さない（使える版を半分だけ消すと、.ok があるのに起動できない）。
    //   ・その版を展開中の起動係がいれば触らない（版ごとの Mutex が取れない）
    //   ・その版のアプリが動いていれば触らない（exe を排他で開けない）
    //   ・まず名前を「<版>.<pid>.old」に変え（中のファイルを誰かが開いていれば失敗する＝版はそのまま残る）、それから消す。
    //     .old や以前の作りの .tmp<pid> は版ではない残骸なので、途中までしか消えなくても害はない（次の起動で続きを消す）
    static void CleanOld(string root, string keep)
    {
        string[] dirs;
        try { dirs = Directory.GetDirectories(root); } catch { return; }
        foreach (var d in dirs)
        {
            string name = Path.GetFileName(d);
            if (string.Equals(name, keep, StringComparison.OrdinalIgnoreCase)) continue;
            try
            {
                if (name.EndsWith(".old", StringComparison.OrdinalIgnoreCase) || name.IndexOf(".tmp", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Directory.Delete(d, true);
                    continue;
                }
                using (var mtx = new Mutex(false, "Touchoku_app_extract_" + name))
                {
                    bool got = false;
                    try { got = mtx.WaitOne(0); } catch (AbandonedMutexException) { got = true; }
                    if (!got) continue;                       // その版をいま展開している起動係がいる
                    try
                    {
                        if (InUse(d)) continue;
                        string trash = d + "." + Process.GetCurrentProcess().Id + ".old";
                        if (Directory.Exists(trash)) continue;
                        try { Directory.Move(d, trash); } catch { continue; }   // 使用中のファイルがある → 版ごと残す
                        try { Directory.Delete(trash, true); } catch { }
                    }
                    finally { mtx.ReleaseMutex(); }
                }
            }
            catch { }
        }
    }

    // その版のアプリが動いているか: exe を排他（FileShare.None）で開けなければ使用中
    static bool InUse(string d)
    {
        string exe = Path.Combine(d, APP_EXE);
        if (!File.Exists(exe)) return false;
        try
        {
            using (new FileStream(exe, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
            return false;
        }
        catch (IOException) { return true; }
        catch (UnauthorizedAccessException) { return true; }
    }

    static void Fail(string msg)
    {
        MessageBox.Show(msg, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
