using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using CvGenerator.Core.Services;

namespace CvGenerator.Docx.Services;

public sealed class PdfConversionService
{
    public string? FindLibreOffice(string? configuredPath)
    {
        if (!string.IsNullOrWhiteSpace(configuredPath))
        {
            var configured = ResolveCandidate(configuredPath);
            if (configured != null)
                return configured;
        }

        foreach (var candidate in EnumerateCandidates())
        {
            var resolved = ResolveCandidate(candidate);
            if (resolved != null)
                return resolved;
        }

        return null;
    }

    private static IEnumerable<string> EnumerateCandidates()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            foreach (var fromRegistry in ReadWindowsInstallPaths())
                yield return Path.Combine(fromRegistry, "program", "soffice.exe");

            var roots = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                Environment.GetEnvironmentVariable("ProgramW6432") ?? "",
                @"C:\Program Files",
                @"C:\Program Files (x86)"
            }.Where(r => !string.IsNullOrWhiteSpace(r)).Distinct(StringComparer.OrdinalIgnoreCase);

            foreach (var root in roots)
            {
                yield return Path.Combine(root, "LibreOffice", "program", "soffice.exe");

                foreach (var dir in SafeEnumerateLibreOfficeDirs(root))
                    yield return Path.Combine(dir, "program", "soffice.exe");
            }
        }
        else
        {
            yield return "/usr/bin/soffice";
            yield return "/usr/bin/libreoffice";
            yield return "/snap/bin/libreoffice";
            yield return "/usr/lib/libreoffice/program/soffice";
        }

        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var dir in pathEnv.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            yield return Path.Combine(dir, "soffice");
            yield return Path.Combine(dir, "soffice.exe");
            yield return Path.Combine(dir, "libreoffice");
            yield return Path.Combine(dir, "libreoffice.exe");
        }
    }

    private static IEnumerable<string> SafeEnumerateLibreOfficeDirs(string root)
    {
        try
        {
            if (!Directory.Exists(root))
                return Array.Empty<string>();
            return Directory.EnumerateDirectories(root, "LibreOffice*").ToList();
        }
        catch
        {
            return Array.Empty<string>();
        }
    }

    private static IEnumerable<string> ReadWindowsInstallPaths()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            yield break;

        string?[] keys =
        [
            @"HKEY_LOCAL_MACHINE\SOFTWARE\LibreOffice\UNO\InstallPath",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath",
            @"HKEY_CURRENT_USER\SOFTWARE\LibreOffice\UNO\InstallPath"
        ];

        foreach (var key in keys)
        {
            string? value = null;
            try
            {
                value = Microsoft.Win32.Registry.GetValue(key!, "", null) as string;
            }
            catch
            {
                // ignore
            }

            if (!string.IsNullOrWhiteSpace(value))
                yield return value!;
        }
    }

    private static string? ResolveCandidate(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return null;

        if (Directory.Exists(path))
        {
            var nested = new[]
            {
                Path.Combine(path, "soffice.exe"),
                Path.Combine(path, "program", "soffice.exe"),
                Path.Combine(path, "soffice"),
                Path.Combine(path, "program", "soffice")
            };
            foreach (var n in nested)
            {
                if (File.Exists(n))
                    return n;
            }

            return null;
        }

        return File.Exists(path) ? path : null;
    }

    public string ConvertToPdf(string docxPath, string outputDirectory, string? libreOfficePath, IssueCollector issues)
    {
        Directory.CreateDirectory(outputDirectory);
        var resumePdf = Path.Combine(outputDirectory, "resume.pdf");

        // Prefer Microsoft Word on Windows — keeps DOCX fonts. Tests: CVGEN_SKIP_WORD_PDF=1
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) &&
            !IsWordPdfSkipped())
        {
            // Never open resume.docx in Word — a stuck Word dialog locks it (white/slow open).
            // Convert a temp copy, then force-kill only Word processes we started.
            if (TryConvertWithWordIsolated(docxPath, resumePdf, issues))
                return resumePdf;
        }

        var exe = FindLibreOffice(libreOfficePath);
        if (exe is null)
        {
            issues.Error(
                "LibreOfficeMissing",
                "Pdf",
                "LibreOffice (soffice.exe) not found. Install LibreOffice, or install Microsoft Word for best PDF font match.");
            return string.Empty;
        }

        try
        {
            const string pdfFilter =
                "pdf:writer_pdf_Export:" +
                "{" +
                "\"SelectPdfVersion\":{\"type\":\"long\",\"value\":\"2\"}," +
                "\"EmbedStandardFonts\":{\"type\":\"boolean\",\"value\":\"true\"}" +
                "}";

            var profileDir = Path.Combine(Path.GetTempPath(), "CvGenerator-lo-profile");
            Directory.CreateDirectory(profileDir);
            var profileUri = new Uri(
                profileDir.EndsWith(Path.DirectorySeparatorChar)
                    ? profileDir
                    : profileDir + Path.DirectorySeparatorChar).AbsoluteUri;

            var psi = new ProcessStartInfo
            {
                FileName = exe,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            psi.ArgumentList.Add($"-env:UserInstallation={profileUri}");
            psi.ArgumentList.Add("--headless");
            psi.ArgumentList.Add("--nologo");
            psi.ArgumentList.Add("--nofirststartwizard");
            psi.ArgumentList.Add("--norestore");
            psi.ArgumentList.Add("--convert-to");
            psi.ArgumentList.Add(pdfFilter);
            psi.ArgumentList.Add("--outdir");
            psi.ArgumentList.Add(outputDirectory);
            psi.ArgumentList.Add(docxPath);

            using var process = Process.Start(psi);
            if (process is null)
            {
                issues.Error("LibreOfficeStartFailed", "Pdf", "Failed to start LibreOffice process.", detail: exe);
                return string.Empty;
            }

            var stdout = new StringBuilder();
            var stderr = new StringBuilder();
            process.OutputDataReceived += (_, e) => { if (e.Data != null) stdout.AppendLine(e.Data); };
            process.ErrorDataReceived += (_, e) => { if (e.Data != null) stderr.AppendLine(e.Data); };
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            if (!process.WaitForExit(120_000))
            {
                try { process.Kill(entireProcessTree: true); } catch { /* ignore */ }
                issues.Error("LibreOfficeTimeout", "Pdf", "LibreOffice conversion timed out after 120s.", detail: exe);
                return string.Empty;
            }

            var expected = Path.Combine(
                outputDirectory,
                Path.GetFileNameWithoutExtension(docxPath) + ".pdf");

            if (process.ExitCode != 0 || !File.Exists(expected))
            {
                issues.Error(
                    "LibreOfficeFailed",
                    "Pdf",
                    $"PDF conversion failed (exit {process.ExitCode}).",
                    detail: $"{exe}\n{stdout}\n{stderr}");
                return string.Empty;
            }

            if (!string.Equals(expected, resumePdf, StringComparison.OrdinalIgnoreCase))
            {
                if (File.Exists(resumePdf))
                    File.Delete(resumePdf);
                File.Move(expected, resumePdf);
            }

            issues.Warning(
                "PdfViaLibreOffice",
                "Pdf",
                "PDF via LibreOffice. Fonts can still look different from Word (e.g. Calibri→Carlito) unless those exact fonts are installed. Install Microsoft Word for the closest match, or install the template fonts on this PC.");

            return resumePdf;
        }
        catch (Exception ex)
        {
            issues.Error("LibreOfficeFailed", "Pdf", $"PDF conversion error: {ex.Message}", detail: ex.ToString());
            return string.Empty;
        }
    }

    private static bool IsWordPdfSkipped()
        => string.Equals(
            Environment.GetEnvironmentVariable("CVGEN_SKIP_WORD_PDF"),
            "1",
            StringComparison.Ordinal);

    /// <summary>
    /// Out-of-process Word via cscript on a temp copy of the DOCX.
    /// Avoids locking resume.docx and avoids in-process COM modal dialogs in the WPF host.
    /// </summary>
    [SupportedOSPlatform("windows")]
    private static bool TryConvertWithWordIsolated(string docxPath, string pdfPath, IssueCollector issues)
    {
        var workDir = Path.Combine(Path.GetTempPath(), "CvGenerator-word-" + Guid.NewGuid().ToString("N"));
        var scriptPath = Path.Combine(workDir, "export.vbs");
        var tempDocx = Path.Combine(workDir, "source.docx");
        var tempPdf = Path.Combine(workDir, "out.pdf");
        var beforePids = SnapshotWordPids();

        try
        {
            Directory.CreateDirectory(workDir);
            File.Copy(docxPath, tempDocx, overwrite: true);
            TryDelete(pdfPath);
            TryDelete(tempPdf);

            WriteWordExportScript(scriptPath, tempDocx, tempPdf);

            foreach (var cscript in EnumerateCscriptHosts())
            {
                // Any leftover Word from a previous attempt still blocks Quit with dialogs.
                KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());

                TryDelete(tempPdf);
                var psi = new ProcessStartInfo
                {
                    FileName = cscript,
                    ArgumentList = { "//nologo", "//T:90", scriptPath },
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    WorkingDirectory = workDir
                };

                using var process = Process.Start(psi);
                if (process is null)
                    continue;

                if (!process.WaitForExit(95_000))
                {
                    try { process.Kill(entireProcessTree: true); } catch { /* ignore */ }
                    KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());
                    continue;
                }

                // Always tear down Word we started — dialogs make Quit hang otherwise.
                KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());

                if (process.ExitCode == 0 && File.Exists(tempPdf) && new FileInfo(tempPdf).Length > 0)
                {
                    File.Copy(tempPdf, pdfPath, overwrite: true);
                    issues.Info("PdfViaWord", "Pdf", "PDF exported via Microsoft Word (fonts match the DOCX).");
                    return true;
                }
            }

            issues.Info(
                "WordPdfUnavailable",
                "Pdf",
                "Word PDF export failed or timed out — falling back to LibreOffice.");
            return false;
        }
        catch (Exception ex)
        {
            KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());
            issues.Info(
                "WordPdfUnavailable",
                "Pdf",
                "Word PDF export failed — falling back to LibreOffice.",
                detail: ex.Message);
            return false;
        }
        finally
        {
            KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());
            TryDeleteDirectory(workDir);
        }
    }

    [SupportedOSPlatform("windows")]
    private static void WriteWordExportScript(string scriptPath, string docxPath, string pdfPath)
    {
        var docxEsc = EscapeVbs(docxPath);
        var pdfEsc = EscapeVbs(pdfPath);
        // Silent open: no conversion/encoding dialogs; never touch the real resume.docx.
        // Saved=True + Close(0) avoids "save changes?" and stuck Quit dialogs.
        var script = $$"""
            On Error Resume Next
            Dim w, docs, d, i
            Set w = CreateObject("Word.Application")
            If Err.Number <> 0 Or w Is Nothing Then
              WScript.Quit 2
            End If

            w.Visible = False
            w.DisplayAlerts = 0
            w.ScreenUpdating = False
            w.AutomationSecurity = 3
            Err.Clear

            Set docs = w.Documents
            Set d = docs.Open("{{docxEsc}}", False, True, False, , , False, , , , , False, False, , True)
            If Err.Number <> 0 Or d Is Nothing Then
              Call ForceQuit(w)
              WScript.Quit 3
            End If

            Err.Clear
            d.ExportAsFixedFormat "{{pdfEsc}}", 17, False, 0, 0, , , 0, True, True, 0, True, False, False
            If Err.Number <> 0 Then
              Call SafeCloseAll(w)
              Call ForceQuit(w)
              WScript.Quit 4
            End If

            Call SafeCloseAll(w)
            Call ForceQuit(w)
            WScript.Quit 0

            Sub SafeCloseAll(app)
              On Error Resume Next
              Dim j, doc
              For j = app.Documents.Count To 1 Step -1
                Set doc = app.Documents(j)
                doc.Saved = True
                doc.Close 0
              Next
            End Sub

            Sub ForceQuit(app)
              On Error Resume Next
              Call SafeCloseAll(app)
              app.DisplayAlerts = 0
              app.Quit 0
            End Sub
            """;
        File.WriteAllText(scriptPath, script, Encoding.ASCII);
    }

    [SupportedOSPlatform("windows")]
    private static HashSet<int> SnapshotWordPids()
    {
        var set = new HashSet<int>();
        try
        {
            foreach (var p in Process.GetProcessesByName("WINWORD"))
            {
                try { set.Add(p.Id); }
                catch { /* ignore */ }
                finally
                {
                    try { p.Dispose(); } catch { /* ignore */ }
                }
            }
        }
        catch
        {
            // ignore
        }

        return set;
    }

    [SupportedOSPlatform("windows")]
    private static void KillWordPids(HashSet<int> pids)
    {
        if (pids.Count == 0)
            return;

        foreach (var pid in pids)
        {
            try
            {
                using var p = Process.GetProcessById(pid);
                if (!p.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase))
                    continue;
                p.Kill(entireProcessTree: true);
                p.WaitForExit(5_000);
            }
            catch
            {
                // already exited / access denied
            }
        }
    }

    private static IEnumerable<string> EnumerateCscriptHosts()
    {
        var windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        // Prefer native bitness first; SysWOW64 as fallback for 32-bit Office.
        yield return Path.Combine(windir, "System32", "cscript.exe");
        var wow = Path.Combine(windir, "SysWOW64", "cscript.exe");
        if (File.Exists(wow))
            yield return wow;
    }

    private static string EscapeVbs(string path)
        => path.Replace("\"", "\"\"", StringComparison.Ordinal);

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // ignore
        }
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
                Directory.Delete(path, recursive: true);
        }
        catch
        {
            // ignore locked temp leftovers
        }
    }
}
