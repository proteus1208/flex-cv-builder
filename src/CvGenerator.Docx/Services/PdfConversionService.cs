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

        // Standard publisher for DOCX→PDF on Windows is Microsoft Word
        // (File → Save As → PDF / ExportAsFixedFormat). LibreOffice is fallback only —
        // it often drops shapes/bars and substitutes fonts.
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) &&
            !IsWordPdfSkipped())
        {
            if (TryConvertWithWord(docxPath, resumePdf, issues))
                return resumePdf;
        }

        var exe = FindLibreOffice(libreOfficePath);
        if (exe is null)
        {
            issues.Error(
                "LibreOfficeMissing",
                "Pdf",
                "Word PDF export failed and LibreOffice (soffice.exe) was not found. Fix Word automation or install LibreOffice.");
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
                "PDF via LibreOffice (not Word). Bars/shapes and fonts often differ from the DOCX. Close stuck Word dialogs, then regenerate so Word Save-As-PDF can run.");

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
    /// Uses Word's built-in PDF publisher (same as File → Save As → PDF) on a temp copy.
    /// Tries 32-bit cscript first — Office 2016 is usually 32-bit and 64-bit cscript cannot create it.
    /// </summary>
    [SupportedOSPlatform("windows")]
    private static bool TryConvertWithWord(string docxPath, string pdfPath, IssueCollector issues)
    {
        var workDir = Path.Combine(Path.GetTempPath(), "CvGenerator-word-" + Guid.NewGuid().ToString("N"));
        var scriptPath = Path.Combine(workDir, "export.vbs");
        var statusPath = Path.Combine(workDir, "status.txt");
        var tempDocx = Path.Combine(workDir, "source.docx");
        var tempPdf = Path.Combine(workDir, "out.pdf");
        var beforePids = SnapshotWordPids();
        var attemptLog = new StringBuilder();

        try
        {
            Directory.CreateDirectory(workDir);
            File.Copy(docxPath, tempDocx, overwrite: true);
            TryDelete(pdfPath);
            WriteWordExportScript(scriptPath, tempDocx, tempPdf, statusPath);

            foreach (var cscript in EnumerateCscriptHosts())
            {
                KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());
                TryDelete(tempPdf);
                TryDelete(statusPath);

                var psi = new ProcessStartInfo
                {
                    FileName = cscript,
                    // No //T — Word 2016 cold start can exceed 90s; we enforce timeout ourselves.
                    ArgumentList = { "//nologo", scriptPath },
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    WorkingDirectory = workDir
                };

                using var process = Process.Start(psi);
                if (process is null)
                {
                    attemptLog.AppendLine($"{cscript}: failed to start");
                    continue;
                }

                if (!process.WaitForExit(180_000))
                {
                    try { process.Kill(entireProcessTree: true); } catch { /* ignore */ }
                    KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());
                    attemptLog.AppendLine($"{cscript}: timed out after 180s");
                    continue;
                }

                // Drain pipes after exit to avoid blocking the child while running.
                var stderrText = "";
                try { stderrText = process.StandardError.ReadToEnd(); } catch { /* ignore */ }

                KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());

                var status = File.Exists(statusPath) ? File.ReadAllText(statusPath).Trim() : "(no status)";
                attemptLog.AppendLine($"{cscript}: exit={process.ExitCode}; {status}");
                if (!string.IsNullOrWhiteSpace(stderrText))
                    attemptLog.AppendLine(stderrText.Trim());

                if (process.ExitCode == 0 && File.Exists(tempPdf) && new FileInfo(tempPdf).Length > 0)
                {
                    File.Copy(tempPdf, pdfPath, overwrite: true);
                    issues.Info(
                        "PdfViaWord",
                        "Pdf",
                        "PDF published via Microsoft Word (File → Save As → PDF). Layout/fonts match the DOCX.");
                    return true;
                }
            }

            issues.Warning(
                "WordPdfUnavailable",
                "Pdf",
                "Word Save-As-PDF failed — falling back to LibreOffice (bars/shapes may be missing).",
                detail: attemptLog.ToString().Trim());
            return false;
        }
        catch (Exception ex)
        {
            KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());
            issues.Warning(
                "WordPdfUnavailable",
                "Pdf",
                "Word Save-As-PDF failed — falling back to LibreOffice (bars/shapes may be missing).",
                detail: ex.ToString());
            return false;
        }
        finally
        {
            KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());
            TryDeleteDirectory(workDir);
        }
    }

    [SupportedOSPlatform("windows")]
    private static void WriteWordExportScript(string scriptPath, string docxPath, string pdfPath, string statusPath)
    {
        var docxEsc = EscapeVbs(docxPath);
        var pdfEsc = EscapeVbs(pdfPath);
        var statusEsc = EscapeVbs(statusPath);

        // Same engine as Word UI: File → Save As → PDF (wdFormatPDF = 17).
        // Fallback: ExportAsFixedFormat. Always operate on a temp copy.
        var script = $$"""
            Option Explicit
            On Error Resume Next

            Dim fso, statusFile, w, d, errOpen, errSave, errExport
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set statusFile = fso.CreateTextFile("{{statusEsc}}", True)
            statusFile.WriteLine "start"

            Set w = CreateObject("Word.Application")
            If Err.Number <> 0 Or w Is Nothing Then
              statusFile.WriteLine "create_failed err=" & Err.Number & " " & Err.Description
              statusFile.Close
              WScript.Quit 2
            End If
            statusFile.WriteLine "word_created version=" & w.Version

            w.Visible = False
            w.DisplayAlerts = 0
            w.ScreenUpdating = False
            Err.Clear
            w.AutomationSecurity = 3
            Err.Clear

            Set d = w.Documents.Open("{{docxEsc}}", False, True, False)
            errOpen = Err.Number
            If errOpen <> 0 Or d Is Nothing Then
              statusFile.WriteLine "open_failed err=" & errOpen & " " & Err.Description
              Call ForceQuit(w)
              statusFile.Close
              WScript.Quit 3
            End If
            statusFile.WriteLine "opened"

            Err.Clear
            d.SaveAs2 "{{pdfEsc}}", 17
            errSave = Err.Number
            If errSave <> 0 Or Not fso.FileExists("{{pdfEsc}}") Then
              statusFile.WriteLine "saveas2_failed err=" & errSave & " " & Err.Description
              Err.Clear
              d.SaveAs "{{pdfEsc}}", 17
              errSave = Err.Number
            End If

            If errSave <> 0 Or Not fso.FileExists("{{pdfEsc}}") Then
              statusFile.WriteLine "saveas_failed err=" & errSave & " " & Err.Description
              Err.Clear
              d.ExportAsFixedFormat "{{pdfEsc}}", 17, False, 0, 0
              errExport = Err.Number
              If errExport <> 0 Or Not fso.FileExists("{{pdfEsc}}") Then
                statusFile.WriteLine "export_failed err=" & errExport & " " & Err.Description
                Call SafeClose(d)
                Call ForceQuit(w)
                statusFile.Close
                WScript.Quit 4
              End If
              statusFile.WriteLine "exported_ok"
            Else
              statusFile.WriteLine "saveas_ok"
            End If

            Call SafeClose(d)
            Call ForceQuit(w)
            statusFile.WriteLine "done"
            statusFile.Close
            WScript.Quit 0

            Sub SafeClose(doc)
              On Error Resume Next
              If Not doc Is Nothing Then
                doc.Saved = True
                doc.Close 0
              End If
            End Sub

            Sub ForceQuit(app)
              On Error Resume Next
              Dim j, doc
              If app Is Nothing Then Exit Sub
              For j = app.Documents.Count To 1 Step -1
                Set doc = app.Documents(j)
                doc.Saved = True
                doc.Close 0
              Next
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
        var sysWow = Path.Combine(windir, "SysWOW64", "cscript.exe");
        var sys32 = Path.Combine(windir, "System32", "cscript.exe");

        // Office 2016 is commonly 32-bit — 64-bit cscript cannot CreateObject(Word).
        if (File.Exists(sysWow))
            yield return sysWow;
        yield return sys32;
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
