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

        // Standard DOCX→PDF publisher is Microsoft Word (preserves bars/shapes/fonts).
        // LibreOffice often drops DrawingML accent bars — only use it if Word is unavailable
        // or CVGEN_ALLOW_LO_PDF=1 after a Word failure.
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) &&
            !IsEnvTrue("CVGEN_SKIP_WORD_PDF"))
        {
            if (TryConvertWithWord(docxPath, resumePdf, issues))
                return resumePdf;

            var wordInstalled = FindWinWordExe() != null;
            if (wordInstalled && !IsEnvTrue("CVGEN_ALLOW_LO_PDF"))
            {
                issues.Error(
                    "WordPdfFailed",
                    "Pdf",
                    "Word is installed but Save-As-PDF failed, so PDF was not published. " +
                    "LibreOffice would drop bars/shapes. Close all Word windows/dialogs, " +
                    "then Generate All again. Or open resume.docx in Word → File → Export → Create PDF/XPS. " +
                    "Set CVGEN_ALLOW_LO_PDF=1 only if you accept LibreOffice layout differences.",
                    detail: issues.Issues.LastOrDefault(i => i.Code == "WordPdfUnavailable")?.Detail);
                return string.Empty;
            }
        }

        return ConvertWithLibreOffice(docxPath, outputDirectory, resumePdf, libreOfficePath, issues);
    }

    private string ConvertWithLibreOffice(
        string docxPath,
        string outputDirectory,
        string resumePdf,
        string? libreOfficePath,
        IssueCollector issues)
    {
        var exe = FindLibreOffice(libreOfficePath);
        if (exe is null)
        {
            issues.Error(
                "LibreOfficeMissing",
                "Pdf",
                "No PDF publisher available (Word failed or skipped; LibreOffice soffice.exe not found).");
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
                "PDF via LibreOffice. Accent bars/shapes from Word often disappear. Prefer Word Save-As-PDF.");

            return resumePdf;
        }
        catch (Exception ex)
        {
            issues.Error("LibreOfficeFailed", "Pdf", $"PDF conversion error: {ex.Message}", detail: ex.ToString());
            return string.Empty;
        }
    }

    private static bool IsEnvTrue(string name)
        => string.Equals(Environment.GetEnvironmentVariable(name), "1", StringComparison.Ordinal);

    /// <summary>
    /// Word's built-in PDF publisher on a temp copy (never locks resume.docx).
    /// Uses 32-bit cscript/PowerShell first — Office 2016 is usually 32-bit; this app is win-x64.
    /// </summary>
    [SupportedOSPlatform("windows")]
    private static bool TryConvertWithWord(string docxPath, string pdfPath, IssueCollector issues)
    {
        var workDir = Path.Combine(Path.GetTempPath(), "CvGenerator-word-" + Guid.NewGuid().ToString("N"));
        var scriptPath = Path.Combine(workDir, "export.vbs");
        var psPath = Path.Combine(workDir, "export.ps1");
        var statusPath = Path.Combine(workDir, "status.txt");
        var tempDocx = Path.Combine(workDir, "source.docx");
        var tempPdf = Path.Combine(workDir, "out.pdf");
        var beforePids = SnapshotWordPids();
        var attemptLog = new StringBuilder();
        var winWord = FindWinWordExe();
        attemptLog.AppendLine($"winword={winWord ?? "(not found)"}");
        attemptLog.AppendLine($"process={RuntimeInformation.ProcessArchitecture}");

        try
        {
            Directory.CreateDirectory(workDir);
            // Export the DOCX as-is (temp copy only to avoid locking resume.docx).
            // Do not rewrite borders — that changed bar style/gaps vs the good DOCX.
            File.Copy(docxPath, tempDocx, overwrite: true);

            TryDelete(pdfPath);
            WriteWordExportScript(scriptPath, tempDocx, tempPdf, statusPath);
            WriteWordExportPowerShell(psPath, tempDocx, tempPdf, statusPath);

            // 1) 32-bit cscript (Office 2016 x86)
            // 2) 64-bit cscript (64-bit Office)
            // 3) 32-bit PowerShell COM
            // 4) 64-bit PowerShell COM
            foreach (var host in EnumerateWordHosts())
            {
                KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());
                TryDelete(tempPdf);
                TryDelete(statusPath);

                var psi = CreateHostStartInfo(host, scriptPath, psPath, workDir);
                using var process = Process.Start(psi);
                if (process is null)
                {
                    attemptLog.AppendLine($"{host.Label}: failed to start");
                    continue;
                }

                if (!process.WaitForExit(180_000))
                {
                    try { process.Kill(entireProcessTree: true); } catch { /* ignore */ }
                    KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());
                    attemptLog.AppendLine($"{host.Label}: timed out after 180s");
                    continue;
                }

                var stderrText = "";
                try { stderrText = process.StandardError.ReadToEnd(); } catch { /* ignore */ }
                var stdoutText = "";
                try { stdoutText = process.StandardOutput.ReadToEnd(); } catch { /* ignore */ }

                KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());

                var status = File.Exists(statusPath) ? File.ReadAllText(statusPath).Trim() : "(no status)";
                attemptLog.AppendLine($"{host.Label}: exit={process.ExitCode}; {status}");
                if (!string.IsNullOrWhiteSpace(stderrText))
                    attemptLog.AppendLine("stderr: " + stderrText.Trim());
                if (!string.IsNullOrWhiteSpace(stdoutText))
                    attemptLog.AppendLine("stdout: " + stdoutText.Trim());

                if (File.Exists(tempPdf) && new FileInfo(tempPdf).Length > 0)
                {
                    File.Copy(tempPdf, pdfPath, overwrite: true);
                    issues.Info(
                        "PdfViaWord",
                        "Pdf",
                        "PDF published via Microsoft Word (preserves bars, shapes, and fonts).");
                    return true;
                }
            }

            issues.Warning(
                "WordPdfUnavailable",
                "Pdf",
                "Word Save-As-PDF failed.",
                detail: attemptLog.ToString().Trim());
            return false;
        }
        catch (Exception ex)
        {
            KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());
            issues.Warning(
                "WordPdfUnavailable",
                "Pdf",
                "Word Save-As-PDF failed.",
                detail: attemptLog + Environment.NewLine + ex);
            return false;
        }
        finally
        {
            KillWordPids(SnapshotWordPids().Except(beforePids).ToHashSet());
            TryDeleteDirectory(workDir);
        }
    }

    private readonly record struct WordHost(string Label, string FileName, string Kind);

    private static IEnumerable<WordHost> EnumerateWordHosts()
    {
        var windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        var wowCscript = Path.Combine(windir, "SysWOW64", "cscript.exe");
        var sysCscript = Path.Combine(windir, "System32", "cscript.exe");
        var wowPs = Path.Combine(windir, "SysWOW64", "WindowsPowerShell", "v1.0", "powershell.exe");
        var sysPs = Path.Combine(windir, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");

        if (File.Exists(wowCscript))
            yield return new WordHost("cscript-x86", wowCscript, "vbs");
        yield return new WordHost("cscript-x64", sysCscript, "vbs");
        if (File.Exists(wowPs))
            yield return new WordHost("powershell-x86", wowPs, "ps1");
        yield return new WordHost("powershell-x64", sysPs, "ps1");
    }

    private static ProcessStartInfo CreateHostStartInfo(WordHost host, string vbs, string ps1, string workDir)
    {
        var psi = new ProcessStartInfo
        {
            FileName = host.FileName,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            WorkingDirectory = workDir
        };

        if (host.Kind == "vbs")
        {
            psi.ArgumentList.Add("//nologo");
            psi.ArgumentList.Add(vbs);
        }
        else
        {
            psi.ArgumentList.Add("-NoProfile");
            psi.ArgumentList.Add("-ExecutionPolicy");
            psi.ArgumentList.Add("Bypass");
            psi.ArgumentList.Add("-File");
            psi.ArgumentList.Add(ps1);
        }

        return psi;
    }

    [SupportedOSPlatform("windows")]
    private static void WriteWordExportScript(string scriptPath, string docxPath, string pdfPath, string statusPath)
    {
        var docxEsc = EscapeVbs(docxPath);
        var pdfEsc = EscapeVbs(pdfPath);
        var statusEsc = EscapeVbs(statusPath);

        // ReadOnly must be False — SaveAs/Export can fail on read-only docs.
        // ExportAsFixedFormat first (same as File → Export → Create PDF/XPS).
        var script = $$"""
            Option Explicit
            On Error Resume Next

            Dim fso, statusFile, w, d, n, errN
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set statusFile = fso.CreateTextFile("{{statusEsc}}", True)
            statusFile.WriteLine "start"

            Set w = CreateObject("Word.Application")
            errN = Err.Number
            If errN <> 0 Or w Is Nothing Then
              statusFile.WriteLine "create_failed err=" & errN & " " & Err.Description
              statusFile.Close
              WScript.Quit 2
            End If
            statusFile.WriteLine "word_created version=" & CStr(w.Version)

            w.Visible = False
            w.DisplayAlerts = 0
            w.ScreenUpdating = False
            Err.Clear
            w.AutomationSecurity = 3
            Err.Clear

            ' ConfirmConversions=False, ReadOnly=False, AddToRecentFiles=False
            Set d = w.Documents.Open("{{docxEsc}}", False, False, False)
            errN = Err.Number
            If errN <> 0 Or d Is Nothing Then
              statusFile.WriteLine "open_failed err=" & errN & " " & Err.Description
              Call ForceQuit(w)
              statusFile.Close
              WScript.Quit 3
            End If
            statusFile.WriteLine "opened"

            If fso.FileExists("{{pdfEsc}}") Then fso.DeleteFile "{{pdfEsc}}", True
            Err.Clear
            d.ExportAsFixedFormat "{{pdfEsc}}", 17, False, 0, 0, , , 0, True, True, 0, True, False, False
            errN = Err.Number
            If errN <> 0 Or Not fso.FileExists("{{pdfEsc}}") Then
              statusFile.WriteLine "export_failed err=" & errN & " " & Err.Description
              Err.Clear
              d.SaveAs2 "{{pdfEsc}}", 17
              errN = Err.Number
              If errN <> 0 Or Not fso.FileExists("{{pdfEsc}}") Then
                statusFile.WriteLine "saveas2_failed err=" & errN & " " & Err.Description
                Err.Clear
                d.SaveAs "{{pdfEsc}}", 17
                errN = Err.Number
                If errN <> 0 Or Not fso.FileExists("{{pdfEsc}}") Then
                  statusFile.WriteLine "saveas_failed err=" & errN & " " & Err.Description
                  Call SafeClose(d)
                  Call ForceQuit(w)
                  statusFile.Close
                  WScript.Quit 4
                End If
                statusFile.WriteLine "saveas_ok"
              Else
                statusFile.WriteLine "saveas2_ok"
              End If
            Else
              statusFile.WriteLine "export_ok size=" & fso.GetFile("{{pdfEsc}}").Size
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
    private static void WriteWordExportPowerShell(string psPath, string docxPath, string pdfPath, string statusPath)
    {
        // PowerShell COM — same Word PDF publisher; useful when cscript host bitness is wrong.
        var script = $$"""
            $ErrorActionPreference = 'Continue'
            $status = '{{statusPath.Replace("'", "''")}}'
            $docx = '{{docxPath.Replace("'", "''")}}'
            $pdf = '{{pdfPath.Replace("'", "''")}}'
            function Write-Status([string]$line) {
              Add-Content -LiteralPath $status -Value $line -Encoding ASCII
            }
            Set-Content -LiteralPath $status -Value 'start' -Encoding ASCII
            $w = $null
            $d = $null
            try {
              $w = New-Object -ComObject Word.Application
              Write-Status ("word_created version=" + $w.Version)
              $w.Visible = $false
              $w.DisplayAlerts = 0
              try { $w.ScreenUpdating = $false } catch {}
              try { $w.AutomationSecurity = 3 } catch {}
              $d = $w.Documents.Open($docx, $false, $false, $false)
              Write-Status 'opened'
              if (Test-Path -LiteralPath $pdf) { Remove-Item -LiteralPath $pdf -Force }
              $d.ExportAsFixedFormat($pdf, 17, $false, 0, 0, $null, $null, 0, $true, $true, 0, $true, $false, $false)
              if (-not (Test-Path -LiteralPath $pdf)) {
                Write-Status 'export_failed'
                $wdFormatPDF = 17
                $d.SaveAs([ref]$pdf, [ref]$wdFormatPDF)
              }
              if (-not (Test-Path -LiteralPath $pdf)) {
                Write-Status 'saveas_failed'
                exit 4
              }
              Write-Status ('export_ok size=' + (Get-Item -LiteralPath $pdf).Length)
              exit 0
            }
            catch {
              Write-Status ('exception ' + $_.Exception.Message)
              exit 5
            }
            finally {
              if ($d -ne $null) {
                try { $d.Saved = $true } catch {}
                try { $d.Close([ref]0) } catch {}
                try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($d) } catch {}
              }
              if ($w -ne $null) {
                try { $w.DisplayAlerts = 0 } catch {}
                try { $w.Quit([ref]0) } catch {}
                try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($w) } catch {}
              }
              [GC]::Collect()
              [GC]::WaitForPendingFinalizers()
            }
            """;
        File.WriteAllText(psPath, script, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    [SupportedOSPlatform("windows")]
    private static string? FindWinWordExe()
    {
        string?[] registryPaths =
        [
            @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Word\InstallRoot",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Word\InstallRoot",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Word\InstallRoot",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Word\InstallRoot",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\14.0\Word\InstallRoot"
        ];

        foreach (var key in registryPaths)
        {
            try
            {
                var path = Microsoft.Win32.Registry.GetValue(key!, "Path", null) as string;
                if (string.IsNullOrWhiteSpace(path))
                    continue;
                var exe = Path.Combine(path, "WINWORD.EXE");
                if (File.Exists(exe))
                    return exe;
            }
            catch
            {
                // ignore
            }
        }

        string[] guesses =
        [
            @"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
            @"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
            @"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
            @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
        ];
        return guesses.FirstOrDefault(File.Exists);
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
