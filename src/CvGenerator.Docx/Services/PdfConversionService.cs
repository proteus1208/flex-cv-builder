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
            if (TryConvertWithWordCom(docxPath, resumePdf, issues))
                return resumePdf;

            if (TryConvertWithWordScript(docxPath, resumePdf, issues))
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
            // PDF/A-2b embeds the fonts LibreOffice actually used.
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

    /// <summary>Late-bound Word COM. BitmapMissingFonts=false embeds fonts instead of rasterizing.</summary>
    [SupportedOSPlatform("windows")]
    private static bool TryConvertWithWordCom(string docxPath, string pdfPath, IssueCollector issues)
    {
        object? wordApp = null;
        object? documents = null;
        object? doc = null;
        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType is null)
                return false;

            wordApp = Activator.CreateInstance(wordType);
            if (wordApp is null)
                return false;

            TrySet(wordApp, "Visible", false);
            TrySet(wordApp, "DisplayAlerts", 0);
            TrySet(wordApp, "ScreenUpdating", false);

            documents = wordApp.GetType().InvokeMember(
                "Documents",
                System.Reflection.BindingFlags.GetProperty,
                null,
                wordApp,
                null);

            if (documents is null)
                return false;

            doc = documents.GetType().InvokeMember(
                "Open",
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                documents,
                [docxPath, false, true, false]);

            if (doc is null)
                return false;

            if (File.Exists(pdfPath))
                File.Delete(pdfPath);

            // ExportAsFixedFormat(
            //   OutputFileName, ExportFormat=PDF(17), OpenAfterExport=false,
            //   OptimizeFor=Print(0), Range=AllDocument(0), From, To, Item=DocumentContent(0),
            //   IncludeDocProps=true, KeepIRM=true, CreateBookmarks=Heading(1),
            //   DocStructureTags=true, BitmapMissingFonts=false, UseISO19005_1=false)
            var missing = Type.Missing;
            doc.GetType().InvokeMember(
                "ExportAsFixedFormat",
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                doc,
                [
                    pdfPath,
                    17,      // wdExportFormatPDF
                    false,   // OpenAfterExport
                    0,       // wdExportOptimizeForPrint
                    0,       // wdExportAllDocument
                    missing, // From
                    missing, // To
                    0,       // wdExportDocumentContent
                    true,    // IncludeDocProps
                    true,    // KeepIRM
                    1,       // wdExportCreateHeadingBookmarks
                    true,    // DocStructureTags
                    false,   // BitmapMissingFonts — embed fonts, don't draw bitmaps
                    false    // UseISO19005_1
                ]);

            if (!File.Exists(pdfPath))
                return false;

            issues.Info("PdfViaWord", "Pdf", "PDF exported via Microsoft Word (fonts match the DOCX).");
            return true;
        }
        catch (Exception ex)
        {
            issues.Info(
                "WordComUnavailable",
                "Pdf",
                "Word COM PDF export unavailable — trying script fallback / LibreOffice.",
                detail: ex.Message);
            TryDelete(pdfPath);
            return false;
        }
        finally
        {
            TryCom(doc, "Close", [0]);
            TryCom(wordApp, "Quit", [0]);
            ReleaseCom(ref doc);
            ReleaseCom(ref documents);
            ReleaseCom(ref wordApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }

    /// <summary>
    /// cscript CreateObject often finds Click-to-Run Word when in-process COM does not
    /// (32/64-bit mismatch). Tries System32 and SysWOW64 cscript.
    /// </summary>
    [SupportedOSPlatform("windows")]
    private static bool TryConvertWithWordScript(string docxPath, string pdfPath, IssueCollector issues)
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), "cvgen-word-pdf-" + Guid.NewGuid().ToString("N") + ".vbs");
        try
        {
            if (File.Exists(pdfPath))
                File.Delete(pdfPath);

            var docxEsc = EscapeVbs(docxPath);
            var pdfEsc = EscapeVbs(pdfPath);
            var script = $$"""
                On Error Resume Next
                Dim w, d
                Set w = CreateObject("Word.Application")
                If Err.Number <> 0 Or w Is Nothing Then
                  WScript.Quit 2
                End If
                w.Visible = False
                w.DisplayAlerts = 0
                Set d = w.Documents.Open("{{docxEsc}}", False, True, False)
                If Err.Number <> 0 Or d Is Nothing Then
                  w.Quit 0
                  WScript.Quit 3
                End If
                d.ExportAsFixedFormat "{{pdfEsc}}", 17, False, 0, 0, , , 0, True, True, 1, True, False, False
                If Err.Number <> 0 Then
                  d.Close 0
                  w.Quit 0
                  WScript.Quit 4
                End If
                d.Close 0
                w.Quit 0
                WScript.Quit 0
                """;
            File.WriteAllText(scriptPath, script, Encoding.ASCII);

            foreach (var cscript in EnumerateCscriptHosts())
            {
                TryDelete(pdfPath);
                var psi = new ProcessStartInfo
                {
                    FileName = cscript,
                    ArgumentList = { "//nologo", scriptPath },
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using var process = Process.Start(psi);
                if (process is null)
                    continue;

                if (!process.WaitForExit(120_000))
                {
                    try { process.Kill(entireProcessTree: true); } catch { /* ignore */ }
                    continue;
                }

                if (process.ExitCode == 0 && File.Exists(pdfPath))
                {
                    issues.Info("PdfViaWordScript", "Pdf", "PDF exported via Microsoft Word script (fonts match the DOCX).");
                    return true;
                }
            }

            issues.Info("WordScriptUnavailable", "Pdf", "Word script PDF export failed — falling back to LibreOffice.");
            return false;
        }
        catch (Exception ex)
        {
            issues.Info("WordScriptUnavailable", "Pdf", "Word script PDF export failed — falling back to LibreOffice.", detail: ex.Message);
            return false;
        }
        finally
        {
            TryDelete(scriptPath);
        }
    }

    private static IEnumerable<string> EnumerateCscriptHosts()
    {
        var windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        yield return Path.Combine(windir, "System32", "cscript.exe");
        yield return Path.Combine(windir, "SysWOW64", "cscript.exe");
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

    private static void TrySet(object target, string property, object value)
    {
        try
        {
            target.GetType().InvokeMember(
                property,
                System.Reflection.BindingFlags.SetProperty,
                null,
                target,
                [value]);
        }
        catch
        {
            // ignore optional properties
        }
    }

    private static void TryCom(object? target, string method, object?[]? args)
    {
        if (target is null)
            return;
        try
        {
            target.GetType().InvokeMember(
                method,
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                target,
                args);
        }
        catch
        {
            // ignore cleanup failures
        }
    }

    [SupportedOSPlatform("windows")]
    private static void ReleaseCom(ref object? com)
    {
        if (com is null)
            return;
        try { Marshal.FinalReleaseComObject(com); }
        catch { /* ignore */ }
        com = null;
    }
}
