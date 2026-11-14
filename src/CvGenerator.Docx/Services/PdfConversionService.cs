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

        // Prefer Microsoft Word on Windows — it keeps DOCX fonts/styles most faithfully.
        // Tests can set CVGEN_SKIP_WORD_PDF=1 to force the LibreOffice path.
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) &&
            !IsWordPdfSkipped() &&
            TryConvertWithWord(docxPath, resumePdf, issues))
        {
            return resumePdf;
        }

        var exe = FindLibreOffice(libreOfficePath);
        if (exe is null)
        {
            issues.Error(
                "LibreOfficeMissing",
                "Pdf",
                "LibreOffice (soffice.exe) not found. Install from https://www.libreoffice.org/download/ or click “Set LibreOffice” and pick …\\LibreOffice\\program\\soffice.exe");
            return string.Empty;
        }

        try
        {
            // PDF/A-2b embeds used fonts so the PDF does not silently remap faces.
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

            issues.Info(
                "PdfViaLibreOffice",
                "Pdf",
                "PDF exported via LibreOffice (PDF/A). For closest font match to Word, install the template’s fonts on this PC or install Microsoft Word.");

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
    /// Uses late-bound Word COM when installed. Best font/style fidelity on Windows.
    /// </summary>
    [SupportedOSPlatform("windows")]
    private static bool TryConvertWithWord(string docxPath, string pdfPath, IssueCollector issues)
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

            // Suppress UI prompts that can leave Word hanging / files locked.
            TrySet(wordApp, "Visible", false);
            TrySet(wordApp, "DisplayAlerts", 0); // wdAlertsNone
            TrySet(wordApp, "ScreenUpdating", false);

            documents = wordApp.GetType().InvokeMember(
                "Documents",
                System.Reflection.BindingFlags.GetProperty,
                null,
                wordApp,
                null);

            if (documents is null)
                return false;

            // Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, ...)
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

            // WdExportFormat.wdExportFormatPDF = 17
            doc.GetType().InvokeMember(
                "ExportAsFixedFormat",
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                doc,
                [pdfPath, 17]);

            if (!File.Exists(pdfPath))
                return false;

            issues.Info("PdfViaWord", "Pdf", "PDF exported via Microsoft Word (best font match to the DOCX).");
            return true;
        }
        catch (Exception ex)
        {
            issues.Info(
                "WordPdfUnavailable",
                "Pdf",
                "Microsoft Word PDF export unavailable — falling back to LibreOffice.",
                detail: ex.Message);
            try
            {
                if (File.Exists(pdfPath))
                    File.Delete(pdfPath);
            }
            catch
            {
                // ignore
            }

            return false;
        }
        finally
        {
            // wdDoNotSaveChanges = 0 — must close before Quit or resume.docx stays locked.
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
