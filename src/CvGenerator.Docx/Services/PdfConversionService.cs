using System.Diagnostics;
using System.Runtime.InteropServices;
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

                // Versioned installs: LibreOffice 24, LibreOffice 7.6, etc.
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

        // User may pick the LibreOffice folder or program folder instead of soffice.exe
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
            Directory.CreateDirectory(outputDirectory);

            // PDF/A-2b + embed standard fonts so glyphs used in the DOCX are embedded
            // instead of being remapped to a different face in the PDF viewer.
            // (plain `--convert-to pdf` often substitutes fonts and looks unlike the DOCX.)
            const string pdfFilter =
                "pdf:writer_pdf_Export:" +
                "{" +
                "\"SelectPdfVersion\":{\"type\":\"long\",\"value\":\"2\"}," +
                "\"EmbedStandardFonts\":{\"type\":\"boolean\",\"value\":\"true\"}" +
                "}";

            var psi = new ProcessStartInfo
            {
                FileName = exe,
                ArgumentList =
                {
                    "--headless",
                    "--nologo",
                    "--nofirststartwizard",
                    "--norestore",
                    "--convert-to",
                    pdfFilter,
                    "--outdir",
                    outputDirectory,
                    docxPath
                },
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            // Isolated profile avoids stale Writer PDF-export defaults from a prior GUI session.
            var profileDir = Path.Combine(Path.GetTempPath(), "CvGenerator-lo-profile");
            Directory.CreateDirectory(profileDir);
            var profileUri = new Uri(profileDir.EndsWith(Path.DirectorySeparatorChar)
                ? profileDir
                : profileDir + Path.DirectorySeparatorChar).AbsoluteUri;
            psi.ArgumentList.Insert(0, $"-env:UserInstallation={profileUri}");

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

            var resumePdf = Path.Combine(outputDirectory, "resume.pdf");
            if (!string.Equals(expected, resumePdf, StringComparison.OrdinalIgnoreCase))
            {
                if (File.Exists(resumePdf))
                    File.Delete(resumePdf);
                File.Move(expected, resumePdf);
                return resumePdf;
            }

            return expected;
        }
        catch (Exception ex)
        {
            issues.Error("LibreOfficeFailed", "Pdf", $"PDF conversion error: {ex.Message}", detail: ex.ToString());
            return string.Empty;
        }
    }
}
