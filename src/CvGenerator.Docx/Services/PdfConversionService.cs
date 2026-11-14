using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using CvGenerator.Core.Services;

namespace CvGenerator.Docx.Services;

public sealed class PdfConversionService
{
    public string? FindLibreOffice(string? configuredPath)
    {
        if (!string.IsNullOrWhiteSpace(configuredPath) && File.Exists(configuredPath))
            return configuredPath;

        var candidates = new List<string>();

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            var pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var pf86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            candidates.Add(Path.Combine(pf, "LibreOffice", "program", "soffice.exe"));
            candidates.Add(Path.Combine(pf86, "LibreOffice", "program", "soffice.exe"));
            candidates.Add(@"C:\Program Files\LibreOffice\program\soffice.exe");
            candidates.Add(@"C:\Program Files (x86)\LibreOffice\program\soffice.exe");
        }
        else
        {
            candidates.Add("/usr/bin/soffice");
            candidates.Add("/usr/bin/libreoffice");
            candidates.Add("/snap/bin/libreoffice");
        }

        foreach (var c in candidates)
        {
            if (File.Exists(c))
                return c;
        }

        // PATH lookup
        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var dir in pathEnv.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            foreach (var name in new[] { "soffice", "soffice.exe", "libreoffice", "libreoffice.exe" })
            {
                var full = Path.Combine(dir, name);
                if (File.Exists(full))
                    return full;
            }
        }

        return null;
    }

    public string ConvertToPdf(string docxPath, string outputDirectory, string? libreOfficePath, IssueCollector issues)
    {
        var exe = FindLibreOffice(libreOfficePath);
        if (exe is null)
        {
            issues.Error(
                "LibreOfficeMissing",
                "Pdf",
                "LibreOffice (soffice) not found. Install LibreOffice or set the path in settings.");
            return string.Empty;
        }

        try
        {
            Directory.CreateDirectory(outputDirectory);
            var psi = new ProcessStartInfo
            {
                FileName = exe,
                ArgumentList =
                {
                    "--headless",
                    "--nologo",
                    "--nofirststartwizard",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    outputDirectory,
                    docxPath
                },
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process is null)
            {
                issues.Error("LibreOfficeStartFailed", "Pdf", "Failed to start LibreOffice process.");
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
                issues.Error("LibreOfficeTimeout", "Pdf", "LibreOffice conversion timed out after 120s.");
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
                    detail: $"{stdout}\n{stderr}");
                return string.Empty;
            }

            // Normalize to resume.pdf if needed
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
