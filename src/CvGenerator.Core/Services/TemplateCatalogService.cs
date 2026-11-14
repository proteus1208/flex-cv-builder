namespace CvGenerator.Core.Services;

public sealed class TemplateInfo
{
    public required string Id { get; init; }
    public required string Name { get; init; }
    public required string FullPath { get; init; }
}

public sealed class TemplateCatalogService
{
    public IReadOnlyList<TemplateInfo> ListTemplates(string templatesRoot)
    {
        if (!Directory.Exists(templatesRoot))
            return Array.Empty<TemplateInfo>();

        return Directory.EnumerateFiles(templatesRoot, "*.docx", SearchOption.AllDirectories)
            .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal))
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .Select(f => new TemplateInfo
            {
                Id = Path.GetRelativePath(templatesRoot, f).Replace('\\', '/'),
                Name = Path.GetFileNameWithoutExtension(f),
                FullPath = f
            })
            .ToList();
    }

    public string? ResolvePath(string templatesRoot, string? templateId)
    {
        if (string.IsNullOrWhiteSpace(templateId))
            return null;

        if (Path.IsPathRooted(templateId) && File.Exists(templateId))
            return templateId;

        var combined = Path.Combine(templatesRoot, templateId);
        return File.Exists(combined) ? combined : null;
    }

    public string ImportTemplate(string templatesRoot, string sourceFilePath)
    {
        Directory.CreateDirectory(templatesRoot);
        var name = Path.GetFileName(sourceFilePath);
        var dest = Path.Combine(templatesRoot, name);
        if (File.Exists(dest))
        {
            var stem = Path.GetFileNameWithoutExtension(name);
            var ext = Path.GetExtension(name);
            dest = Path.Combine(templatesRoot, $"{stem}-{DateTime.Now:yyyyMMddHHmmss}{ext}");
        }

        File.Copy(sourceFilePath, dest, overwrite: false);
        return dest;
    }
}
