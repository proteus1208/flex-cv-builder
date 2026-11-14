using CvGenerator.Core.Models;
using CvGenerator.Docx.Services;

namespace CvGenerator.Docx.Tests;

public class TemplateSyntaxValidatorTests
{
    private readonly TemplateSyntaxValidator _sut = new();

    [Fact]
    public void Detects_unpaired_ArrayStart_and_ArrayEnd()
    {
        var issues = _sut.ValidateText("${ArrayStart experience} hi ${name}");
        Assert.Contains(issues, i => i.Code == "UnpairedArrayStart" && i.Severity == IssueSeverity.Error);

        issues = _sut.ValidateText("${ArrayEnd experience}");
        Assert.Contains(issues, i => i.Code == "UnpairedArrayEnd");
    }

    [Fact]
    public void Detects_legacy_start_keyword()
    {
        var issues = _sut.ValidateText("${ start experience}");
        Assert.Contains(issues, i => i.Code == "InvalidPlaceholderKeyword");
    }

    [Fact]
    public void Accepts_valid_placeholders()
    {
        var issues = _sut.ValidateText("""
            ${name}
            ${ArrayStart experience}
            ${experience_i.title}
            ${ArrayEnd experience}
            """);
        Assert.DoesNotContain(issues, i => i.Severity == IssueSeverity.Error);
    }

    [Fact]
    public void Accepts_parent_qualified_nested_loops()
    {
        var issues = _sut.ValidateText("""
            ${ArrayStart skills}
            ${skills_i.category}
            ${ArrayStart skills_i.detail}
            ${skills_i.detail_i.name}
            ${skills_i.detail_i}
            ${ArrayEnd skills_i.detail}
            ${ArrayEnd skills}
            """);
        Assert.DoesNotContain(issues, i => i.Severity == IssueSeverity.Error);
        Assert.DoesNotContain(issues, i => i.Code == "OrphanLoopField");
    }
}
