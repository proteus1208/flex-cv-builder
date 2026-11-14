using CvGenerator.Core.Models;
using CvGenerator.Docx.Services;

namespace CvGenerator.Docx.Tests;

public class TemplateSyntaxValidatorTests
{
    private readonly TemplateSyntaxValidator _sut = new();

    [Fact]
    public void Detects_unpaired_start_and_end()
    {
        var issues = _sut.ValidateText("${ start experience} hi ${name}");
        Assert.Contains(issues, i => i.Code == "UnpairedLoopStart" && i.Severity == IssueSeverity.Error);

        issues = _sut.ValidateText("${ end experience}");
        Assert.Contains(issues, i => i.Code == "UnpairedLoopEnd");
    }

    [Fact]
    public void Detects_invalid_keyword_and_too_many_parts()
    {
        var issues = _sut.ValidateText("${ starts experience} ${foo bar baz}");
        Assert.Contains(issues, i => i.Code == "InvalidPlaceholderKeyword");
        Assert.Contains(issues, i => i.Code == "InvalidPlaceholder");
    }

    [Fact]
    public void Accepts_valid_placeholders()
    {
        var issues = _sut.ValidateText("""
            ${name}
            ${ start experience}
            ${experience.i.title}
            ${ end experience}
            """);
        Assert.DoesNotContain(issues, i => i.Severity == IssueSeverity.Error);
    }
}
