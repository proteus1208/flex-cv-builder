using CvGenerator.Core.Models;
using CvGenerator.Core.Services;

namespace CvGenerator.Core.Tests;

public class LiveJsonValidatorTests
{
    [Fact]
    public void Warns_on_missing_and_empty_fields()
    {
        var sut = new LiveJsonValidator();
        var result = sut.Validate(
            """{"name":"Ada","role":""}""",
            new[]
            {
                ("simple", "name", "${name}"),
                ("simple", "role", "${role}"),
                ("simple", "linkedin", "${linkedin}"),
                ("loop-start", "experience", "${ start experience}")
            },
            Array.Empty<GenerationIssue>());

        Assert.Contains(result.Issues, i => i.Path == "role" && i.Code == "EmptyField");
        Assert.Contains(result.Issues, i => i.Path == "linkedin" && i.Code == "MissingField");
        Assert.Contains(result.Issues, i => i.Path == "experience" && i.Code == "MissingLoopArray");
        Assert.Contains("json-mark warn", result.HighlightHtml);
    }
}
