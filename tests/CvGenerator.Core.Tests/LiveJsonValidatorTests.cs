using CvGenerator.Core.Models;
using CvGenerator.Core.Services;

namespace CvGenerator.Core.Tests;

public class LiveJsonValidatorTests
{
    [Fact]
    public void Warns_when_json_key_does_not_match_template()
    {
        var sut = new LiveJsonValidator();
        var result = sut.Validate(
            """{"nam":"Ada","role":"Eng"}""",
            new (string Kind, string? JsonPath, string Token)[]
            {
                ("simple", "name", "${name}"),
                ("simple", "role", "${role}"),
                ("array-start", "experience", "${ArrayStart experience}"),
                ("loop-field", "experience_i.title", "${experience_i.title}")
            },
            Array.Empty<GenerationIssue>());

        Assert.Contains(result.Issues, i => i.Path == "name" && i.Code == "MissingField");
        Assert.Contains(result.Issues, i => i.Path == "experience" && i.Code == "MissingLoopArray");
        // Highlight must mirror source JSON from the first paint (no <br/> doubling).
        Assert.DoesNotContain("<br", result.HighlightHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("&quot;nam&quot;", result.HighlightHtml);
    }

    [Fact]
    public void Does_not_paint_unrelated_keys_when_array_field_missing()
    {
        var sut = new LiveJsonValidator();
        var result = sut.Validate(
            """
            {
              "name": "Ada",
              "title": "RootTitle",
              "experience": [ { "company": "A" } ]
            }
            """,
            new (string Kind, string? JsonPath, string Token)[]
            {
                ("loop-field", "experience_i.title", "${experience_i.title}")
            },
            Array.Empty<GenerationIssue>());

        Assert.Contains(result.Issues, i => i.Code == "MissingLoopField");
        // Root "title" should not be warn-marked just because array item lacks title.
        Assert.DoesNotContain("json-mark warn\">&quot;title&quot;", result.HighlightHtml);
    }

    [Fact]
    public void Validates_parent_qualified_nested_ArrayStart()
    {
        var sut = new LiveJsonValidator();
        var result = sut.Validate(
            """
            {
              "name": "Ada",
              "role": "Eng",
              "topSkills": ["C#"],
              "skills": [
                { "category": "Backend", "detail": [ { "name": "C#" } ] }
              ]
            }
            """,
            new (string Kind, string? JsonPath, string Token)[]
            {
                ("array-start", "skills", "${ArrayStart skills}"),
                ("array-start", "skills_i.detail", "${ArrayStart skills_i.detail}"),
                ("loop-field", "skills_i.detail_i.name", "${skills_i.detail_i.name}")
            },
            Array.Empty<GenerationIssue>());

        Assert.DoesNotContain(result.Issues, i => i.Code is "MissingLoopArray" or "MissingLoopField");
    }

    [Fact]
    public void Warns_when_parent_qualified_nested_array_missing()
    {
        var sut = new LiveJsonValidator();
        var result = sut.Validate(
            """
            {
              "name": "Ada",
              "role": "Eng",
              "topSkills": ["C#"],
              "skills": [ { "category": "Backend" } ]
            }
            """,
            new (string Kind, string? JsonPath, string Token)[]
            {
                ("array-start", "skills_i.detail", "${ArrayStart skills_i.detail}")
            },
            Array.Empty<GenerationIssue>());

        Assert.Contains(result.Issues, i => i.Code == "MissingLoopArray");
    }
}
