using CvGenerator.Core.Models;
using CvGenerator.Core.Services;

namespace CvGenerator.Core.Tests;

public class LiveJsonValidatorTests
{
    [Fact]
    public void Accepts_absolute_dotted_array_paths_and_alias_item_fields()
    {
        var sut = new LiveJsonValidator();
        var result = sut.Validate(
            """
            {
              "experience": {
                "company_1": {
                  "industries": ["Education", "FinTech"],
                  "skills": { "primary": ["C#", ".NET"] }
                }
              },
              "industries": ["Healthcare"]
            }
            """,
            new (string Kind, string? JsonPath, string Token)[]
            {
                ("array-start", "experience.company_1.industries", "${ArrayStart experience.company_1.industries industry}"),
                ("loop-field", "experience.company_1.industries_i", "${industry}"),
                ("array-start", "experience.company_1.skills.primary", "${ArrayStart experience.company_1.skills.primary topskill}"),
                ("loop-field", "experience.company_1.skills.primary_i", "${topskill}"),
                ("array-start", "industries", "${ArrayStart industries industry}"),
                ("loop-field", "industries_i", "${industry}")
            },
            Array.Empty<GenerationIssue>());

        Assert.DoesNotContain(result.Issues, i => i.Code is "MissingLoopArray" or "MissingField" or "MissingLoopField");
    }

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

    [Fact]
    public void Does_not_flag_canonical_alias_loop_fields_as_missing_root_keys()
    {
        var sut = new LiveJsonValidator();
        var result = sut.Validate(
            """
            {
              "name": "Ada",
              "role": "Eng",
              "topSkills": ["C#"],
              "experience": [ { "title": "Lead", "company": "A" } ]
            }
            """,
            new (string Kind, string? JsonPath, string Token)[]
            {
                ("array-start", "experience", "${ArrayStart experience exp}"),
                ("loop-field", "experience_i.title", "${exp.title}"),
                ("loop-field", "experience_i.company", "${exp.company}")
            },
            Array.Empty<GenerationIssue>());

        Assert.DoesNotContain(result.Issues, i => i.Code == "MissingField");
        Assert.DoesNotContain(result.Issues, i => i.Code is "MissingLoopArray" or "MissingLoopField");
    }
}
