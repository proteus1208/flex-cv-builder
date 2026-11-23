namespace CvGenerator.App.Services;

/// <summary>Selectable / copyable grammar reference shown in the Help dialog.</summary>
public static class GrammarHelpText
{
    public const string Title = "Template & JSON grammar";

    public const string Body = """
CV Generator grammar
====================

JSON is free-form. Put any structure you need. The DOCX template pulls
values with ${...} placeholders. Markup inside JSON string values can
style text when it is written into the DOCX.


1) Simple fields
----------------
Template:
  ${name}
  ${contact.email}
  ${linkedin}

JSON:
  {
    "name": "Ada Lovelace",
    "contact": { "email": "ada@example.com" },
    "linkedin": "[Ada](https://linkedin.com/in/ada)"
  }

Notes:
  - Paths use dots for nested objects.
  - Missing fields become empty and show a warning in Debug Error List.


2) Arrays (loops)
-----------------
Classic (_i form):
  ${ArrayStart experience}
  ${experience_i.title} — ${experience_i.company}
  ${ArrayEnd experience}

With iterate alias (recommended):
  ${ArrayStart experience exp}
  ${exp.title} — ${exp.company}
  ${ArrayEnd experience}

JSON:
  {
    "experience": [
      { "title": "Engineer", "company": "A" },
      { "title": "Lead", "company": "B" }
    ]
  }

Notes:
  - ${ArrayStart varname iteratevarname} binds each item to iteratevarname.
  - Then use ${iteratevarname.field} instead of ${varname_i.field}.
  - ArrayEnd matches the array source (varname). Alias on End is optional.
  - Any root-level JSON array works, e.g. ${ArrayStart extra_infos info}${info}...
  - Nested under an alias: ${ArrayStart exp.skills.primary topskill}${topskill}...
  - Absolute dotted paths from JSON root are also allowed, e.g.
    ${ArrayStart experience.company_1.skills.primary topskill}${topskill}...
  - Everything between the markers is repeated per item; styles are kept.


2b) Join lists without a trailing comma
---------------------------------------
Wrong (trailing comma):
  Technology Stack: ${ArrayStart exp.skills skill}${skill}, ${ArrayEnd exp.skills}
  → React, Node.js,

Correct:
  Technology Stack: ${ArrayStart exp.skills skill}${skill}${JoinStart}, ${JoinEnd}${ArrayEnd exp.skills}
  → React, Node.js

  ${ArrayStart skills}${skills_i}${JoinStart}, ${JoinEnd}${ArrayEnd skills}

Notes:
  - Put the separator only between ${JoinStart} and ${JoinEnd}.
  - It is inserted between items, never after the last item.


3) Nested arrays
----------------
With aliases:
  ${ArrayStart skills skill}
  ${skill.category}: ${ArrayStart skill.detail detail}${detail}${JoinStart}, ${JoinEnd}${ArrayEnd skill.detail}
  ${ArrayEnd skills}

Classic (_i) still works:
  ${ArrayStart skills}
  ${skills_i.category}
  ${ArrayStart skills_i.detail}
  - ${skills_i.detail_i.name}
  ${ArrayEnd skills_i.detail}
  ${ArrayEnd skills}

JSON:
  {
    "skills": [
      {
        "category": "Backend",
        "detail": [
          { "name": "C#", "level": "expert" }
        ],
        "tags": ["api", "dotnet"]
      }
    ]
  }

Notes:
  - Nested start can use the parent alias: ${ArrayStart skill.detail d}
  - Or classic parent-qualified: ${ArrayStart skills_i.detail}
  - Scalar string items: ${tag} or ${skills_i.tags_i}


4) Rich text inside JSON string values
--------------------------------------
These markers work inside any string that fills a ${placeholder}:

  **bold**
  *italic*
  __underline__
  [label](https://example.com)

Examples:
  "summary": "Shipped **PDF** export with *fast* __preview__"
  "linkedin": "[Cody Lee Berry](https://linkedin.com/in/example)"
  "website": "[Portfolio](https://example.com)"

Rules:
  - Links must be http, https, or mailto.
  - Markup is only expanded when substituting into the DOCX.
  - You can also bold/color the whole ${placeholder} in Word; that style
    is kept as the base, and these markers add on top.


5) Output folder fields
-----------------------
The export folder is named from:

  directory   (optional string)

Folder shape:
  {OutputRoot}/{directory}/
  or, if directory is missing/empty:
  {OutputRoot}/{yyyy-MM-dd_HH-mm-ss}/

Artifacts written there:
  resume.json
  resume.docx
  resume.pdf   (Word Export preferred; LibreOffice fallback)
  notes.txt    (from JSON "notes")


6) Quick cheat sheet
--------------------
  ${path}                         root / dotted field
  ${ArrayStart name}              begin loop (classic)
  ${ArrayStart name alias}        begin loop with iterate alias
  ${ArrayEnd name}                end loop
  ${JoinStart}...${JoinEnd}       separator between items only
  ${name_i.field}                 classic field on loop item
  ${alias.field}                  field via iterate alias
  ${alias}                        scalar item via alias
  ${parent_i.child}               nested ArrayStart (classic)
  ${alias.child}                  nested ArrayStart via parent alias
  **text**  *text*  __text__      bold / italic / underline
  [label](url)                    hyperlink
""";
}
