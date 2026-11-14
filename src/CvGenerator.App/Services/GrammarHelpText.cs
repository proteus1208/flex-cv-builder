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
Template:
  ${ArrayStart experience}
  ${experience_i.title} — ${experience_i.company}
  ${ArrayEnd experience}

JSON:
  {
    "experience": [
      { "title": "Engineer", "company": "A" },
      { "title": "Lead", "company": "B" }
    ]
  }

Notes:
  - ArrayStart / ArrayEnd names must match exactly.
  - Inside the loop, use arrayname_i.field (example: experience_i.title).
  - Everything between the markers (paragraphs, tables, images) is repeated
    once per array item. Styles from the template are kept.


3) Nested arrays (preferred: parent-qualified)
----------------------------------------------
Template:
  ${ArrayStart skills}
  ${skills_i.category}
  ${ArrayStart skills_i.detail}
  - ${skills_i.detail_i.name} (${skills_i.detail_i.level})
  ${ArrayEnd skills_i.detail}
  ${ArrayStart skills_i.tags}
  #${skills_i.tags_i}
  ${ArrayEnd skills_i.tags}
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
  - Prefer ${ArrayStart skills_i.detail} so the parent is obvious.
  - Relative names still work: ${ArrayStart detail} inside a skills loop.
  - For string arrays, use ${skills_i.tags_i} with no field after _i.


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
These JSON keys name the export folder:

  name
  role
  topSkills   (array; first 3 used)

Folder shape:
  {OutputRoot}/{yyyy-MM-dd}/{name} - {role} - {skill1}-{skill2}-{skill3}/

Artifacts written there:
  resume.json
  resume.docx
  resume.pdf   (needs LibreOffice)
  notes.txt    (from JSON "notes")


6) Quick cheat sheet
--------------------
  ${path}                         root / dotted field
  ${ArrayStart name}              begin loop
  ${ArrayEnd name}                end loop
  ${name_i.field}                 field on loop item
  ${parent_i.child}               nested ArrayStart name
  ${parent_i.child_i.field}       nested item field
  ${parent_i.child_i}             nested scalar item
  **text**  *text*  __text__      bold / italic / underline
  [label](url)                    hyperlink
""";
}
