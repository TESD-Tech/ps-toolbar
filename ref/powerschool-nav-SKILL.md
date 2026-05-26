---
name: powerschool-nav
description: Guidance for adding navigation links to PowerSchool plugin portals. Covers the Enhanced Navigation pagecataloging JSON system (admin portal) and wildcard-based nav injection (teachers and guardian portals). This skill should be used when adding a plugin page to PowerSchool's navigation menus, configuring pagecataloging JSON, or injecting nav links via PowerSchool wildcards.
---

# PowerSchool Navigation

Adding nav links to PowerSchool plugins is non-obvious because each portal uses a completely different mechanism. Do not guess — the patterns are strict and the tooling is specific.

## Portal → Mechanism Map

| Portal | Mechanism | Requires PS version |
|--------|-----------|---------------------|
| Admin | Pagecataloging JSON (`pagecataloging/` folder) | 23.5.0.0+ |
| Teachers | Wildcard footer file + HTML page changes | Any |
| Guardian | Wildcard footer file + HTML page changes | Any |

Read the appropriate reference file before implementing.

## Admin Portal — Pagecataloging JSON

> Full property reference: `references/ps-enhanced-navigation.md`

**File location:** `src/powerschool/pagecataloging/{plugin-name}.json`

`ps-package` picks up the `pagecataloging/` folder and includes it in the **main UI ZIP** (not the DATA zip). No additional build config needed.

**Minimal example — adding a link to the student detail nav:**

```json
{
  "pages": [
    {
      "htmlID": "navMyPluginReport",
      "title": "My Plugin",
      "version": "26.05.01",
      "contextType": "student",
      "requiredContext": "student",
      "sortOrder": 50,
      "pageURL": "/admin/my-plugin/dashboard.html",
      "parentHTMLID": "navStudentAcademicPerformanceSection",
      "districtLevelContext": 1
    }
  ]
}
```

**Common gotchas:**

- `htmlID` must be globally unique across all plugins and PS core. Once set, it cannot be changed.
- `version` must be incremented any time the record changes, or PS will not update the DB record.
- `districtLevelContext`: `0` = both school and district, `1` = school only, `2` = district only. Default is `0`. ELD-style reports that are school-specific should use `1`.
- `pageURL` must start with `/admin`. PS-HTML tags like `~(studentfrn)` are supported in URLs.
- To find `parentHTMLID` values for existing categories: right-click the nav item in PS admin → Inspect → look for `data-custom-info` attribute.

## Teachers Portal — Wildcard Injection

> Full patterns and DOM targets: `references/wildcard-injection.md`

Two things are required:

**1. Create the wildcard footer file:**

`src/powerschool/WEB_ROOT/wildcards/teachers_nav_css.{plugin-slug}.leftnav.footer.txt`

This file auto-appends to `~[wc:teachers_nav_css]` output on every teacher page that calls that wildcard.

**2. Add the nav wildcards to each teacher HTML page:**

```html
~[wc:teachers_header_css]
~[wc:teachers_navigation_css]
~[wc:teachers_nav_css]
```

`teachers_navigation_css` renders the sidebar panel structure. `teachers_nav_css` is the hook the wildcard file injects into. Teacher pages may only have `teachers_header_css` by default — both additions are required.

**DOM injection target:**
```javascript
const ul = document.querySelector('#nav-main ul');
ul.prepend(li);
```

## Guardian Portal — Wildcard Injection

> Full patterns and DOM targets: `references/wildcard-injection.md`

Two things are required:

**1. Create the wildcard footer file:**

`src/powerschool/WEB_ROOT/wildcards/guardian_header_css.{plugin-slug}.leftnav.footer.txt`

Guardian pages already call `~[wc:guardian_header_css]`, so this wildcard fires automatically — no change needed to guardian HTML pages for the script injection.

**2. Add the nav panel wildcard to each guardian HTML page:**

```html
~[wc:guardian_header_css]
~[wc:guardian_navigation_css]
```

`guardian_navigation_css` renders the nav panel. It may be missing from guardian pages by default.

**DOM injection target:**
```javascript
const menu = document.querySelector('[role="menu"]');
menu.prepend(li);
```

## Elementary School Conditional (TESD)

To restrict a nav link to elementary schools only (school ID > 4000 in TESD):

```
~[if#isElem.~(curschoolid)>4000]
  ... nav injection script ...
[/if#isElem]
```

Remove or adjust this conditional if the link should appear for all schools.

## References

- `references/ps-enhanced-navigation.md` — Full pagecataloging JSON property reference, supported PS-HTML tags, known `parentHTMLID` values, version requirements
- `references/wildcard-injection.md` — Wildcard naming conventions, teacher and guardian DOM targets, complete file examples
