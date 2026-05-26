---
name: powerschool-plugin-development
description: Expert guidance for PowerSchool plugin development including folder structure (permissions_root, user_schema_root, queries_root, WEB_ROOT, pagecataloging, MessageKeys), portal organization (admin, teachers, guardian, subs), plugin.xml configuration, ps-package build automation, and deployment. Use when building PowerSchool plugins, organizing plugin structure, or packaging plugins.
---

# PowerSchool Plugin Development

This skill provides expert guidance for PowerSchool plugin development fundamentals including folder structure, portal organization, configuration, and build automation.

## PowerSchool Folder Structure

PowerSchool plugins use a standardized folder structure. Here's what each folder contains and its purpose:

### permissions_root/
**⚠️ WARNING:** `permissions_root/` is NOT a recognized PowerSchool plugin directory and will cause installation errors if included in your plugin package.

Permissions are managed through PowerSchool's admin interface after plugin installation, not through plugin files.

**❌ Do NOT include this directory in your plugin:**
```
src/powerschool/
├── permissions_root/    # ERROR: Not supported!
│   └── permissions.xml  # Will cause "unrecognized file" error
```

Instead, configure permissions in PowerSchool after installation:
1. System > Security > Permission Groups
2. Create custom permission groups
3. Assign to user roles

If your organization uses a custom permissions system, store permission XML files outside the PowerSchool plugin structure for reference only.

### user_schema_root/
Stores database schema definitions for custom tables and fields.
- XML files defining custom database structures
- Used when adding custom data to PowerSchool's database
- Defines custom student fields, tables, or extensions
- **Build destination:** `schema/` (DATA variant only)

**IMPORTANT:** PowerSchool requires the `psExtension` format with PowerSchool namespace. The old `<tables>` format will be REJECTED during installation.

**CORRECT Example structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<psExtension xmlns="http://www.powerschool.com"
             xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
             xsi:schemaLocation="http://www.powerschool.com psextension.xsd">
    <extensionname>U_My_Table</extensionname>
    <extendedTable type="standalone" dbTableName="u_my_table" childName="u_my_table" comment="Description">
        <field name="id" type="Integer"/>
        <field name="studentsdcid" type="Integer"/>
        <field name="name" type="String" length="255"/>
        <field name="created_date" type="Date"/>
    </extendedTable>
</psExtension>
```

**Key patterns:**
- Custom tables must use `u_` prefix (e.g., `u_student_awards`)
- Each table gets its own XML file
- Use `type="standalone"` for independent tables
- Use `coreTable="Students"` for student-extended tables
- Field types: `Integer`, `String` (with length), `Date`, `Boolean`

For complete schema requirements and validation rules, see `references/schema-requirements.md`.

### queries_root/
Contains named query definitions (XML files) that can be reused throughout PowerSchool.
- Pre-defined SQL queries callable by name
- Avoids writing SQL directly in pages
- Reusable across multiple pages
- **Build destination:** `dist/` (standard plugin)

**CRITICAL:** PowerSchool does NOT support subdirectories in `queries_root/`. All query XML files must be at the root level. Subdirectories will cause installation errors.

**✅ CORRECT Structure:**
```
queries_root/
├── my_queries.xml
└── admin_queries.xml
```

**❌ WRONG - Will fail installation:**
```
queries_root/
├── core/
│   └── queries.xml    # ERROR: No subdirectories allowed!
```

**Example query file:**
```xml
<queries xmlns="http://powerschool.com/dtds/query-1.0.dtd">
    <query name="my_plugin.get_data" coreTable="students">
        <args>
            <arg name="studentsdcid" type="primitive"/>
        </args>
        <sql><![CDATA[
            SELECT /* query here */
        ]]></sql>
    </query>
</queries>
```

For complete examples with complex queries, see `references/schema-requirements.md`.

### WEB_ROOT/
The main folder for web assets and user interface.
- HTML pages, JavaScript, CSS, images
- Contains the portal subdirectories (see Portal Structure below)
- Your plugin's visible user interface
- **Build destination:** `dist/` (standard plugin)

**⚠️ PSHTML Processing Requirement:** PowerSchool only processes PSHTML template tags (`~(...)`, `~[...]`) in files served through the correct portal paths. Files must be placed inside the appropriate portal subdirectory (`WEB_ROOT/admin/`, `WEB_ROOT/teachers/`, etc.) — **not** at `WEB_ROOT/` root. Files at the wrong level are not served through the portal and PSHTML tags will not be evaluated. Wildcards must similarly be in `WEB_ROOT/wildcards/`.

### pagecataloging/
Contains JSON files that register your plugin's pages with PowerSchool's navigation.
- Defines where pages appear in menus
- Controls page visibility by user type
- Maps navigation structure
- **Build destination:** `dist/` (standard plugin)

**Example structure:**
```json
{
  "pages": [
    {
      "htmlID": "navMyPluginPage",
      "title": "My Plugin Page",
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

See the **powerschool-nav** skill for the full property reference, supported `parentHTMLID` values, and wildcard-based nav injection for teacher/guardian portals.

### MessageKeys/
Stores internationalization/localization files.
- Text strings that can be translated into different languages
- Makes plugins multi-language capable
- Key-value pairs for UI text
- **Build destination:** `schema/` (DATA variant only)

**Example structure:**
```properties
my_plugin.title=Plugin Title
my_plugin.form.label=Form Label
```

Usage in HTML: `~[text:my_plugin.title]`

## WEB_ROOT Portal Structure

The WEB_ROOT folder is organized into 4 main portal subdirectories, each serving different user types:

### admin/
Administrative portal pages for district and school administrators.
- District-level administration pages
- School-level administration pages
- System configuration interfaces
- Reporting and analytics for admins
- **Typical files:** HTML pages, admin-specific JavaScript, management interfaces

### teachers/
Teacher portal pages for teacher-specific workflows.
- Gradebook interfaces
- Attendance tools
- Student information views
- Teacher-specific reporting
- **Typical files:** HTML pages, teacher workflow scripts, classroom management tools

### guardian/
Parent/guardian portal pages for family access.
- Student progress viewing
- Grade access
- Attendance monitoring
- Communication tools
- **Typical files:** HTML pages, parent-facing interfaces, read-only student data views

### subs/
Substitute teacher portal pages for substitute workflows.
- Substitute-specific views
- Limited classroom access
- Basic student information
- Temporary access tools
- **Typical files:** HTML pages, simplified teacher interfaces, restricted access pages

### Portal Organization Tips
- Keep portal-specific code isolated within each directory
- Shared utilities can be placed in WEB_ROOT root or a common/ subdirectory
- Consider user permissions when deciding which portal to use
- **UI Technology Options:**
  - Traditional HTML pages for simple interfaces
  - AngularJS for legacy plugins and client-side data access (see **powerschool-ui** Part 3)
  - Modern frameworks (Svelte/Vue) for rich, interactive UIs (see **powerschool-ui** Part 2)

## plugin.xml Configuration

The plugin.xml file is your plugin's manifest and configuration file.

### ⚠️ CRITICAL: Element Ordering Matters

PowerSchool validates `plugin.xml` with strict element ordering. You MUST follow this order or installation will fail:

1. `<oauth/>` (optional)
2. `<access_request>` (BEFORE publisher!)
3. `<publisher>`
4. Other elements

**❌ WRONG - This will cause "cvc-complex-type.2.4.a" error:**
```xml
<plugin ...>
    <publisher .../>      <!-- ERROR: publisher before access_request -->
    <access_request .../> <!-- Will fail validation! -->
</plugin>
```

**✅ CORRECT - Element order:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<plugin xmlns="http://plugin.powerschool.pearson.com"
        xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
        xsi:schemaLocation="http://plugin.powerschool.pearson.com plugin.xsd"
        name="My Plugin"
        version="26.01.01"
        description="Plugin description">

    <!-- 1. OAuth (optional) -->
    <oauth/>

    <!-- 2. Access Request (BEFORE publisher) -->
    <access_request>
        <field table="u_my_table" field="*" access="FullAccess"/>
        <field table="students" field="student_number" access="ViewOnly"/>
    </access_request>

    <!-- 3. Publisher (AFTER access_request) -->
    <publisher name="Organization">
        <contact email="support@example.com"/>
    </publisher>
</plugin>
```

### Key Elements
- **name**: Plugin display name
- **version**: CalVer format (YY.MM.PATCH, e.g., 26.01.01)
- **description**: Brief plugin description
- **access_request**: Database permissions (MUST come before publisher)
- **publisher**: Organization info (MUST come after access_request)

For complete examples including OAuth and custom links, see `references/examples.md`.

### CalVer Versioning (YY.MM.PATCH)
- **YY**: Two-digit year (e.g., 25 for 2025)
- **MM**: Two-digit month (e.g., 07 for July)
- **PATCH**: Two-digit patch number (01-99)
- If year or month changes, patch resets to 01
- Otherwise, patch increments

**Examples:**
- `25.01.01` → `25.01.02` (patch increment)
- `25.01.15` → `25.02.01` (month changed, patch reset)
- `25.12.05` → `26.01.01` (year changed, patch reset)

### DATA Variant
The build process creates a second "DATA" variant of plugin.xml:
- Appends " DATA" to plugin name
- Removes access_request elements
- Used for schema-only deployment
- Plugin names > 35 characters are truncated

**Purpose:** Install DATA variant first to create custom tables, then install standard plugin that uses those tables.

## ps-package Build Automation

The ps-package CLI tool automates the build and packaging process.

### Build Command
```bash
node index.js
# or
pnpm run build
```

### Build Process Flow
1. **Version Management**: Reads current version, generates new CalVer version
2. **Directory Setup**: Ensures dist/, plugin_archive/, schema/ exist
3. **Version Updates**: Updates package.json, plugin.xml, pagecataloging files
4. **XML Generation**: Creates standard plugin.xml (dist/) and DATA variant (schema/)
5. **Build Directory Prep**: Merges PowerSchool folders, removes junk files
6. **Framework Handling**: If projectType is 'svelte' or 'vue', copies public/build to dist/WEB_ROOT/{pluginName}
7. **Archive Creation**: Creates two ZIP files:
   - {pluginName}-{version}.zip (from dist/)
   - DATA-{pluginName}-{version}.zip (from schema/)
8. **Archive Pruning**: Keeps only 10 most recent archives

### Expected Project Structure
```
project-root/
├── src/
│   └── powerschool/          # PowerSchool-specific files
│       ├── permissions_root/
│       ├── user_schema_root/
│       ├── queries_root/
│       ├── WEB_ROOT/
│       ├── pagecataloging/
│       └── MessageKeys/
├── plugin.xml                # Plugin manifest (at root)
├── package.json
├── dist/                     # Generated by build
├── schema/                   # Generated by build
└── plugin_archive/           # Generated by build
```

### Build Configuration
Key configuration in package.json:
```json
{
  "psPackage": {
    "pluginName": "my-plugin",
    "sourceDir": "src",
    "powerSchoolSourceDir": "src/powerschool",
    "buildDir": "dist",
    "schemaDir": "schema",
    "archiveDir": "plugin_archive",
    "projectType": "vue"
  }
}
```

**Configuration options:**
- **sourceDir**: src/ (source files)
- **powerSchoolSourceDir**: src/powerschool/ (PS-specific files)
- **buildDir**: dist/ (build output)
- **schemaDir**: schema/ (schema-only plugin)
- **archiveDir**: plugin_archive/ (ZIP archives)
- **projectType**: 'vue' or 'svelte' (affects build steps)
- **junkFiles**: Files to remove (.DS_Store, Thumbs.db, robots.txt, etc.)

## Plugin Installation Workflow

### Step 1: Install DATA Variant (if using custom tables)
1. Go to **System > Plugin Management Dashboard**
2. Upload **DATA-{pluginName}-{version}.zip**
3. Enable plugin
4. This creates custom database tables

### Step 2: Install Standard Plugin
1. Upload **{pluginName}-{version}.zip**
2. Enable plugin
3. Configure permissions in Security settings

**Why this order?** The standard plugin may reference custom tables created by the DATA variant, so those tables must exist first.

## General Troubleshooting

### Build Errors
- **Missing folders**: Ensure all required PowerSchool folders exist in src/powerschool/
- **Version format**: Verify plugin.xml uses CalVer format (YY.MM.PATCH)
- **Archive issues**: Check that dist/ and schema/ directories are properly created
- **Framework build fails**: Verify public/build/ exists if using Svelte/Vue

### Plugin Installation Issues

**Common Schema Errors:**
- **"Badly-formatted XML: unexpected element"**: Using `<tables>` instead of `<psExtension>` → See `references/schema-requirements.md`
- **"Plugin file contains an unrecognized file"**: Subdirectories in queries_root or permissions_root directory → Flatten structure
- **"cvc-complex-type.2.4.a: Invalid content"**: Wrong element order in plugin.xml → Put access_request BEFORE publisher

**Other Issues:**
- **Navigation issues**: Validate pagecataloging/ JSON files
- **Custom table errors**: Ensure table names start with `u_` prefix
- **Field type errors**: Use Integer, String (with length), Date, Boolean in psExtension format

**For complete troubleshooting guide, see `references/schema-requirements.md`**

### General Tips
- Always test builds before deployment
- Keep plugin.xml at project root (not in src/)
- Use psExtension format for ALL schema files
- No subdirectories in queries_root/
- Do NOT include permissions_root/ in plugin package
- Follow PowerSchool naming conventions
- Archive pruning keeps only 10 recent builds by modification time
- Install DATA variant before standard plugin if using custom tables

---

## Complete Examples and Walkthrough

This skill includes comprehensive bundled resources:

### references/schema-requirements.md
**⭐ CRITICAL REFERENCE - Read this first when creating schemas**

Complete, validated guide for PowerSchool plugin schema requirements:
- **psExtension format** (the ONLY format that works)
- Field type mappings (Integer, String, Date, Boolean)
- plugin.xml element ordering requirements
- queries_root structure rules (no subdirectories!)
- Common installation errors and fixes
- Complete working examples
- Validation checklist

**When to use:** ALWAYS reference this when creating user_schema_root files or troubleshooting installation errors. This document contains real-world validated information from successful plugin installations.

### references/checklist-undefined-lookup-bug.md
Root cause and fix for "undefined, undefined" appearing in Registration Alert Email field when school lookup keys are missing or school not yet selected. Key pattern: always `.filter(Boolean)` before joining, guard early return if key field is empty.

### references/walkthrough.md
Complete end-to-end guide for creating a plugin from scratch:
- Project setup and directory structure
- Creating each configuration file in order
- Build configuration
- Installation and deployment steps
- Testing and troubleshooting
- Version management

**When to use:** Follow this walkthrough when creating a new plugin or onboarding developers to PowerSchool plugin development.

---

## Related Skills

This skill works best in combination with:
- **powerschool-data**: Access PowerSchool data via API or database queries. Includes REST API patterns, SQL style guide, and 720-table data dictionary reference.
- **powerschool-ui**: Build user interfaces with traditional HTML or modern frameworks (Svelte/Vue). Includes build configuration, framework integration, and AngularJS patterns.
