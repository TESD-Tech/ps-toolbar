---
name: powerschool-data
description: Complete PowerSchool data access guide covering both REST API and database access patterns. Includes authentication, endpoints, SQL queries, schema reference, and guidance on when to use each approach. Use when working with PowerSchool data, making API calls, writing queries, or designing data access strategies.
---

# PowerSchool Data Access

This skill provides comprehensive guidance for accessing PowerSchool data through both the REST API and direct database access.

## Choosing Your Data Access Method

PowerSchool offers two primary methods for accessing data. Here's when to use each:

### Use the REST API When:
- **External integration**: Accessing PowerSchool from outside applications
- **Cross-platform**: Building integrations that work across different PowerSchool instances
- **Limited scope**: Only need standard student/staff/school data
- **Security priority**: Want OAuth-based authentication and built-in permissions
- **Versioned access**: Need stable, versioned API endpoints
- **Rate limiting acceptable**: Can work within API rate limits

### Use Direct Database Access When:
- **Plugin development**: Building PowerSchool plugins with internal access
- **Complex queries**: Need advanced joins, aggregations, or custom logic
- **Performance critical**: Need faster access to large datasets
- **Custom tables**: Working with custom extensions (u_* tables)
- **Bulk operations**: Processing large amounts of data efficiently
- **PowerQueries**: Leveraging reusable named queries

### Hybrid Approach
Many PowerSchool plugins use both:
- Database queries for complex internal operations
- API for external integrations or standardized access
- PowerQueries (database) callable via API endpoints

---

# Part 1: PowerSchool REST API

## API Overview

The PowerSchool API is a RESTful web service for programmatic data access.

### API Characteristics
- **RESTful**: Uses standard HTTP methods (GET, POST, PUT, DELETE)
- **JSON-based**: Requests and responses use JSON format
- **OAuth 2.0**: Uses OAuth 2.0 for authentication
- **Versioned**: API endpoints include version numbers
- **Rate limited**: Subject to rate limiting based on server configuration

## Authentication

PowerSchool API uses OAuth 2.0 with client credentials flow.

### Authentication Process
1. **Plugin Installation**: Creates client ID and secret
2. **Token Request**: Exchange credentials for access token
3. **API Calls**: Include access token in Authorization header
4. **Token Refresh**: Tokens expire and must be refreshed

### Access Token Request
```javascript
// Typical token request pattern
POST /oauth/access_token
Content-Type: application/x-www-form-urlencoded

grant_type=client_credentials
&client_id=YOUR_CLIENT_ID
&client_secret=YOUR_CLIENT_SECRET
```

### Authorization Header
```javascript
// Include token in all API requests
Authorization: Bearer {access_token}
```

### Token Management Tips
- Cache tokens to avoid unnecessary requests
- Implement token refresh logic before expiration
- Store credentials securely (never in client-side code)
- Handle authentication errors gracefully

## Common API Endpoints

### Students
- `GET /ws/v1/student/{id}` - Get single student
- `GET /ws/v1/district/student/count` - Get student count
- `POST /ws/v1/student` - Create student
- `PUT /ws/v1/student/{id}` - Update student
- `GET /ws/v1/student?q=...` - Query students

### Schools
- `GET /ws/v1/school/{id}` - Get single school
- `GET /ws/v1/district/school` - Get all schools
- `GET /ws/v1/school/{id}/student` - Get school students

### Staff
- `GET /ws/v1/staff/{id}` - Get single staff member
- `GET /ws/v1/district/staff` - Get all staff
- `POST /ws/v1/staff` - Create staff member

### Custom Tables
- `GET /ws/schema/table/{tablename}` - Get table schema
- `GET /ws/schema/table/{tablename}/record/{recordid}` - Get record
- `POST /ws/schema/table/{tablename}` - Create record
- `PUT /ws/schema/table/{tablename}/record/{recordid}` - Update record
- `DELETE /ws/schema/table/{tablename}/record/{recordid}` - Delete record

#### Saving to Extension Tables from Plugin Pages

When building PowerSchool plugin pages that need to save custom field values to extension tables (like `U_DEF_EXT_STUDENTS`), you must use a specific pattern that differs from standard custom table updates. This pattern was discovered through extensive trial and error and is critical for success.

##### Context: When to Use This Pattern

Use this pattern when:
- Building plugin customization pages (e.g., student detail pages)
- Need to save custom field values to extension tables from client-side JavaScript
- Working with PowerSchool's Angular-based UI that loads asynchronously
- Want to update extension table fields without full page reload

##### Getting Student Context with PowerSchool Template Tags

PowerSchool template tags execute server-side and can inject values directly into your JavaScript. Use these patterns to get the student context and initial field values:

```javascript
// Extract student DCID from student foreign reference number
// The studentfrn format is typically "PS_{DCID}", so we extract from position 4 onward
const studentDcid = '~(*evaluate mid(~(studentfrn),4,99999))';

// Read initial value from extension table to set UI state
// This reads directly from the U_Students_Extension table
const initialValue = '~([Students.U_Students_Extension]FIELD_NAME)';
```

**Key Points:**
- `~(studentfrn)` returns the student foreign reference number (e.g., "PS_12345")
- `~(*evaluate mid(...,4,99999))` extracts the DCID portion
- `~([Students.U_Students_Extension]FIELD_NAME)` reads current field value
- Values are typically returned as strings: '1', '0', or '' (empty for null)
- Template tags execute server-side during page render, injecting values as JavaScript strings

##### The Critical Save Pattern

**Endpoint Structure:**
```
PUT /ws/schema/table/U_DEF_EXT_STUDENTS/{studentDcid}
```

**Critical Payload Structure** (this is the key discovery):
```javascript
const payload = {
    name: 'STUDENTS',  // CRITICAL: Must reference parent table, not extension table
    tables: {
        U_DEF_EXT_STUDENTS: {  // Extension table name as key
            FIELD_NAME: 'value'  // Field updates go here
        }
    }
};
```

**Complete Save Implementation:**
```javascript
async function saveCustomField(studentDcid, fieldName, newValue) {
    const endpoint = `/ws/schema/table/U_DEF_EXT_STUDENTS/${studentDcid}`;
    const payload = {
        name: 'STUDENTS',  // Parent table - REQUIRED
        tables: {
            U_DEF_EXT_STUDENTS: {
                [fieldName]: newValue
            }
        }
    };

    try {
        const response = await fetch(endpoint, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            // Try to parse error message, fall back to status text
            const errorData = await response.json().catch(() => ({
                message: response.statusText
            }));
            throw new Error(errorData.message || `HTTP ${response.status}`);
        }

        const result = await response.json();
        console.log('Field saved successfully:', result);
        return result;
    } catch (error) {
        console.error('Error saving field:', error);
        throw error; // Re-throw for caller to handle
    }
}
```

##### Understanding the Payload Structure

**Why this structure is required:**

1. **`name: "STUDENTS"`** - Must reference the parent table, not the extension table
   - Extension tables (U_DEF_EXT_STUDENTS) are linked to parent tables (STUDENTS)
   - PowerSchool needs to know which parent table context to use
   - This is NOT optional - the save will fail without it

2. **`tables` object** - Contains nested extension table updates
   - Key is the extension table name (U_DEF_EXT_STUDENTS)
   - Value is an object with field-value pairs
   - Allows updating multiple fields in one call

3. **DCID in endpoint** - Uses the parent table's DCID, not the extension record ID
   - `studentDcid` is from the students table, not U_DEF_EXT_STUDENTS
   - PowerSchool automatically finds the related extension record
   - If no extension record exists, one is created automatically

##### Complete Working Example with UI Integration

This example shows the full pattern including MutationObserver for Angular integration:

```javascript
(function() {
  'use strict';

  // Step 1: Get student context using PowerSchool template tags
  const studentDcid = '~(*evaluate mid(~(studentfrn),4,99999))';
  const initialValue = '~([Students.U_Students_Extension]ONLINE_REG_PIPELINE)';
  const fieldName = 'ONLINE_REG_PIPELINE';

  // Step 2: Wait for Angular components to load using MutationObserver
  const observer = new MutationObserver((mutations, obs) => {
    // Look for a specific element that indicates page is ready
    const anchor = document.querySelector('pss-add-favorite a[id="btnAddRemoveFavorite"]');

    // Only inject once
    if (anchor && !document.querySelector('.custom-field-container')) {
      const parent = anchor.closest('pss-add-favorite');
      if (parent && parent.parentNode) {

        // Step 3: Create UI element (checkbox example)
        const container = document.createElement('div');
        container.className = 'custom-field-container';
        container.innerHTML = `
          <input type="checkbox" id="${fieldName}">
          <label for="${fieldName}">Enable Feature</label>
        `;

        const checkbox = container.querySelector('input');

        // Step 4: Initialize checkbox state from template value
        // PowerSchool often returns '1' or '0' as strings, or empty string if null
        checkbox.checked = (initialValue === '1');

        // Step 5: Handle changes with save and error handling
        checkbox.addEventListener('change', async (e) => {
          const newValue = e.target.checked ? '1' : '0';
          const endpoint = `/ws/schema/table/U_DEF_EXT_STUDENTS/${studentDcid}`;

          // Step 6: Build payload with critical structure
          const payload = {
              name: 'STUDENTS',  // CRITICAL: Parent table name
              tables: {
                  U_DEF_EXT_STUDENTS: {
                      [fieldName]: newValue
                  }
              }
          };

          try {
            const response = await fetch(endpoint, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                const errorData = await response.json().catch(() => ({
                    message: response.statusText
                }));
                throw new Error(errorData.message || `HTTP ${response.status}`);
            }

            const result = await response.json();
            console.log('Field saved successfully:', result);

          } catch (error) {
             console.error('Error saving field:', error);
             alert('Error saving change: ' + error.message);

             // Step 7: Revert UI state on error
             e.target.checked = !e.target.checked;
          }
        });

        // Step 8: Insert into DOM
        parent.parentNode.insertBefore(container, parent.nextSibling);

        // Step 9: Stop observing once injected
        obs.disconnect();
      }
    }
  });

  // Start observing for Angular components
  observer.observe(document.body, { childList: true, subtree: true });
})();
```

##### Common Pitfalls and Solutions

**❌ What NOT to do:**

1. **Don't use extension table record ID in endpoint**
   ```javascript
   // WRONG - using extension table's own ID
   PUT /ws/schema/table/U_DEF_EXT_STUDENTS/{extensionRecordId}
   ```
   **Solution:** Use the parent table's DCID (studentDcid)

2. **Don't omit the `name` field**
   ```javascript
   // WRONG - missing parent table reference
   {
     tables: { U_DEF_EXT_STUDENTS: { FIELD: 'value' } }
   }
   ```
   **Solution:** Always include `name: "STUDENTS"`

3. **Don't use flat payload structure**
   ```javascript
   // WRONG - flat structure doesn't work
   {
     name: 'STUDENTS',
     FIELD_NAME: 'value'
   }
   ```
   **Solution:** Nest fields under `tables.U_DEF_EXT_STUDENTS`

4. **Don't forget error handling and UI reversion**
   ```javascript
   // WRONG - no error handling
   checkbox.addEventListener('change', (e) => {
     fetch(endpoint, { ... });  // Fire and forget
   });
   ```
   **Solution:** Use async/await with try/catch and revert UI on error

5. **Don't ignore async page loading**
   ```javascript
   // WRONG - assumes elements exist immediately
   const element = document.querySelector('.some-element');
   element.addEventListener(...);  // May be null
   ```
   **Solution:** Use MutationObserver to wait for Angular components

**✅ What TO do:**

- ✅ Use `~(*evaluate mid(~(studentfrn),4,99999))` to extract student DCID
- ✅ Include `name: "STUDENTS"` in every payload
- ✅ Nest field updates under `tables.U_DEF_EXT_STUDENTS`
- ✅ Implement comprehensive error handling with UI state reversion
- ✅ Use MutationObserver to wait for Angular components to render
- ✅ Test checkbox/UI state initialization with PowerSchool template values
- ✅ Use `response.ok` check before parsing JSON
- ✅ Log successes to console for debugging
- ✅ Provide user feedback (alerts) on errors

##### Extension Tables for Other Parent Tables

This pattern works for other extension tables too. Adjust the parent table name:

```javascript
// For teachers
const payload = {
    name: 'TEACHERS',  // Parent table
    tables: {
        U_DEF_EXT_TEACHERS: {
            FIELD_NAME: 'value'
        }
    }
};

// For schools
const payload = {
    name: 'SCHOOLS',  // Parent table
    tables: {
        U_DEF_EXT_SCHOOLS: {
            FIELD_NAME: 'value'
        }
    }
};
```

**Pattern:** The `name` field always references the parent table, while the endpoint and `tables` object reference the extension table.

##### Cross-References

- See [Custom Table Queries](#custom-table-queries) for reading extension table data via SQL
- See [API Usage Patterns](#api-usage-patterns) for general error handling approaches
- See [PowerSchool Plugin Development](../powerschool-plugin-development/SKILL.md) for information about extension table definitions in user_schema_root/
- For more about PowerSchool template tags, see the powerschool-ui skill

### PowerQueries via API
- `POST /ws/schema/query/{queryname}` - Execute named query
- Execute PowerQueries defined in queries_root/ via API
- More efficient than standard API queries
- Can include complex joins and custom logic

## API Usage Patterns

### Making API Calls from Plugins

#### Server-Side Pattern (Recommended)
Make API calls from server-side code (not client JavaScript) to protect credentials:
- Use server-side languages (Java, Python, etc.)
- Store credentials in secure server configuration
- Proxy API calls through your plugin's backend

#### Client-Side Pattern (Use with Caution)
If making calls from browser:
- Never expose client credentials
- Use session-based authentication
- Implement proper CORS handling
- Consider security implications

### Error Handling
```javascript
// Always handle API errors gracefully
try {
  const response = await fetch(apiEndpoint, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    }
  });

  if (!response.ok) {
    // Handle HTTP errors
    throw new Error(`API error: ${response.status}`);
  }

  const data = await response.json();
  return data;
} catch (error) {
  // Handle network or parsing errors
  console.error('API call failed:', error);
  // Implement fallback or user notification
}
```

### Pagination
Many endpoints return paginated results:
- Use `page` and `pagesize` parameters
- Default page size is typically 100
- Check response for total count and pagination metadata

### Query Filters
PowerSchool API supports filtering:
- Use `q` parameter for queries
- Syntax: `field==value` or `field=ge=value`
- Multiple conditions: `field1==value1;field2==value2`
- Operators: `==` (equals), `!=` (not equals), `=ge=` (greater than or equal), `=le=` (less than or equal)

## API Best Practices

### Performance
- **Cache responses**: Don't repeatedly fetch the same data
- **Use specific queries**: Fetch only what you need
- **Batch requests**: Group related operations when possible
- **Respect rate limits**: Implement backoff strategies

### Security
- **Never expose credentials**: Keep client secrets server-side
- **Validate input**: Sanitize data before API calls
- **Handle errors**: Don't expose API errors to end users
- **Use HTTPS**: Always use secure connections

### Data Management
- **Validate responses**: Check data structure and completeness
- **Handle missing data**: Not all fields may be populated
- **Respect data types**: Follow PowerSchool's data type definitions
- **Use transactions**: For multi-step operations requiring consistency

## API Response Formats

### Successful Response
```json
{
  "students": {
    "student": [
      {
        "id": 12345,
        "student_number": "100001",
        "first_name": "John",
        "last_name": "Doe",
        "grade_level": 9
      }
    ]
  }
}
```

### Error Response
```json
{
  "message": "Invalid access token",
  "error": "invalid_token"
}
```

## API Troubleshooting

### Authentication Issues
- **Invalid credentials**: Verify client ID and secret
- **Token expired**: Implement token refresh logic
- **Missing Authorization header**: Ensure header is properly set
- **Wrong grant type**: Use client_credentials flow

### Request Errors
- **404 Not Found**: Check endpoint URL and resource ID
- **400 Bad Request**: Validate request body and parameters
- **403 Forbidden**: Check permissions and access rights
- **429 Too Many Requests**: Implement rate limiting and backoff

### Data Issues
- **Empty responses**: Verify query parameters and filters
- **Unexpected format**: Check API version and response structure
- **Missing fields**: Not all fields are available in all contexts
- **Type mismatches**: Ensure data types match API expectations

### Network Issues
- **Timeout**: Increase timeout or optimize queries
- **CORS errors**: Check CORS configuration (client-side calls)
- **Connection refused**: Verify PowerSchool server is accessible
- **SSL errors**: Check certificate configuration

---

# Part 2: PowerSchool Database Access

## Database Dictionary Reference

A comprehensive PowerSchool data dictionary is available in this skill directory, organized into **15 topic-based files** for efficient navigation. The dictionary covers **720 core tables** from PowerSchool SIS Release 25.9.0, plus **5 Pennsylvania state plugin tables**.

### Quick Reference by Topic

| File | Tables | Description |
|------|--------|-------------|
| `dict-students-core.md` | 40 | Students, guardians, contacts, demographics |
| `dict-attendance.md` | 36 | Attendance records, codes, tracking |
| `dict-scheduling.md` | 56 | Bell schedules, calendars, periods, terms |
| `dict-courses-sections.md` | 21 | Courses, sections, enrollments (CC table) |
| `dict-grades-assignments.md` | 70 | Gradebook, assignments, standards, scores |
| `dict-schools-district.md` | 14 | Schools, departments, district config |
| `dict-staff-teachers.md` | 11 | Teachers, staff, users |
| `dict-health.md` | 45 | Health records, immunizations, screenings |
| `dict-incidents.md` | 16 | Discipline, behavior tracking |
| `dict-fees-financials.md` | 9 | Fee management, payments |
| `dict-careertech-graduation.md` | 50 | Graduation plans, CTE programs |
| `dict-testing.md` | 6 | Assessments, test scores |
| `dict-reporting.md` | 102 | Reports, state extracts, data export |
| `dict-system-admin.md` | 244 | Plugins, auth, roles, system tables |
| `dict-pa-state-tables.md` | 5 | Pennsylvania state plugin tables (S_PA_*, S_STU_EDFI_X) — join patterns and key fields |
| `dict-pa-student-fields.md` | 6 | Full PA student extension field lists: S_PA_STU_X (48 fields), S_PA_STU_Homeless_X, S_PA_STU_CTE_X, S_PA_REN_X, S_STU_CRDC_X, S_PA_SEN_SPED_X |
| `dict-pa-staff-courses.md` | 11 | PA course/section/staff/school extension fields: S_PA_CRS_X, S_PA_SEC_X, S_PA_CC_X, S_PA_SSF_X, S_PA_GEN_X, S_PA_LOG_X, S_PA_SCH_X, S_PA_USR_X, and CRDC variants |
| `dict-pa-lookup-codes.md` | — | PA lookup code tables: EL status codes, LIEP type codes, NCES language code schema, attendance codes |

**Index file**: `dict-index.md` contains a master alphabetical table lookup and topic navigation.

### When a Table Isn't in the Dictionary

The dictionary covers core PowerSchool SIS tables. Two categories of tables are intentionally absent:

- **State plugin tables** (`S_PA_*`, `S_OH_*`, `S_CA_*`, etc.) — installed by state-specific reporting plugins, not part of core PowerSchool
- **District custom tables** (`U_*` or other local tables) — created by the district's own schema definitions

**Before guessing at field names or join patterns, search the project codebase for existing usage.** Report HTML files in this project already use many of these tables with proven, working join patterns. Finding an existing example is far more reliable than inference.

Search the report directory for the table name:
```bash
grep -ril "S_PA_LANGUAGE_CODE_S" src/powerschool/WEB_ROOT/admin/tesd_custom_reports/
```

Or search across all HTML and JSON files:
```bash
grep -ri "TABLE_NAME_HERE" src/powerschool/WEB_ROOT/admin/tesd_custom_reports/*.html
```

Once you find a file using the table, read that file to extract the exact field names and JOIN syntax that is already proven to work in this environment.

### How to Use the Dictionary

1. **Find your topic**: Use the table above to identify which file contains your tables
2. **Look up specific tables**: Check `dict-index.md` for alphabetical lookup
3. **Browse by topic**: Each topic file has a table of contents at the top

## Common PowerSchool Tables

PowerSchool's database contains hundreds of tables. Here are the most commonly used:

### Core Student Tables
- **Students**: Main student table with demographic information
  - Contains: student_number, first_name, last_name, grade_level, etc.
  - Primary key: id
  - Most commonly queried table

- **StudentCoreFields**: Core student data fields
  - Extended student information
  - Custom fields and extensions

- **Enrollment**: Student enrollment records
  - Track student school assignments
  - Entry and exit dates

- **Attendance**: Daily attendance records
  - Daily attendance tracking
  - Attendance codes and patterns

- **Attendance_Code**: Attendance code definitions
  - Defines attendance types (present, absent, tardy, etc.)
  - School-specific configurations

### School Tables
- **Schools**: School information
  - School names, identifiers, addresses
  - School configuration data

### Course and Scheduling
- **Courses**: Course definitions
  - Course names, numbers, credits
  - Course catalog information

- **Sections**: Course sections (class instances)
  - Specific class sections
  - Teacher assignments
  - Scheduling information

- **CC**: Course/Class table (student enrollments)
  - Links students to sections
  - Grade and credit tracking

### Staff Tables
- **Teachers**: Teacher/staff information
  - Staff demographics
  - Contact information

### Grades and Assignments
- **StoredGrades**: Final grades storage
  - Report card grades
  - Term grades

- **Assignment**: Individual assignments
  - Gradebook assignments
  - Assignment details

- **AssignmentScore**: Student assignment scores
  - Individual assignment grades
  - Score tracking

### Custom Tables
- **u_*** tables**: User-defined custom tables
  - Created via user_schema_root/
  - Custom plugin data

## SQL Style Guide

When writing SQL queries for PowerSchool, follow these formatting conventions for better readability and easier debugging:

### Formatting Conventions

**No Table Aliases:**
Avoid table aliases. Use full table names throughout the query. This makes queries:
- More explicit and self-documenting
- Easier to understand when returning to code after time away
- Simpler for developers unfamiliar with your alias conventions
- More searchable (you can grep for "students.student_number" without knowing the alias)
- Less prone to confusion when multiple queries use different alias conventions

*Exception: Use aliases only when absolutely necessary (e.g., self-joins where the same table appears multiple times).*

**Leading Commas:**
Use leading commas (comma at the start of the line) instead of trailing commas. This makes it easier to:
- Comment out lines during debugging without syntax errors
- Visually scan the query structure
- Quickly identify missing commas (they're all aligned on the left)

**Indentation:**
Use tabs for indentation to show query structure clearly. Indent:
- SELECT columns (after the SELECT keyword)
- JOIN conditions (after the FROM)
- WHERE conditions (after the WHERE keyword)
- ORDER BY columns (after the ORDER BY keyword)

**Line Breaks:**
Put each clause element on its own line:
- Each SELECT column on its own line
- Each JOIN on its own line
- Each WHERE condition on its own line (for complex queries)
- Each ORDER BY column on its own line

### Example Format

```sql
-- Good: No aliases, leading commas, tabs, one item per line
SELECT students.student_number
	, students.first_name
	, students.last_name
	, students.grade_level
FROM students
WHERE students.enroll_status = 0
	AND students.schoolid = 123
ORDER BY students.student_number
```

```sql
-- Bad: Trailing commas, aliases, everything on one line
SELECT s.student_number, s.first_name, s.last_name FROM students s WHERE s.enroll_status = 0
```

## Database Query Patterns

### Basic Query Structure

```sql
-- Select student information
SELECT students.student_number
	, students.first_name
	, students.last_name
	, students.grade_level
FROM students
WHERE students.enroll_status = 0  -- Currently enrolled
	AND students.schoolid = 123     -- Specific school
```

### Joining Tables

```sql
-- Student with school information
SELECT students.student_number
	, students.first_name
	, students.last_name
	, schools.name AS school_name
FROM students
	INNER JOIN schools ON students.schoolid = schools.school_number
WHERE students.enroll_status = 0
```

### Attendance Queries

```sql
-- Student attendance with codes
SELECT students.student_number
	, attendance.att_date
	, attendance_code.att_code
	, attendance_code.description
FROM students
	INNER JOIN attendance ON students.id = attendance.studentid
	INNER JOIN attendance_code ON attendance.attendance_codeid = attendance_code.id
WHERE attendance.att_date >= SYSDATE - 30  -- Last 30 days
ORDER BY students.student_number
	, attendance.att_date
```

### Course Enrollment

```sql
-- Student course enrollments
SELECT students.student_number
	, courses.course_name
	, sections.section_number
	, teachers.last_name AS teacher
FROM students
	INNER JOIN cc ON students.id = cc.studentid
	INNER JOIN sections ON cc.sectionid = sections.id
	INNER JOIN courses ON sections.course_number = courses.course_number
	LEFT JOIN teachers ON sections.teacher = teachers.id
WHERE cc.termid >= 2500  -- Current term (adjust as needed)
	AND students.enroll_status = 0
ORDER BY students.student_number
	, courses.course_name
```

## Field Lookup Strategies

### Finding Fields by Concept
When looking for fields, consider:
1. **Table name hints**: Use the data dictionary to search for relevant tables
2. **Field naming conventions**: PowerSchool uses descriptive field names
3. **Related tables**: Many concepts span multiple tables
4. **Custom extensions**: Check for u_* custom tables and fields

### Common Field Patterns
- **ID fields**: Primary keys usually named `id`
- **Foreign keys**: Often named `{table}id` (e.g., studentid, schoolid)
- **Dates**: Often end in `_date` (e.g., att_date, entry_date)
- **Codes**: Often end in `_code` or `code`
- **Custom fields**: User-defined fields often prefixed with `u_`

### Using the Data Dictionary Files

Each topic file contains:
1. Table of contents linking to all tables in that file
2. Table name and version information
3. Field names and data types
4. Field descriptions

**Quick lookup workflow**:
1. Need student data? → Open `dict-students-core.md`
2. Need attendance tables? → Open `dict-attendance.md`
3. Not sure which file? → Check `dict-index.md` for alphabetical lookup

For detailed examples using specific tables, see the **powerschool-plugin-development** skill's queries_root examples.

## Database Best Practices

### Query Performance
- **Use indexes**: Query indexed fields when possible (id, student_number, etc.)
- **Limit results**: Use WHERE clauses to filter data
- **Avoid SELECT ***: Select only needed columns
- **Use appropriate joins**: INNER JOIN vs LEFT JOIN based on needs

### Data Integrity
- **Check for NULLs**: Many fields may be NULL
- **Validate IDs**: Verify foreign key relationships exist
- **Use transactions**: For operations requiring consistency
- **Respect constraints**: Follow primary key and unique constraints

### Security
- **Sanitize inputs**: Prevent SQL injection
- **Limit permissions**: Use appropriate database user permissions
- **Audit access**: Log database queries in production
- **Protect sensitive data**: Handle PII appropriately

### PowerSchool-Specific Considerations
- **enroll_status**: Filter by `enroll_status = 0` for currently enrolled students
- **schoolid**: Most queries need school context
- **termid**: Academic terms affect many queries
- **yearid**: School year context (negative numbers, e.g., -1 for current year)

## Common Database Patterns

### Current Students

```sql
WHERE enroll_status = 0
```

### Active Records

```sql
WHERE active = 1
```

### Date Ranges

```sql
WHERE att_date BETWEEN TO_DATE('2024-01-01', 'YYYY-MM-DD')
	AND TO_DATE('2024-12-31', 'YYYY-MM-DD')
```

### Custom Table Queries

```sql
-- Query custom table linked to students
SELECT students.student_number
	, u_custom_table.custom_field1
	, u_custom_table.custom_field2
FROM students
	INNER JOIN u_custom_table ON students.dcid = u_custom_table.studentsdcid
WHERE students.enroll_status = 0
ORDER BY students.student_number
```

## Database Troubleshooting

### Query Errors
- **Invalid column**: Check field name spelling and table
- **Invalid table**: Verify table exists and name is correct
- **No rows returned**: Check WHERE clause and data existence
- **Type mismatch**: Ensure correct data types in comparisons

### Performance Issues
- **Slow queries**: Add appropriate indexes, optimize joins
- **Timeout**: Reduce result set, add filters
- **Locking**: Avoid long-running transactions

### Data Issues
- **NULL values**: Check for NULL in conditions (IS NULL vs = NULL)
- **Incorrect relationships**: Verify foreign key joins
- **Missing data**: Check enrollment status, active flags
- **Date issues**: Verify date format and timezone handling

---

# Part 3: PowerQueries

PowerQueries are reusable named queries defined in the queries_root/ folder of PowerSchool plugins.

## PowerQueries Overview

### What are PowerQueries?
- Named SQL queries stored as XML files
- Defined server-side in queries_root/
- Callable from pages or API endpoints
- More secure than dynamic SQL
- Reusable across multiple contexts

### When to Use PowerQueries
- **Reusable queries**: Same query needed in multiple places
- **Security-sensitive**: Avoid SQL injection risks
- **Complex logic**: Joins, aggregations, custom calculations
- **API integration**: Expose complex queries via API endpoints
- **Maintenance**: Easier to update centralized queries

### PowerQueries vs Direct SQL

**PowerQueries (Recommended)**:
- Defined in queries_root/
- Server-side execution
- Better security
- Easier maintenance
- Reusable across pages and API
- Versioned with plugin

**Direct SQL (Use with Caution)**:
- Dynamic queries
- More flexible
- Requires careful input validation
- Higher security risk
- Harder to maintain

## PowerQuery Definition (XML)

Basic structure in queries_root/:

```xml
<query name="my_custom_query">
  <description>Description of what this query does</description>
  <sql>
    SELECT students.student_number
    	, students.first_name
    	, students.last_name
    FROM students
    WHERE students.schoolid = :schoolid
    	AND students.enroll_status = 0
  </sql>
</query>
```

## Calling PowerQueries

### From Pages
Use PowerSchool tags to execute queries in pages.

### Via API
```javascript
POST /ws/schema/query/my_custom_query
Content-Type: application/json

{
  "schoolid": 123
}
```

### With Parameters
PowerQueries support parameterized queries:
- Use `:parameter_name` in SQL
- Pass parameters when calling the query
- Prevents SQL injection

## Custom Table Development

When creating custom tables via user_schema_root/:

### Naming Conventions
- Use `u_` prefix for custom tables
- Use `u_` prefix for custom fields in existing tables
- Use descriptive names

### Linking to Core Tables
- Use `{table}dcid` for foreign keys (e.g., studentsdcid)
- DCID is PowerSchool's global unique identifier
- Maintain referential integrity

### Field Types
- Choose appropriate Oracle data types
- Consider field length requirements
- Use NOT NULL constraints appropriately

## PSHTML Template Tag Limitations

### Extension Table Access in PSHTML Conditionals

**CRITICAL:** State plugin extension tables (like `S_PA_STU_X`, `S_OH_*`, etc.) **CANNOT** be directly referenced in PSHTML conditionals using the standard `~([Students.TableName]field)` syntax.

**❌ This will NOT work:**
```html
~[if#isELD.~([Students.S_PA_STU_X]LEP_ELL_Status_Code)=01]
<!-- This fails because S_PA_STU_X is a state extension table -->
[/if#isELD]
```

**✅ Use tlist_sql instead:**
```html
~[tlist_sql;
SELECT CASE WHEN S_PA_STU_X.LEP_ELL_Status_Code = '01' THEN '1' ELSE '0' END as is_eld
FROM students 
LEFT JOIN S_PA_STU_X ON S_PA_STU_X.studentsdcid = students.dcid
WHERE students.id = ~(curstudid)]

~[if#isELD.~(is_eld)=1]
<!-- ELD-specific content -->
[else#isELD]
<!-- Non-ELD content -->
[/if#isELD]

[/tlist_sql]
```

### What Works vs. What Doesn't

**Direct PSHTML field access (`~([Students]field)`) works for:**
- Core PowerSchool tables (students, schools, etc.)
- District custom fields added to core tables
- Standard PowerSchool extension tables

**Requires tlist_sql approach:**
- State plugin tables (`S_PA_*`, `S_OH_*`, `S_CA_*`, etc.)
- Complex joins between multiple tables
- Calculated fields or CASE statements
- Custom tables (`U_*`) not directly linked to core tables

### Why This Limitation Exists

State extension tables are installed by separate plugins and aren't part of PowerSchool's core template tag system. The PSHTML processor only has direct access to core schema tables, not plugin-added extensions.

**Pattern for any state extension field access:**
1. Use `tlist_sql` block to query the extension table
2. Join to core table via `studentsdcid` or appropriate foreign key
3. Use `WHERE students.id = ~(curstudid)` for current student context
4. Return simple values (strings/numbers) for use in conditionals
5. Reference the returned field alias in PSHTML conditionals

---

# Part 4: Combining Data Access Methods

## Hybrid Strategies

Many PowerSchool plugins use multiple data access methods:

### Example: Admin Dashboard
- **API**: Fetch student counts and summary data
- **Database**: Complex attendance reports with joins
- **PowerQueries**: Reusable queries for common reports

### Example: Teacher Portal
- **Database**: Real-time gradebook calculations
- **PowerQueries via API**: Student rosters with custom fields
- **API**: Integration with external grading systems

### Example: Parent Portal
- **API**: Standard student information
- **PowerQueries**: Custom attendance summaries
- **Database**: (Avoid direct access from public-facing portals)

## Data Access Decision Flow

```
Need to access PowerSchool data?
├─ External integration?
│  └─ YES → Use REST API
├─ Complex joins/aggregations?
│  └─ YES → Use Database/PowerQueries
├─ Custom tables (u_*)?
│  └─ YES → Use Database
├─ Reusable query needed?
│  └─ YES → Create PowerQuery (callable via API or direct)
└─ Standard CRUD operations?
   └─ Use API (simpler, more stable)
```

## Client-Side Data Access: psQuery

### Overview

The `psQuery` service provides client-side JavaScript access to PowerSchool's database through the psQuery API endpoints. This allows AngularJS plugins to perform CRUD operations without server-side code.

**When to use psQuery:**
- Building AngularJS-based plugins
- Client-side data manipulation without server endpoints
- Quick prototypes and simple data operations
- Portal-aware data access (admin/teachers/guardian)

**When NOT to use:**
- Complex server-side logic required
- Building with modern frameworks (use REST API instead)
- Operations requiring more than 20 fields at once
- Complex transactions or batch operations

### Portal-Aware Endpoints

psQuery automatically detects the current portal and uses appropriate endpoints:
- **Admin:** `/admin/psQuery/psQueryA.html`
- **Teachers:** `/teachers/psQuery/psQueryT.html`
- **Guardian:** `/guardian/psQuery/psQueryP.html`

### Basic Usage Pattern

```javascript
// In AngularJS controller
angular.module('myApp', ['psQueryModule'])
    .controller('MyCtrl', ['$scope', '$psq', function($scope, $psq) {
        // $psq service is available
    }]);
```

### CRUD Operations

#### Get Data (Admin Portal Only)

```javascript
// Query students by grade level
$psq('students')
    .get('grade_level = 9', ['first_name', 'last_name', 'student_number'], function(results) {
        console.log(results);
        // [{ first_name: 'John', last_name: 'Doe', student_number: '123456' }, ...]
    });
```

**Important:** The `.get()` method only works in admin portal. Teacher and guardian portals will throw an error.

#### Insert Record

```javascript
$psq('students')
    .insert({
        first_name: 'Jane',
        last_name: 'Smith',
        grade_level: 10,
        enroll_status: 0
    }, function(newId) {
        console.log('New student ID:', newId);
    });
```

#### Update Record

```javascript
$psq('students')
    .update(12345, {
        grade_level: 11,
        enroll_status: 0
    }, function(id) {
        console.log('Updated student:', id);
    });
```

#### Delete Record

```javascript
$psq('students')
    .delete(12345, function(id) {
        console.log('Deleted student:', id);
    });
```

#### Insert Child Record (Parent-Child Relationship)

```javascript
// Insert attendance record for a student
$psq('attendance')
    .insertChild(
        { table: 'students', id: 12345 },
        {
            att_date: '01/15/2024',
            attendance_codeid: 1,
            att_mode_code: 'ATT_ModeDaily'
        },
        function(newAttendanceId) {
            console.log('New attendance ID:', newAttendanceId);
        }
    );
```

### Table Name Conventions

**Stock Tables:** Use table name directly (dcid as ID field)
```javascript
$psq('students')
$psq('attendance')
$psq('cc')
```

**Custom Tables:** Use full U_* name (id as ID field)
```javascript
$psq('U_MY_CUSTOM_TABLE')
$psq('U_STUDENT_PROGRAMS')
```

### Stock Table Reference

psQuery includes mappings for 224 PowerSchool stock tables. Common tables include:

- `students` (001) - Main student table
- `courses` (002) - Course catalog
- `sections` (003) - Course sections
- `cc` (004) - Course enrollments
- `teachers` (005) - Teacher records
- `attendance` (157) - Attendance records
- `attendance_code` (156) - Attendance codes
- `terms` (013) - School terms
- `schools` (039) - School information

### Field Limits

**Maximum 20 fields per operation** (`psQuery.maxParams = 20`)

For more than 20 fields, split into multiple operations:
```javascript
// First update
$psq('students').update(id, { /* fields 1-20 */ }, function() {
    // Second update
    $psq('students').update(id, { /* fields 21-40 */ }, callback);
});
```

### Date and Format Handling

psQuery expects specific formats:

**Dates:**
- Oracle format: `YYYY-MM-DD` (e.g., `2024-01-15`)
- psQuery format: `MM/DD/YYYY` (e.g., `01/15/2024`)

**Times:**
- Seconds: `0` to `86399` (0 = midnight, 43200 = noon)
- String: `HH:MM AM/PM` (e.g., `08:30 AM`)

Use date/time directives to handle format conversion automatically. See **powerschool-ui** skill for directive templates.

### Portal Permissions

Different portals have different data access:
- **Admin:** Full database access
- **Teacher:** Limited to teacher's students and sections
- **Guardian:** Limited to guardian's children

Always respect portal boundaries in your queries.

### Implementation Details

For complete implementation including:
- Full psQuery service code (`psQueryFactory.js`)
- Integration with AngularJS modules
- Usage examples and patterns
- Portal-aware service development

See **powerschool-ui** skill, Part 3: AngularJS Integration.

### Template Resources

Template files available in powerschool-ui skill:
- `resources/angularjs-templates/psQueryFactory.js` - Complete service implementation
- `resources/angularjs-patterns.md` - Comprehensive usage documentation

---

## Performance Considerations

### API Performance
- Rate limiting can slow down bulk operations
- Network latency for each request
- JSON serialization overhead
- Best for: Smaller datasets, external integrations

### Database Performance
- Direct access is faster
- Better for complex queries
- Can handle large datasets efficiently
- Best for: Internal operations, bulk processing

### PowerQueries Performance
- Optimized server-side execution
- Can be called via API or directly
- Reusable reduces code duplication
- Best for: Complex, frequently-used queries

## Security Considerations

### API Security
- OAuth 2.0 authentication
- Built-in permission checks
- Versioned endpoints (stable contracts)
- Easier to audit external access

### Database Security
- Direct access requires careful permission management
- SQL injection risk with dynamic queries
- Use parameterized queries/PowerQueries
- Best for: Internal, trusted plugin code

## Updating the PA Dictionary

The PA state reporting dictionary files (`dict-pa-student-fields.md`, `dict-pa-staff-courses.md`, `dict-pa-lookup-codes.md`) are generated from the live PA compliance documentation site. When the user says the PA docs are stale or a new PA state reporting installer has been released, regenerate them with:

```bash
node ~/.claude/skills/powerschool-data/update-pa-dict.js
```

This script (requires Node 18+, no other dependencies) crawls `https://ps-compliance.powerschool-docs.com/pssis-pa/latest`, extracts field migration tables, strips navigation/script artifacts, and writes updated dict files directly to the skill directory.

If you have a local directory of already-scraped markdown files, use:
```bash
node ~/.claude/skills/powerschool-data/update-pa-dict.js --from-cache /path/to/scraped/dir
```

The script also updates `dict-index.md` to list the generated files.

---

## Enhancing This Skill

This skill can be enhanced with:
- Real-world integration examples
- Custom authentication wrappers
- Query optimization patterns
- Hybrid data access libraries
- Project-specific PowerQueries
- Updated data dictionary for newer PowerSchool versions

## Related Skills

This skill works best in combination with:
- **powerschool-plugin-development**: Understanding where data access fits in plugin structure
- **powerschool-ui**: Building interfaces that consume this data
