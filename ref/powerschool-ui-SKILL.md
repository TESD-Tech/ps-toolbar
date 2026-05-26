---
name: powerschool-ui
description: Building PowerSchool plugin user interfaces with traditional HTML pages or modern JavaScript frameworks (Svelte, Vue). Includes build configuration, framework integration, portal-specific patterns, and complete development workflow. Use when building PowerSchool UI, integrating frameworks, or designing plugin interfaces.
---

# PowerSchool UI Development

This skill provides comprehensive guidance for building PowerSchool plugin user interfaces, from traditional HTML pages to modern framework integration.

## Choosing Your UI Approach

PowerSchool plugins support multiple UI development approaches:

### Traditional HTML Pages
**Best for:**
- Simple forms and data displays
- Pages that closely match PowerSchool's existing UI
- Quick prototypes and basic interfaces
- Plugins with minimal interactive requirements

**Advantages:**
- Faster to develop for simple use cases
- No build process required
- Easier to match PowerSchool's existing look and feel
- Lower maintenance overhead

### Modern JavaScript Frameworks (Svelte, Vue)
**Best for:**
- Rich, interactive user experiences
- Complex state management
- Real-time data updates
- Single-page applications (SPAs)
- Reusable component libraries

**Advantages:**
- Better developer experience with modern tooling
- Reactive data binding
- Component reusability
- Better code organization for complex UIs

**For comprehensive framework integration guidance, see `references/framework-integration.md`.**

### AngularJS (Legacy)
**Best for:**
- Maintaining existing AngularJS plugins
- Extending plugins already using AngularJS
- Using PowerSchool's built-in AngularJS support
- Client-side data manipulation with psQuery

**Considerations:**
- AngularJS is in long-term support (LTS) mode
- Prefer modern frameworks for new projects
- Good for legacy maintenance and existing ecosystems

**For comprehensive AngularJS guidance, see `resources/angularjs-patterns.md` and template files in `resources/angularjs-templates/`.**

### Hybrid Approach
**Best for:**
- Large plugins with varying complexity needs
- Gradual migration from traditional to modern
- Different requirements across portals

**Pattern:**
- Traditional pages for simple views
- Framework components for complex interactions
- Mix and match based on specific needs

---

# Part 1: Traditional PowerSchool Pages

## Basic HTML Page Structure

PowerSchool pages use standard HTML with PowerSchool-specific tags and patterns.

### Minimal Page Template
```html
<!DOCTYPE html>
<html>
<head>
  <title>My Plugin Page</title>
  <link rel="stylesheet" href="/admin/my-plugin/styles.css">
</head>
<body>
  <h1>My Plugin Page</h1>

  <!-- Your content here -->

  <script src="/admin/my-plugin/script.js"></script>
</body>
</html>
```

### Portal-Specific Pages

Place pages in the appropriate portal directory based on user type:
- `WEB_ROOT/admin/` - Administrator pages
- `WEB_ROOT/teachers/` - Teacher pages
- `WEB_ROOT/guardian/` - Parent/guardian pages
- `WEB_ROOT/subs/` - Substitute teacher pages

See **powerschool-plugin-development** skill for detailed portal information.

### PowerSchool Tags and Patterns

Traditional PowerSchool pages can use:
- Server-side tags for data access (e.g., `~(students.first_name)`)
- PowerSchool session variables
- Built-in PowerSchool JavaScript libraries
- PowerSchool CSS classes for consistent styling

### Data Integration

Access PowerSchool data in traditional pages:
- Server-side data rendering with PowerSchool tags
- AJAX calls to PowerSchool REST API
- Form submissions to custom endpoints
- PowerQueries for complex data needs

See **powerschool-data** skill for data access patterns.

### Styling Traditional Pages

**Match PowerSchool UI:**
- Use PowerSchool's existing CSS: `/images/css/screen.css`
- Follow PowerSchool's visual patterns
- Maintain consistent navigation
- Use standard PowerSchool form elements

**Custom Styling:**
- Place CSS in WEB_ROOT/{portal}/
- Use scoped styles to avoid conflicts
- Consider responsive design
- Test across PowerSchool themes

### Navigation Integration

Register pages in PowerSchool navigation using `pagecataloging/`:
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
      "pageURL": "/admin/my-plugin/page.html",
      "parentHTMLID": "navStudentAcademicPerformanceSection",
      "districtLevelContext": 1
    }
  ]
}
```

See the **powerschool-nav** skill for the full property reference, `parentHTMLID` values, version requirements, and wildcard-based nav injection for teacher/guardian portals.

---

# Part 2: Modern Framework Integration

## Overview

PowerSchool plugins can leverage modern JavaScript frameworks (Svelte and Vue) to build rich, reactive user interfaces. Both frameworks compile to standard JavaScript that runs in PowerSchool's web environment.

### Framework Integration Architecture

**Basic flow:**
1. **Develop** with modern framework in src/
2. **Build** framework to production JavaScript/CSS
3. **Copy** build output to PowerSchool WEB_ROOT structure
4. **Package** complete plugin with ps-package
5. **Deploy** to PowerSchool

### Supported Frameworks
- **Svelte**: Lightweight, compile-time framework (recommended for smaller apps)
- **Vue**: Progressive JavaScript framework (recommended for larger SPAs)

## Svelte 5 + Vite — Custom Elements (Recommended)

This is the current pattern for PS plugins using Svelte. Components compile to Shadow DOM custom elements, which prevents PowerSchool page styles from leaking in.

### Project Structure
```
project-root/
├── src/
│   ├── powerschool/
│   │   └── WEB_ROOT/
│   │       └── admin/          # Portal HTML loader pages
│   │           └── my-plugin/
│   │               └── dashboard.html
│   ├── App.svelte              # Root component (defines the custom element)
│   ├── lib/                    # Components, utilities
│   │   └── injectShadowCss.ts  # CSS injection helper (required)
│   └── main.ts                 # Entry point (imports App)
├── vite.config.ts
├── package.json
└── plugin.xml
```

### vite.config.ts
```typescript
import { defineConfig } from 'vite'
import { svelte } from '@sveltejs/vite-plugin-svelte'
import path from 'path'

export default defineConfig({
  plugins: [
    svelte({
      compilerOptions: { customElement: true },
      emitCss: false,           // REQUIRED: CSS must bundle into JS for Shadow DOM
    }),
  ],
  base: `/{pluginName}/`,
  resolve: { alias: { '$lib': path.resolve(__dirname, 'src/lib') } },
  build: {
    outDir: `dist/WEB_ROOT/{pluginName}/`,
    rollupOptions: {
      input: { app: path.resolve(__dirname, 'src/main.ts') },
      output: {
        format: 'es',
        entryFileNames: '[name].js',
      },
    },
  },
})
```

### Root Component (App.svelte)
```svelte
<svelte:options customElement="my-plugin-app" />

<script lang="ts">
  // Hyphenated HTML attributes need multiple prop aliases
  let {
    portal,
    'year-id': yearId,
  } = $props<{ portal?: string; 'year-id'?: string }>()
</script>

<main>
  <!-- your app -->
</main>

<style>
  /* Styles are bundled into JS and injected into Shadow DOM */
</style>
```

**Svelte 5 runes used in PS plugins:**
- `$props()` — component props (replaces `export let`)
- `$state()` — reactive state (replaces `let` with reactivity)
- `$derived()` / `$derived.by()` — computed values (replaces `$:`)
- `onMount()` — lifecycle, still from `svelte`

**Prop aliasing pattern** — PowerSchool passes hyphenated HTML attributes; handle all casings:
```typescript
let { portal, userType, usertype, 'user-type': userTypeAttr } = $props()
```

### CSS in Shadow DOM

Because the custom element uses Shadow DOM, PowerSchool's page CSS doesn't reach inside. All CSS must be bundled into JS (`emitCss: false`) and injected at mount:

```typescript
// src/lib/injectShadowCss.ts
export function injectShadowCss(shadowRoot: ShadowRoot, css: string) {
  const style = document.createElement('style')
  style.textContent = css
  shadowRoot.appendChild(style)
}
```

Call it in `onMount` after getting the shadow root:
```typescript
onMount(() => {
  const sr = el.getRootNode()
  if (sr instanceof ShadowRoot) injectShadowCss(sr, '')
})
```

### HTML Loader Page
```html
<!DOCTYPE html>
<html>
<head>
  <title>My Plugin</title>
  ~[wc:commonscripts]
  <link href="/images/css/screen.css" rel="stylesheet" media="screen">
</head>
<body>
  ~[wc:teachers_header_css]
  ~[wc:teachers_navigation_css]
  ~[wc:teachers_nav_css]

  <my-plugin-app portal="teacher" year-id="~(curyearid)"></my-plugin-app>

  <script type="module">
    const isDev = location.hostname === 'localhost' || location.hostname === '127.0.0.1'
    const path = isDev ? '/src/main.ts' : '/my-plugin/app.js?v=~(random16)'
    import(/* @vite-ignore */ path).catch(e => console.error('load failed:', e))
  </script>

  ~[wc:teachers_footer_css]
</body>
</html>
```

The `?v=~(random16)` cache-busting param is a PowerSchool template tag that injects a random string — use it on the prod path to prevent browser caching after plugin updates.

### Build & Package Commands
```bash
pnpm dev        # Vite dev server — loads /src/main.ts directly
pnpm build      # Compiles to dist/WEB_ROOT/{pluginName}/app.js
pnpm package    # build + ps-package → ZIP in plugin_archive/
```

**Dev data:** Put a sample `eld.json` (or equivalent) in `public/` so Vite serves it at the expected relative URL during development.

For complete Vite + Svelte 5 configuration details, see `references/framework-integration.md`.

## Quick Start: Vue

### Project Structure
```
project-root/
├── src/
│   ├── powerschool/          # PowerSchool folders
│   │   └── WEB_ROOT/
│   │       └── admin/        # Loader page
│   ├── App.vue               # Vue app entry
│   └── components/           # Vue components
├── dist/                     # Vue build output
├── plugin.xml
├── package.json
└── vite.config.js
```

### Build Configuration
```json
{
  "scripts": {
    "dev": "vite",
    "build": "vite build",
    "package": "npm run build && node /path/to/ps-package/index.js"
  }
}
```

### Vite Configuration
```javascript
// vite.config.js
export default defineConfig({
  base: '/admin/{pluginName}/',
  build: {
    outDir: 'dist/WEB_ROOT/{pluginName}',
    assetsDir: 'assets'
  },
  plugins: [vue()]
});
```

## PowerSchool-Specific Considerations

### Authentication
PowerSchool handles authentication at the server level. Framework apps run in authenticated context—no need to implement authentication in framework code.

### Portal Context
Different portals have different permissions:
- **admin/**: Full administrative access
- **teachers/**: Teacher-scoped data (their students/classes)
- **guardian/**: Parent-limited views (their children)
- **subs/**: Substitute-limited views

Framework apps should respect portal context in data requests and adjust UI based on permissions.

### Data Access
- **PowerSchool REST API**: Standard data access (see **powerschool-data** skill)
- **PowerQueries**: Execute named queries via API endpoints
- **Hybrid approach**: Mix API calls and server-side rendering

### Styling
- Use scoped CSS to avoid conflicts with PowerSchool styles
- Test across different PowerSchool themes
- Consider matching PowerSchool's color schemes for consistency
- Ensure responsive design for mobile access

## Complete Framework Integration Guide

**For comprehensive framework integration documentation, including:**
- Detailed Svelte configuration (Rollup setup, build process)
- Detailed Vue configuration (Vite/webpack setup)
- Framework-specific best practices (state management, API integration)
- Development workflow (local dev, mock data, testing)
- Troubleshooting (build issues, runtime issues, performance)
- Advanced patterns (hybrid approach, multi-portal support, component embedding)

**See: `references/framework-integration.md`**

---

# Part 3: AngularJS Integration

## Overview

PowerSchool supports AngularJS for building interactive plugin interfaces. While AngularJS is in long-term support mode, it remains relevant for:
- Maintaining existing AngularJS-based plugins
- Extending plugins already using AngularJS
- Leveraging PowerSchool's built-in AngularJS libraries
- Using the psQuery service for client-side data access

**For new projects, prefer modern frameworks (Part 2) unless you have specific requirements for AngularJS.**

## When to Use AngularJS

### Good Use Cases
- **Legacy plugin maintenance** - Updating existing AngularJS plugins
- **Existing AngularJS ecosystems** - Extending plugins already using AngularJS
- **PowerSchool's built-in support** - Leveraging included AngularJS libraries
- **Client-side CRUD** - Using psQuery for data manipulation without server-side code

### When NOT to Use
- **New greenfield projects** - Modern frameworks offer better tooling and performance
- **Complex SPAs** - Modern frameworks provide better developer experience
- **Modern development workflow** - TypeScript, testing, and tooling are better with modern frameworks

## Quick Start: AngularJS

### Basic Module Structure (RequireJS/AMD)

PowerSchool uses RequireJS for module loading. All AngularJS code follows the AMD pattern:

```javascript
'use strict';
define([
    'angular',
    'components/angular_libraries/directives/myDirective',
    'components/angular_libraries/services/myService'
], function(angular, myDirective) {
    return angular.module('myPluginModule', ['dependencyModule'])
        .directive('myDirective', myDirective);
});
```

### Loading in PowerSchool Pages

```html
<!DOCTYPE html>
<html>
<head>
    <title>My AngularJS Page</title>
</head>
<body ng-controller="MyController">
    <!-- Your content -->

    <script>
    require([
        'angular',
        'components/angular_libraries/myModule'
    ], function(angular) {
        angular.bootstrap(document, ['myPluginModule']);
    });
    </script>
</body>
</html>
```

### psQuery Service for Client-Side Data Access

The `$psq` service provides client-side CRUD operations:

```javascript
angular.module('myApp', ['psQueryModule'])
    .controller('MyCtrl', ['$scope', '$psq', function($scope, $psq) {
        // Get data (admin portal only)
        $psq('students')
            .get('grade_level = 9', ['first_name', 'last_name'], function(results) {
                $scope.$apply(function() {
                    $scope.students = results;
                });
            });

        // Insert record
        $psq('students').insert({
            first_name: 'Jane',
            last_name: 'Smith'
        }, function(newId) {
            console.log('New ID:', newId);
        });

        // Update record
        $psq('students').update(12345, {
            grade_level: 11
        }, function(id) {
            console.log('Updated:', id);
        });
    }]);
```

**Important:** psQuery `.get()` only works in admin portal. Teacher/guardian portals will throw an error.

### Portal-Aware Pattern

```javascript
const isAdmin = window.location.href.includes("/admin/");
const isTeacher = window.location.href.includes("/teachers/");
const isGuardian = window.location.href.includes("/guardian/");

// Adjust behavior based on portal
if (isAdmin) {
    // Admin: full data access
} else {
    // Teachers/guardians: limited access
}
```

## AngularJS Template Resources

This skill includes production-tested AngularJS templates in `resources/angularjs-templates/`:

### Available Directives
- **dateEntryOracle.js** - Date picker with format conversion (MM/DD/YYYY ↔ YYYY-MM-DD)
- **timeEntrySeconds.js** - Time entry with seconds/string conversion
- **tooltipFollow.js** - Custom HTML tooltips that follow mouse
- **cleanTextArea.js** - Text cleaning (prevents Enter, strips quotes)

### Available Services
- **psQueryFactory.js** - Portal-aware client-side data access service

### Module Template
- **moduleTemplate.js** - AMD module structure template

**Usage example:**
```html
<input type="text"
       ng-model="startDate"
       date-entry-oracle
       data-date-entry-output="oracle">
```

## Complete AngularJS Documentation

**For comprehensive AngularJS integration documentation, including:**
- RequireJS/AMD module patterns
- Creating custom directives with ngModel integration
- psQuery service patterns and CRUD operations
- Portal-aware patterns and permission handling
- Common directive patterns (date entry, time entry, tooltips)
- Integration patterns (embedding in traditional pages)
- Best practices (style scoping, event cleanup, validation)
- Complete template directive documentation

**See: `resources/angularjs-patterns.md`**

---

## Related Skills

This skill works best in combination with:
- **powerschool-plugin-development**: Understanding PowerSchool folder structure, portal organization, and build process
- **powerschool-data**: Accessing PowerSchool data via API or database for your UI components, including REST API patterns and SQL queries
