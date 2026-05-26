<script lang="ts">
  import { resolveIconSvgPaths } from '$lib/icons';

  interface IconFormData {
    id: string;
    icon: string;
    href: string;
    title: string;
    description: string;
    count_sql: string;
    sort_order: number;
    disabled: string;
  }

  interface Props {
    icon: IconFormData;
    onsave: (data: Partial<IconFormData>) => void;
    oncancel: () => void;
  }

  let { icon, onsave, oncancel } = $props<Props>();

  // svelte-ignore state_referenced_locally
  let form = $state<IconFormData>({ ...icon });
  let sqlTestResult = $state('');
  let sqlTesting = $state(false);
  let sqlTestError = $state('');
  let isDisabled = $state(form.disabled === '1');

  function handleDisabledChange() {
    isDisabled = !isDisabled;
    form.disabled = isDisabled ? '1' : '0';
  }

  function handleSubmit(e: Event) {
    e.preventDefault();
    onsave({
      id: form.id,
      icon: form.icon,
      href: form.href,
      title: form.title,
      description: form.description,
      count_sql: form.count_sql,
      sort_order: form.sort_order,
      disabled: form.disabled,
    });
  }

  function setNewId() {
    form.id = form.title.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
  }

  async function testSql() {
    if (!form.count_sql.trim()) return;
    sqlTesting = true;
    sqlTestError = '';
    sqlTestResult = '';
    try {
      const res = await fetch(`/admin/ps-toolbar/sql-test.html?sql=${encodeURIComponent(form.count_sql)}`, {
        credentials: 'same-origin',
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      if (data.error) {
        sqlTestError = data.error;
      } else {
        sqlTestResult = data.query_time_ms >= 0
          ? `Returned ${data.row_count} rows in ${data.query_time_ms}ms`
          : `Returned ${data.row_count} rows`;
      }
    } catch (e: any) {
      sqlTestError = e.message;
    } finally {
      sqlTesting = false;
    }
  }

  function selectIcon(name: string) {
    form.icon = name;
  }

  const availableIcons = [
    // Ref originals
    'mail', 'clipboard', 'sports', 'language', 'health', 'people-plus',
    // Legacy
    'bell', 'alert', 'star', 'external', 'user', 'computer-alt',
    // Academic subjects
    'math', 'science', 'book', 'pencil', 'paint', 'music', 'gym', 'globe',
    // School logistics
    'bus', 'lunch', 'library', 'calendar', 'clock', 'grade', 'checklist', 'schedule', 'graduation',
    // Support services
    'heart', 'nurse',
    // Communication
    'megaphone', 'newspaper', 'phone',
    // People
    'family',
    // Facilities
    'building', 'door', 'playground', 'stage',
    // Technology
    'tablet', 'laptop', 'wifi', 'database',
    // Programs & ideas
    'lightbulb', 'accessibility', 'gear',
    // Finance
    'dollar', 'receipt',
    // Achievement & safety
    'award', 'shield', 'thumbs-up',
    // Support & goals
    'hand', 'target', 'flag',
    // Access & security
    'key', 'lock',
    // Communication & collaboration
    'chat', 'team',
    // Data & sharing
    'chart', 'upload', 'download',
    // Events
    'camera',
  ] as const;
</script>

<div class="form-overlay">
  <form class="icon-form" onsubmit={handleSubmit}>
    <h3>{icon.id ? 'Edit Icon' : 'Add Icon'}</h3>

    <label class="field">
      <span class="field-label">Title</span>
      <input 
        type="text" 
        bind:value={form.title}
        oninput={() => { if (!icon.id) setNewId(); }}
        placeholder="e.g. Messages"
        required
      />
    </label>

    <label class="field">
      <span class="field-label">ID</span>
      <input 
        type="text" 
        bind:value={form.id}
        placeholder="e.g. messages"
        required
        disabled={!!icon.id}
      />
      <span class="field-hint">Unique key used by the toolbar. Cannot be changed after creation.</span>
    </label>

    <label class="field">
      <span class="field-label">Icon</span>
      <div class="icon-selector">
        {#each availableIcons as name}
          <button 
            type="button"
            class="icon-option"
            class:selected={form.icon === name}
            onclick={() => selectIcon(name)}
            title={name}
          >
            {#if resolveIconSvgPaths(name)}
              <span class="picker-icon">
                {@html resolveIconSvgPaths(name)!}
              </span>
            {:else}
              <span>{name.charAt(0).toUpperCase()}</span>
            {/if}
          </button>
        {/each}
        <input 
          type="text" 
          bind:value={form.icon}
          placeholder="Or type custom name"
          class="icon-custom-input"
        />
      </div>
    </label>

    <label class="field">
      <span class="field-label">Target URL</span>
      <input 
        type="text" 
        bind:value={form.href}
        placeholder="e.g. /admin/messages.html"
        required
      />
    </label>

    <label class="field">
      <span class="field-label">Hover Tooltip</span>
      <input 
        type="text" 
        bind:value={form.description}
        placeholder="e.g. View your school messages"
      />
      <span class="field-hint">Shown when the user hovers over the icon in the toolbar.</span>
    </label>

    <label class="field">
      <span class="field-label">SQL Query</span>
      <textarea 
        bind:value={form.count_sql}
        rows="4"
        placeholder="SELECT COUNT(*) FROM ... WHERE ... = ~(curuserid)"
      ></textarea>
      <div class="sql-actions">
        <button type="button" class="btn btn-sm" onclick={testSql} disabled={sqlTesting || !form.count_sql.trim()}>
          {sqlTesting ? 'Testing...' : 'Test SQL'}
        </button>
        {#if sqlTestResult}
          <span class="sql-result ok">{sqlTestResult}</span>
        {/if}
        {#if sqlTestError}
          <span class="sql-result error">{sqlTestError}</span>
        {/if}
      </div>
    </label>

    <div class="form-row">
      <label class="field">
        <span class="field-label">Sort Order</span>
        <input type="number" bind:value={form.sort_order} min="0" />
      </label>

      <label class="field checkbox-field">
        <input type="checkbox" checked={isDisabled} onchange={handleDisabledChange} />
        <span class="field-label">Disabled</span>
      </label>
    </div>

    <div class="form-actions">
      <button type="submit" class="btn btn-primary">
        {icon.id ? 'Save Changes' : 'Create Icon'}
      </button>
      <button type="button" class="btn" onclick={oncancel}>Cancel</button>
    </div>
  </form>
</div>

<style>
  .form-overlay {
    background: #f9fafb;
    border: 1px solid #d1d5db;
    border-radius: 8px;
    padding: 20px;
    margin-bottom: 16px;
  }

  .icon-form h3 {
    margin: 0 0 16px;
    font-size: 16px;
    font-weight: 600;
  }

  .field {
    display: block;
    margin-bottom: 14px;
  }

  .field-label {
    display: block;
    font-size: 12px;
    font-weight: 600;
    color: #374151;
    margin-bottom: 4px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }

  .field-hint {
    display: block;
    font-size: 11px;
    color: #9ca3af;
    margin-top: 2px;
  }

  input[type="text"], input[type="number"], textarea {
    width: 100%;
    padding: 8px 10px;
    border: 1px solid #d1d5db;
    border-radius: 6px;
    font-size: 13px;
    font-family: inherit;
    box-sizing: border-box;
  }

  input:disabled {
    background: #f3f4f6;
    color: #9ca3af;
  }

  textarea {
    font-family: 'SF Mono', Menlo, monospace;
    font-size: 12px;
    resize: vertical;
  }

  .icon-selector {
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
    align-items: center;
  }

  .icon-option {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 34px;
    height: 34px;
    border: 2px solid transparent;
    border-radius: 6px;
    background: #fff;
    cursor: pointer;
    transition: all 0.15s;
  }

  .icon-option:hover {
    border-color: #93c5fd;
    background: #eff6ff;
  }

  .icon-option.selected {
    border-color: #2563eb;
    background: #dbeafe;
  }

  .icon-option :global(svg) {
    width: 20px;
    height: 20px;
    display: block;
  }

  .icon-custom-input {
    flex: 1;
    min-width: 120px;
    padding: 6px 8px;
    border: 1px solid #d1d5db;
    border-radius: 6px;
    font-size: 12px;
    margin-left: 4px;
  }

  .sql-actions {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-top: 6px;
  }

  .sql-result {
    font-size: 12px;
    padding: 2px 8px;
    border-radius: 4px;
  }

  .sql-result.ok {
    background: #dcfce7;
    color: #166534;
  }

  .sql-result.error {
    background: #fee2e2;
    color: #991b1b;
  }

  .form-row {
    display: flex;
    gap: 16px;
    align-items: flex-end;
  }

  .form-row .field {
    flex: 1;
  }

  .checkbox-field {
    display: flex;
    align-items: center;
    gap: 6px;
    padding-bottom: 4px;
  }

  .checkbox-field .field-label {
    margin-bottom: 0;
  }

  .form-actions {
    display: flex;
    gap: 8px;
    margin-top: 20px;
  }

  .btn {
    padding: 6px 14px;
    border: 1px solid #d1d5db;
    border-radius: 6px;
    background: #fff;
    cursor: pointer;
    font-size: 13px;
    transition: all 0.15s;
  }

  .btn:hover { background: #f3f4f6; }
  .btn-primary { background: #2563eb; color: #fff; border-color: #2563eb; }
  .btn-primary:hover { background: #1d4ed8; }
  .btn-sm { padding: 3px 8px; font-size: 12px; }
</style>