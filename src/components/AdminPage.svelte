<svelte:options customElement="ps-toolbar-admin" />
<script lang="ts">
  import { resolveIconSvgPaths } from '$lib/icons';
  import IconForm from './IconForm.svelte';

  interface IconRecord {
    dcid: number;
    id: string;
    icon: string;
    href: string;
    title: string;
    description: string;
    count_sql: string;
    sort_order: number;
    disabled: string;
  }

  type Tab = 'manage' | 'performance';

  let currentTab = $state<Tab>('manage');
  let icons = $state<IconRecord[]>([]);
  let loading = $state(true);
  let error = $state('');
  let editingIcon = $state<IconRecord | null>(null);
  let showForm = $state(false);
  let perfData = $state<any>(null);
  let perfLoading = $state(false);
  let saveMessage = $state('');

  const apiBase = '/ws/schema/table/U_TESD_PS_TOOLBAR_ICONS';

  async function fetchIcons() {
    loading = true;
    error = '';
    try {
      const res = await fetch(apiBase, { credentials: 'same-origin' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      icons = Array.isArray(data) ? data : [];
    } catch (e: any) {
      error = `Failed to load icons: ${e.message}`;
      icons = [];
    } finally {
      loading = false;
    }
  }

  async function fetchPerfData() {
    perfLoading = true;
    try {
      const res = await fetch('/tmp/ps-toolbar-perf-log.json', { credentials: 'same-origin' });
      if (res.ok) {
        perfData = await res.json();
      }
    } catch {
      perfData = null;
    } finally {
      perfLoading = false;
    }
  }

  async function deleteIcon(record: IconRecord) {
    if (!confirm(`Delete icon "${record.title}" (${record.id})?`)) return;
    try {
      const res = await fetch(`${apiBase}/${record.dcid}`, {
        method: 'DELETE',
        credentials: 'same-origin',
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      icons = icons.filter(i => i.dcid !== record.dcid);
      showSaveMessage(`Deleted "${record.title}"`);
    } catch (e: any) {
      showSaveMessage(`Delete failed: ${e.message}`);
    }
  }

  function startAdd() {
    editingIcon = null;
    showForm = true;
  }

  function startEdit(record: IconRecord) {
    editingIcon = { ...record };
    showForm = true;
  }

  function cancelForm() {
    editingIcon = null;
    showForm = false;
  }

  async function handleSave(record: Partial<IconRecord>) {
    try {
      if (editingIcon) {
        // Update existing
        const res = await fetch(`${apiBase}/${editingIcon.dcid}`, {
          method: 'PUT',
          credentials: 'same-origin',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(record),
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        showSaveMessage(`Updated "${record.title}"`);
      } else {
        // Create new — generate DCID ref
        const res = await fetch(apiBase, {
          method: 'POST',
          credentials: 'same-origin',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(record),
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        showSaveMessage(`Created "${record.title}"`);
      }
      showForm = false;
      editingIcon = null;
      await fetchIcons();
    } catch (e: any) {
      showSaveMessage(`Save failed: ${e.message}`);
    }
  }

  function showSaveMessage(msg: string) {
    saveMessage = msg;
    setTimeout(() => { saveMessage = ''; }, 3000);
  }

  $effect(() => {
    fetchIcons();
  });

  $effect(() => {
    if (currentTab === 'performance') {
      fetchPerfData();
    }
  });

  const sortedIcons = $derived(
    [...icons].sort((a, b) => (a.sort_order ?? 0) - (b.sort_order ?? 0))
  );
</script>

<div class="admin-container">
  <!-- Header -->
  <div class="admin-header">
    <h2>PS Toolbar Management</h2>
    {#if saveMessage}
      <div class="toast">{saveMessage}</div>
    {/if}
  </div>

  <!-- Tabs -->
  <div class="tabs" role="tablist">
    <button 
      class="tab" 
      class:active={currentTab === 'manage'}
      onclick={() => { currentTab = 'manage'; }}
      role="tab"
      aria-selected={currentTab === 'manage'}
    >Manage Icons</button>
    <button 
      class="tab" 
      class:active={currentTab === 'performance'}
      onclick={() => { currentTab = 'performance'; }}
      role="tab"
      aria-selected={currentTab === 'performance'}
    >Performance</button>
  </div>

  <!-- Tab: Manage Icons -->
  {#if currentTab === 'manage'}
    {#if showForm}
      <IconForm 
        icon={editingIcon ? {
          id: editingIcon.id,
          icon: editingIcon.icon,
          href: editingIcon.href,
          title: editingIcon.title,
          description: editingIcon.description,
          count_sql: editingIcon.count_sql,
          sort_order: editingIcon.sort_order,
          disabled: editingIcon.disabled,
        } : {
          id: '', icon: 'mail', href: '', title: '',
          description: '', count_sql: '', sort_order: 0, disabled: '0',
        }}
        onsave={handleSave}
        oncancel={cancelForm}
      />
    {:else if loading}
      <div class="loading">Loading icons...</div>
    {:else if error}
      <div class="error">{error}</div>
    {:else}
      <div class="toolbar-actions">
        <button class="btn btn-primary" onclick={startAdd}>+ Add Icon</button>
        <button class="btn btn-secondary" onclick={fetchIcons}>↻ Refresh</button>
      </div>

      <table class="icon-table">
        <thead>
          <tr>
            <th></th>
            <th>ID</th>
            <th>Title</th>
            <th>Icon</th>
            <th>Hover</th>
            <th>SQL</th>
            <th>Order</th>
            <th>On</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          {#each sortedIcons as icon (icon.dcid)}
            <tr class:disabled={icon.disabled === '1'}>
              <td class="icon-preview-cell">
                {#if resolveIconSvgPaths(icon.icon)}
                  <span class="preview-svg"> 
                    {@html resolveIconSvgPaths(icon.icon)!}
                  </span>
                {:else}
                  <span class="preview-fallback">{icon.title?.charAt(0) || '?'}</span>
                {/if}
              </td>
              <td><code>{icon.id}</code></td>
              <td>{icon.title}</td>
              <td><code>{icon.icon}</code></td>
              <td class="desc-cell" title={icon.description}>{icon.description || '—'}</td>
              <td class="sql-cell" title={icon.count_sql}>
                {icon.count_sql ? (icon.count_sql.length > 40 ? icon.count_sql.slice(0, 40) + '…' : icon.count_sql) : '—'}
              </td>
              <td>{icon.sort_order ?? 0}</td>
              <td>
                <span class="status-dot" class:active={icon.disabled !== '1'}></span>
              </td>
              <td class="action-cell">
                <button class="btn btn-sm" onclick={() => startEdit(icon)}>Edit</button>
                <button class="btn btn-sm btn-danger" onclick={() => deleteIcon(icon)}>Del</button>
              </td>
            </tr>
          {/each}
        </tbody>
      </table>

      {#if icons.length === 0}
        <div class="empty-state">
          <p>No icons configured yet. Click <strong>+ Add Icon</strong> to create one.</p>
        </div>
      {/if}
    {/if}

  <!-- Tab: Performance -->
  {:else if currentTab === 'performance'}
    {#if perfLoading}
      <div class="loading">Loading performance data...</div>
    {:else if perfData}
      <div class="perf-summary">
        <div class="perf-stat">
          <span class="perf-stat-value">{Object.keys(perfData).length}</span>
          <span class="perf-stat-label">Tracked Queries</span>
        </div>
      </div>

      <table class="icon-table">
        <thead>
          <tr>
            <th>Icon ID</th>
            <th>Last 5</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          {#each Object.entries(perfData) as [iconId, entries]}
            {@const lastEntries = (entries as any[]).slice(-5)}
            {@const avgMs = Math.round(lastEntries.reduce((s: number, e: any) => s + (e.query_time_ms || 0), 0) / lastEntries.length)}
            <tr>
              <td><code>{iconId}</code></td>
              <td class="perf-bars">
                {#each lastEntries as entry}
                  <span 
                    class="perf-bar" 
                    class:slow={entry.query_time_ms >= 500}
                    class:warn={entry.query_time_ms >= 200 && entry.query_time_ms < 500}
                    class:fast={entry.query_time_ms < 200}
                    title={`${entry.query_time_ms}ms at ${entry.time || '?'}`}
                    style="width: {Math.min(entry.query_time_ms / 5, 60)}px"
                  ></span>
                {/each}
              </td>
              <td>
                <span class="badge-status" class:badge-danger={avgMs >= 500} class:badge-warn={avgMs >= 200 && avgMs < 500} class:badge-ok={avgMs < 200}>
                  {avgMs < 200 ? 'Fast' : avgMs < 500 ? 'OK' : 'Slow'}
                </span>
              </td>
            </tr>
          {/each}
        </tbody>
      </table>
    {:else}
      <div class="empty-state">
        <p>No performance data yet. Performance is recorded when the feed endpoint executes queries in PowerSchool.</p>
      </div>
    {/if}
  {/if}
</div>

<style>
  :host {
    --bg: #fff;
    --border: #d1d5db;
    --text: #111827;
    --text-muted: #6b7280;
    --primary: #2563eb;
    --primary-hover: #1d4ed8;
    --danger: #dc2626;
    --danger-hover: #b91c1c;
    --ok: #16a34a;
    --warn: #d97706;
    --slow: #dc2626;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    font-size: 14px;
    color: var(--text);
  }

  .admin-container {
    padding: 16px 0;
  }

  .admin-header {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 16px;
  }

  .admin-header h2 {
    margin: 0;
    font-size: 20px;
    font-weight: 600;
  }

  .toast {
    background: var(--ok);
    color: #fff;
    padding: 6px 14px;
    border-radius: 6px;
    font-size: 13px;
    animation: fadeIn 0.2s;
  }

  @keyframes fadeIn {
    from { opacity: 0; transform: translateY(-4px); }
    to { opacity: 1; transform: translateY(0); }
  }

  .tabs {
    display: flex;
    gap: 0;
    border-bottom: 2px solid var(--border);
    margin-bottom: 16px;
  }

  .tab {
    padding: 8px 20px;
    border: none;
    background: none;
    cursor: pointer;
    font-size: 14px;
    color: var(--text-muted);
    border-bottom: 2px solid transparent;
    margin-bottom: -2px;
    transition: all 0.15s;
  }

  .tab.active {
    color: var(--primary);
    border-bottom-color: var(--primary);
    font-weight: 600;
  }

  .loading, .error, .empty-state {
    text-align: center;
    padding: 32px;
    color: var(--text-muted);
  }

  .error {
    color: var(--danger);
  }

  .toolbar-actions {
    display: flex;
    gap: 8px;
    margin-bottom: 12px;
  }

  .btn {
    padding: 6px 14px;
    border: 1px solid var(--border);
    border-radius: 6px;
    background: var(--bg);
    cursor: pointer;
    font-size: 13px;
    transition: all 0.15s;
  }

  .btn:hover { background: #f3f4f6; }
  .btn-primary { background: var(--primary); color: #fff; border-color: var(--primary); }
  .btn-primary:hover { background: var(--primary-hover); }
  .btn-danger { color: var(--danger); border-color: var(--danger); }
  .btn-danger:hover { background: var(--danger); color: #fff; }
  .btn-sm { padding: 3px 8px; font-size: 12px; }

  .icon-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
  }

  .icon-table th {
    text-align: left;
    padding: 8px 6px;
    border-bottom: 2px solid var(--border);
    font-weight: 600;
    color: var(--text-muted);
    font-size: 12px;
    text-transform: uppercase;
    white-space: nowrap;
  }

  .icon-table td {
    padding: 8px 6px;
    border-bottom: 1px solid #e5e7eb;
    vertical-align: middle;
  }

  .icon-table tr.disabled td {
    opacity: 0.4;
  }

  .icon-table tr:hover td {
    background: #f9fafb;
  }

  .icon-preview-cell {
    width: 36px;
    text-align: center;
  }

  .preview-svg {
    width: 24px;
    height: 24px;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .preview-svg :global(svg) {
    width: 100%;
    height: 100%;
  }

  .preview-fallback {
    font-weight: 700;
    font-size: 12px;
    color: var(--text-muted);
  }

  .desc-cell {
    max-width: 180px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    color: var(--text-muted);
  }

  .sql-cell {
    max-width: 200px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    font-family: 'SF Mono', Menlo, monospace;
    font-size: 11px;
    color: var(--text-muted);
  }

  .status-dot {
    display: inline-block;
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: var(--border);
  }

  .status-dot.active {
    background: var(--ok);
  }

  .action-cell {
    white-space: nowrap;
    text-align: right;
  }

  .perf-summary {
    display: flex;
    gap: 16px;
    margin-bottom: 16px;
  }

  .perf-stat {
    background: #f9fafb;
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 12px 20px;
    text-align: center;
  }

  .perf-stat-value {
    display: block;
    font-size: 24px;
    font-weight: 700;
    color: var(--text);
  }

  .perf-stat-label {
    font-size: 12px;
    color: var(--text-muted);
  }

  .perf-bars {
    display: flex;
    gap: 2px;
    align-items: center;
    min-width: 60px;
  }

  .perf-bar {
    display: inline-block;
    height: 12px;
    border-radius: 2px;
    transition: width 0.3s;
  }

  .perf-bar.fast { background: var(--ok); }
  .perf-bar.warn { background: var(--warn); }
  .perf-bar.slow { background: var(--slow); }

  .badge-status {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 10px;
    font-size: 11px;
    font-weight: 600;
  }

  .badge-ok { background: #dcfce7; color: #166534; }
  .badge-warn { background: #fef3c7; color: #92400e; }
  .badge-danger { background: #fee2e2; color: #991b1b; }
</style>