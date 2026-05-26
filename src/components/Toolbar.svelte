<script lang="ts">
  interface ToolbarItem {
    id: string | number;
    icon: string;
    href: string;
    title: string;
    description?: string;
    count: number;
  }

  interface Props {
    feedUrl?: string;
    pollInterval?: number;
    portal?: 'admin' | 'teachers' | 'guardian' | 'unknown';
  }

  let { 
    feedUrl, 
    pollInterval = 30000,
    portal = 'unknown',
  } = $props<Props>();

  // Determine the correct feed URL based on the portal if not explicitly provided
  const resolvedFeedUrl = $derived.by(() => {
    if (feedUrl) return feedUrl;
    
    // In PowerSchool, the feed endpoint executes count_sql and returns enriched JSON
    if (portal !== 'unknown') {
      return `/${portal}/ps-toolbar/feed.html`;
    }
    
    // Fallback for local dev
    return 'ps-toolbar/feed.html';
  });

  let items = $state<ToolbarItem[]>([]);

  async function fetchFeed() {
    const url = resolvedFeedUrl;
    try {
      const res = await fetch(url, { cache: 'no-store' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      if (Array.isArray(data)) {
        items = data.map((it, idx) => {
          return {
            id: it.id ?? it.title ?? idx,
            icon: it.icon ?? '',
            href: it.href ?? '#',
            title: it.title ?? '',
            description: it.description || it.title || '',
            count: Number(it.count) || 0
          };
        });
      } else {
        items = [];
      }
    } catch (e) {
      items = [];
    }
  }

  $effect(() => {
    // Start polling
    fetchFeed();
    const interval = setInterval(fetchFeed, pollInterval);
    return () => clearInterval(interval);
  });

  import { resolveIconSvgPaths } from '$lib/icons';
  function getIconSvgPaths(name: string): string | null {
    return resolveIconSvgPaths(name);
  }
</script>

<div class="toolbar" role="toolbar" aria-label="Notifications toolbar">
  {#each items as item (item.id)}
    <div class="pds-app-action">
      <a 
        class="button-with-badge" 
        href={item.href} 
        title={item.description || item.title}
        target="_top"
        aria-label={item.description || item.title}
      >
        {#if item.icon && getIconSvgPaths(item.icon)}
          <span class="icon-svg-wrapper"> 
            {@html getIconSvgPaths(item.icon)!}
          </span>
        {:else}
          <span class="icon-fallback" aria-hidden="true">
            {item.title ? item.title.charAt(0).toUpperCase() : '?'}
          </span>
        {/if}
        
        {#if item.count > 0}
          <span class="badge">{item.count}</span>
        {/if}
      </a>
    </div>
  {/each}
</div>

<style>
  :global(.toolbar) {
    display: flex;
    list-style: none;
    margin: 0;
    padding: 0;
    align-items: center;
    color: #fff;
  }

  :global(.pds-app-action) {
    display: flex;
    align-items: center;
    justify-content: center;
    margin: 0;
    padding: 0;
  }

  :global(.button-with-badge) {
    position: relative;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 44px;
    height: 38px;
    color: #fff;
    text-decoration: none;
    transition: background-color 0.2s;
    border-radius: 4px;
  }

  :global(.button-with-badge:hover) {
    background-color: rgba(255, 255, 255, 0.1);
  }

  :global(.icon-svg-wrapper) {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 24px;
    height: 24px;
  }
  :global(.icon-svg-wrapper svg) {
    width: 100%;
    height: 100%;
  }

  :global(.icon-fallback) {
    font-size: 14px;
    font-weight: 700;
    line-height: 1;
    color: #fff;
  }

  :global(.badge) {
    position: absolute;
    top: 2px;
    right: 2px;
    background: #ef4444;
    color: white;
    font-size: 11px;
    font-weight: 600;
    padding: 0 4px;
    border-radius: 10px;
    min-width: 16px;
    height: 16px;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 1px 2px rgba(0,0,0,0.4);
    pointer-events: none;
    line-height: 1;
  }
</style>