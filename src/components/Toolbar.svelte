<script lang="ts">
  import { onMount, onDestroy } from 'svelte';
  import { injectShadowCss } from '$lib/injectShadowCss';

  interface ToolbarItem {
    id: string | number;
    icon: string;
    href: string;
    title: string;
    count: number;
  }

  interface Props {
    feedUrl?: string;
    pollInterval?: number;
    portal?: 'admin' | 'teachers' | 'guardian' | 'unknown';
    // Component reference for ShadowRoot access
    el?: HTMLElement;
  }

  let { 
    feedUrl, 
    pollInterval = 30000,
    portal = 'unknown',
    el
  } = $props<Props>();

  // Determine the correct feed URL based on the portal if not explicitly provided
  const resolvedFeedUrl = $derived(() => {
    if (feedUrl) return feedUrl;
    
    // In PowerSchool, we load the .json.txt which is processed for tlist_sql
    if (portal !== 'unknown') {
      return `/${portal}/ps-toolbar/notifications.json.txt`;
    }
    
    // Fallback for local dev
    return 'ps-toolbar/notifications.json';
  });

  let items = $state<ToolbarItem[]>([]);
  let timer: ReturnType<typeof setInterval>;

  async function fetchFeed() {
    const url = resolvedFeedUrl();
    try {
      const res = await fetch(url, { cache: 'no-store' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      if (Array.isArray(data)) {
        const baseUrl = import.meta.env.BASE_URL;
        items = data.map((it, idx) => {
          let icon = it.icon ?? '';
          if (icon.startsWith('/') && !icon.startsWith(baseUrl)) {
            icon = `${baseUrl.replace(/\/$/, '')}${icon}`;
          }
          return {
            id: it.id ?? it.title ?? idx,
            icon,
            href: it.href ?? '#',
            title: it.title ?? '',
            count: Number(it.count) || 0
          };
        });
      } else {
        items = [];
      }
    } catch (e) {
      console.error(`[PS Toolbar] Failed fetching notifications from ${url}`, e);
      items = [];
    }
  }

  onMount(() => {
    // Inject Shadow DOM styles if we have an element reference
    if (el) {
      const sr = el.getRootNode();
      if (sr instanceof ShadowRoot) {
        injectShadowCss(sr);
      }
    }

    fetchFeed();
    timer = setInterval(fetchFeed, pollInterval);
  });

  onDestroy(() => {
    clearInterval(timer);
  });
</script>

<div class="toolbar" role="toolbar" aria-label="Notifications toolbar">
  {#each items as item (item.id)}
    <div class="pds-app-action">
      <a 
        class="button-with-badge" 
        href={item.href} 
        title={item.title} 
        target="_top"
        aria-label={item.title}
      >
        {#if item.icon}
          <img src={item.icon} alt="" class="pds-icon" />
        {:else}
          <span class="icon-placeholder" aria-hidden="true">🔔</span>
        {/if}
        
        {#if item.count > 0}
          <span class="badge">{item.count}</span>
        {/if}
      </a>
    </div>
  {/each}
</div>

<style>
  /* 
    Styles are defined in injectShadowCss.ts for Shadow DOM compatibility.
    Svelte styles below will only work if customElement: true is NOT used,
    or during local development/testing where Shadow DOM might be bypassed.
  */
</style>
