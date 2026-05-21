<script>
  import { onMount, onDestroy } from 'svelte';
  export let feedUrl = '/notifications.json';
  export let pollInterval = 30000; // ms

  let items = [];
  let timer;

  async function fetchFeed() {
    try {
      const res = await fetch(feedUrl, { cache: 'no-store' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      if (Array.isArray(data)) {
        items = data.map((it, idx) => ({
          id: it.id ?? it.title ?? idx,
          icon: it.icon ?? '',
          href: it.href ?? '#',
          title: it.title ?? '',
          count: Number(it.count) || 0
        }));
      } else {
        items = [];
      }
    } catch (e) {
      console.error('Failed fetching notifications', e);
      items = [];
    }
  }

  onMount(() => {
    fetchFeed();
    timer = setInterval(fetchFeed, pollInterval);
  });

  onDestroy(() => {
    clearInterval(timer);
  });
</script>

<nav class="toolbar" role="toolbar" aria-label="Notifications toolbar">
  {#each items as item (item.id)}
    <a class="notif" href={item.href} title={item.title} target="_blank" rel="noopener noreferrer" aria-label={item.title}>
      {#if item.icon}
        <img src={item.icon} alt="" class="icon" />
      {:else}
        <span class="icon-placeholder" aria-hidden="true">🔔</span>
      {/if}
      {#if item.count > 0}
        <span class="badge">{item.count}</span>
      {/if}
    </a>
  {/each}
</nav>

<style>
  .toolbar{
    display:flex;
    gap:0.5rem;
    align-items:center;
    padding:0.5rem;
    background:#111827;
    color:#fff;
  }
  .notif{
    display:inline-flex;
    align-items:center;
    justify-content:center;
    width:40px;
    height:40px;
    border-radius:6px;
    position:relative;
    color:inherit;
    text-decoration:none;
  }
  .icon{ width:24px; height:24px; display:block }
  .badge{
    position:absolute;
    top:-6px;
    right:-6px;
    background:#ef4444;
    color:#fff;
    font-size:12px;
    padding:2px 6px;
    border-radius:999px;
    min-width:18px;
    text-align:center;
    box-shadow:0 1px 0 rgba(0,0,0,0.2);
  }
  .icon-placeholder { font-size:20px }
</style>
