<script>
  import Toolbar from './components/Toolbar.svelte';
  // Default: local dev feed; in production set VITE_NOTIF_FEED_URL
  const feedUrl = import.meta.env.VITE_NOTIF_FEED_URL ?? '/notifications.json';
</script>

<Toolbar {feedUrl} />

<main class="container">
  <h1>PowerSchool toolbar demo</h1>
  <p>Set VITE_NOTIF_FEED_URL to point to a production JSON feed.</p>
</main>

<style>
  .container { padding: 1rem; max-width: 980px; margin: 0 auto; }
</style>
