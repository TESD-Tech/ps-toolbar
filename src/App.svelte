<svelte:options customElement="ps-toolbar" />
<script lang="ts">
  import Toolbar from './components/Toolbar.svelte';
  
  interface Props {
    feedUrl?: string;
    portal?: 'admin' | 'teachers' | 'guardian' | 'unknown';
    // Handle PowerSchool's various casing for userType/portal
    usertype?: 'admin' | 'teachers' | 'guardian';
    userType?: 'admin' | 'teachers' | 'guardian';
    'user-type'?: 'admin' | 'teachers' | 'guardian';
  }

  let { 
    feedUrl, 
    portal: portalAttr = 'unknown',
    usertype,
    userType,
    'user-type': userTypeHyphen
  } = $props<Props>();

  // Determine portal from various possible prop names
  const resolvedPortal = $derived.by(() => {
    const p = portalAttr || usertype || userType || userTypeHyphen || 'unknown';
    // Normalize 'teacher' to 'teachers' if needed
    if (p === 'teacher') return 'teachers';
    return p as 'admin' | 'teachers' | 'guardian' | 'unknown';
  });

</script>

<div style="display: contents;">
  <Toolbar {feedUrl} portal={resolvedPortal} />
</div>

<style>
  :host {
    display: inline-block;
    vertical-align: middle;
  }
</style>
