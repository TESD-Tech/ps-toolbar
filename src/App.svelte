<svelte:options customElement="ps-toolbar" />
<script lang="ts">
  import { onMount } from 'svelte';
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
  const resolvedPortal = $derived(() => {
    const p = portalAttr || usertype || userType || userTypeHyphen || 'unknown';
    // Normalize 'teacher' to 'teachers' if needed
    if (p === 'teacher') return 'teachers';
    return p as 'admin' | 'teachers' | 'guardian' | 'unknown';
  });

  // Reference to self for Shadow DOM access in child components
  let el = $state<HTMLElement>();

  onMount(() => {
    // el is already bound via bind:this
  });
</script>

<div bind:this={el} style="display: contents;">
  <Toolbar {feedUrl} portal={resolvedPortal()} {el} />
</div>

<style>
  :host {
    display: inline-block;
    vertical-align: middle;
  }
</style>
