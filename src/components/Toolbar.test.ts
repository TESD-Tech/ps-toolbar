import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render, screen, waitFor } from '@testing-library/svelte';
import Toolbar from './Toolbar.svelte';

// Mock fetch
const mockData = [
  {
    id: 'messages',
    icon: 'mail',
    href: '/admin/messages.html',
    title: 'Messages',
    description: 'View your school messages',
    count: 3
  }
];

global.fetch = vi.fn();

describe('Toolbar', () => {
  beforeEach(() => {
    vi.resetAllMocks();
  });

  it('renders notifications from feed with hover tooltip', async () => {
    (global.fetch as any).mockResolvedValue({
      ok: true,
      json: async () => mockData
    });

    render(Toolbar, { feedUrl: '/test.json' });

    await waitFor(() => {
      const link = screen.getByTitle('View your school messages');
      expect(link).toBeInTheDocument();
    });

    const badge = screen.getByText('3');
    expect(badge).toBeInTheDocument();
  });

  it('resolves feed URL based on portal', async () => {
    (global.fetch as any).mockResolvedValue({
      ok: true,
      json: async () => []
    });

    render(Toolbar, { portal: 'admin' });

    await waitFor(() => {
      expect(global.fetch).toHaveBeenCalledWith('/admin/ps-toolbar/notifications.json', expect.any(Object));
    });
  });

  it('handles empty feed', async () => {
    (global.fetch as any).mockResolvedValue({
      ok: true,
      json: async () => []
    });

    render(Toolbar, { feedUrl: '/test.json' });

    await waitFor(() => {
      const nav = screen.getByRole('toolbar');
      expect(nav.children.length).toBe(0);
    });
  });

  it('handles fetch error gracefully', async () => {
    (global.fetch as any).mockRejectedValue(new Error('Network error'));

    render(Toolbar, { feedUrl: '/test.json' });

    await waitFor(() => {
      const nav = screen.getByRole('toolbar');
      expect(nav.children.length).toBe(0);
    });
  });
});
