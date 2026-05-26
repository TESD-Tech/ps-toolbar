import { defineConfig } from 'vite'
import { svelte } from '@sveltejs/vite-plugin-svelte'
import path from 'path'
import fs from 'fs'
import { mockPsApi } from './vite-plugin-mock-ps-api'

const pkg = JSON.parse(fs.readFileSync('package.json', 'utf-8'))
const projectName = pkg.name 

function noEmptyChunks() {
  return {
    name: 'no-empty-chunks',
    generateBundle(_: unknown, bundle: Record<string, { type: string; code?: string }>) {
      for (const chunk of Object.values(bundle))
        if (chunk.type === 'chunk' && !chunk.code?.trim()) chunk.code = 'export{}'
    }
  }
}

// Gather all .svelte files in src/lib/components, src/lib root, and src root
function gatherSvelteEntries() {
  const entries = {};
  const componentDirs = [
    path.resolve(__dirname, 'src/lib/components'),
    path.resolve(__dirname, 'src/lib'),
    path.resolve(__dirname, 'src')
  ];
  for (const dir of componentDirs) {
    if (!fs.existsSync(dir)) continue;
    const files = fs.readdirSync(dir);
    for (const file of files) {
      if (file.endsWith('.svelte')) {
        const name = file.replace(/\.svelte$/, '').toLowerCase();
        entries[name] = path.resolve(dir, file);
      }
    }
  }
  return entries;
}

const componentEntries = gatherSvelteEntries();

export default defineConfig({
  plugins: [
    svelte({
      compilerOptions: { customElement: true, dev: true },
      emitCss: false,
    }),
    mockPsApi(),
  ],
  base: `/${projectName}/`,
  css: {
    modules: {
      localsConvention: 'camelCaseOnly'
    }
  },
  resolve: {
    alias: { '$lib': path.resolve(__dirname, 'src/lib') },
    conditions: ['browser']
  },
  build: {
    outDir: `dist/WEB_ROOT/${projectName}/`,
    rollupOptions: {
      input: {
        main: path.resolve(__dirname, 'index.html'),
        app: path.resolve(__dirname, 'src/main.ts'),
        admin: path.resolve(__dirname, 'src/admin.ts'),
        ...componentEntries
      },
      plugins: [noEmptyChunks()],
      output: {
        format: 'es',
        entryFileNames: '[name].js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name].[ext]',
      },
    },
  },
  test: {
    environment: 'jsdom',
    globals: true,
    setupFiles: ['./src/test/setup.ts'],
    exclude: ['**/node_modules/**', '**/dist/**', '**/ref/**'],
  },
})
