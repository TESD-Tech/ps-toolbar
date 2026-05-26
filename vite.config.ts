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

export default defineConfig({
  plugins: [
    svelte({
      compilerOptions: { customElement: true, dev: true },
      emitCss: false,
    }),
    mockPsApi(),
  ],
  base: `/${projectName}/`,
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
