import react from '@vitejs/plugin-react'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  base: './',

  plugins: [react()],

  test: {
    environment: 'jsdom',
  },

  build: {
    outDir: 'dist',
    emptyOutDir: true,

    rollupOptions: {
      output: {
        manualChunks: undefined,
      },
    },
  },

  server: {
    port: 5173,
    strictPort: true,
  },

  preview: {
    port: 4173,
    strictPort: true,
  },
})