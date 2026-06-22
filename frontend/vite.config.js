import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      '/validate': 'https://qa-tool-1oh2.onrender.com',
      '/health'  : 'https://qa-tool-1oh2.onrender.com',
    }
  }
})
