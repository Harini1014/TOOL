import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      '/validate': 'https://tool-2-3w1t.onrender.com/,
      '/health': 'https://tool-3-vctq.onrender.com,
    }
  }
})
