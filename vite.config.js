import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/ATO_Travel_Diary/',
  server: { port: 5173, strictPort: true },
})
