import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // GitHub Pages URL will be: https://kbsreeganesh.github.io/caseflow/
  base: '/caseflow/',
})
