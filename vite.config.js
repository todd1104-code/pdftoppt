import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/pdftoppt/', // 對應您的 GitHub 專案名稱
  build: {
    outDir: 'dist', // 這次我們可以用回標準的 dist 了，因為不會有雙層鎖定問題
  }
})