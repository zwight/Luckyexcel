import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';
import { univerPlugin } from '@univerjs/vite-plugin';

// https://vite.dev/config/
export default defineConfig({
  plugins: [vue(), univerPlugin()],
  resolve: {
    alias: {
      '@zwight/luckyexcel': '/Users/lx/Desktop/project/Luckyexcel/dist/luckyexcel.esm.js',
    },
  },
});
