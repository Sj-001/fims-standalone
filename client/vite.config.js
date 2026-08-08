import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// During local development, the Vite dev server (port 5173) proxies /api/* to the Express backend
// (port 3000) so the frontend can call relative paths like '/api/extract' in both dev and
// production — in production the backend serves the built frontend directly, so there's no cross
// -origin request at all and no proxy is needed there.
export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      '/api': {
        target: 'http://localhost:3000',
        changeOrigin: true,
      },
    },
  },
});
