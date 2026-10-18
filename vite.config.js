import { defineConfig } from 'vite';

const base = process.env.VITE_BASE || '/';

export default defineConfig({
  base,
  server: {
    host: '0.0.0.0',
    port: 3000,
    allowedHosts: true,
  },
  preview: {
    host: '0.0.0.0',
    port: 3000,
    allowedHosts: true,
  },
});
