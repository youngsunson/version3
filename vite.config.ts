import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';
import os from 'os';

// Check for certificates in multiple locations
function getCertificates() {
  const certLocations = [
    // Check for .crt and .key files (default office-addin-dev-certs format)
    {
      key: path.join(os.homedir(), '.office-addin-dev-certs', 'localhost.key'),
      cert: path.join(os.homedir(), '.office-addin-dev-certs', 'localhost.crt')
    },
    // Check for .pem files
    {
      key: path.join(os.homedir(), '.office-addin-dev-certs', 'localhost-key.pem'),
      cert: path.join(os.homedir(), '.office-addin-dev-certs', 'localhost.pem')
    },
    // Check in local certs folder
    {
      key: path.join(__dirname, 'certs', 'localhost.key'),
      cert: path.join(__dirname, 'certs', 'localhost.crt')
    },
    {
      key: path.join(__dirname, 'certs', 'localhost-key.pem'),
      cert: path.join(__dirname, 'certs', 'localhost.pem')
    }
  ];

  for (const location of certLocations) {
    if (fs.existsSync(location.key) && fs.existsSync(location.cert)) {
      console.log('✅ Found certificates at:', location.cert);
      return {
        key: fs.readFileSync(location.key),
        cert: fs.readFileSync(location.cert)
      };
    }
  }

  console.warn('⚠️ No certificates found. Please run: npx office-addin-dev-certs install --machine');
  return false;
}

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  base: '/bengali-spell-check/', // GitHub Pages base path
  root: '.',
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: {
        main: path.resolve(__dirname, 'index.html'),
        commands: path.resolve(__dirname, 'commands.html')
      }
    }
  },
  publicDir: 'public',
  server: {
    port: 3000,
    host: 'localhost',
    https: getCertificates() || undefined,
    headers: {
      'Access-Control-Allow-Origin': '*'
    }
  }
});
