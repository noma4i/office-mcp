import { copyFileSync, mkdirSync, readdirSync, statSync, rmSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const srcDir = join(__dirname, '../src');
const distDir = join(__dirname, '../dist');

function copyRecursive(src, dest) {
  const stats = statSync(src);

  if (stats.isDirectory()) {
    mkdirSync(dest, { recursive: true });
    const entries = readdirSync(src);

    for (const entry of entries) {
      copyRecursive(join(src, entry), join(dest, entry));
    }
  } else if (stats.isFile()) {
    copyFileSync(src, dest);
  }
}

console.log('Building...');
rmSync(distDir, { recursive: true, force: true });
copyRecursive(srcDir, distDir);
console.log('Build complete!');
