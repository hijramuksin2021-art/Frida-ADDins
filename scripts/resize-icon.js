// Script untuk resize ikon baru ke berbagai ukuran yang dibutuhkan add-in
// Jalankan: node scripts/resize-icon.js

const fs = require('fs');
const path = require('path');

// Cek apakah sharp tersedia
let sharp;
try {
  sharp = require('sharp');
} catch (e) {
  console.log('Sharp tidak tersedia. Menggunakan metode alternatif...');
  console.log('Silakan install sharp: npm install sharp --save-dev');
  console.log('Atau resize manual menggunakan tool lain.');
  process.exit(1);
}

const sizes = [
  { name: 'icon-16.png', size: 16 },
  { name: 'icon-32.png', size: 32 },
  { name: 'icon-80.png', size: 80 },
  { name: 'frida-logo.png', size: 40 }
];

const inputPath = path.join(__dirname, '..', 'assets', 'new-icon.png');
const outputDir = path.join(__dirname, '..', 'assets');

async function resizeIcons() {
  for (const { name, size } of sizes) {
    const outputPath = path.join(outputDir, name);
    await sharp(inputPath)
      .resize(size, size, {
        fit: 'contain',
        background: { r: 0, g: 0, b: 0, alpha: 0 }
      })
      .png()
      .toFile(outputPath);
    console.log(`✓ Created ${name} (${size}x${size})`);
  }
  console.log('\n✅ Semua ikon berhasil di-resize!');
}

resizeIcons().catch(console.error);
