const fs = require('fs');
const path = require('path');

const dirsToClean = [
  'dist',
  'release', 
  'temp',
  'lib',
  '.heft/build-cache',
  '.heft/webpack-cache'
];

console.log('Cleaning build artifacts...\n');

let cleanedCount = 0;
let skippedCount = 0;

for (const dir of dirsToClean) {
  const fullPath = path.join(process.cwd(), dir);
  try {
    if (fs.existsSync(fullPath)) {
      const stats = fs.statSync(fullPath);
      const sizeMB = stats.isDirectory() 
        ? getFolderSize(fullPath) / 1024 / 1024
        : stats.size / 1024 / 1024;
      
      fs.rmSync(fullPath, { recursive: true, force: true });
      console.log(`  ✓ Removed ${dir} (${sizeMB.toFixed(1)} MB)`);
      cleanedCount++;
    } else {
      console.log(`  - Skipped ${dir} (not found)`);
      skippedCount++;
    }
  } catch (error) {
    console.error(`  ✗ Error removing ${dir}:`, error.message);
  }
}

console.log(`\nCleaned ${cleanedCount} directories, skipped ${skippedCount}`);

function getFolderSize(folderPath) {
  let size = 0;
  try {
    const files = fs.readdirSync(folderPath);
    for (const file of files) {
      const filePath = path.join(folderPath, file);
      const stats = fs.statSync(filePath);
      if (stats.isDirectory()) {
        size += getFolderSize(filePath);
      } else {
        size += stats.size;
      }
    }
  } catch {
    // Ignore permission errors
  }
  return size;
}
