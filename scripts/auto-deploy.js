const fs = require('fs');
const { execSync } = require('child_process');

function run(cmd, desc) {
  console.log(`\n========================================`);
  console.log(`▶ ${desc}...`);
  console.log(`Command: ${cmd}`);
  console.log(`========================================`);
  try {
    const out = execSync(cmd, { stdio: 'inherit' });
    return true;
  } catch (err) {
    console.error(`❌ Gagal pada tahap: ${desc}`);
    return false;
  }
}

console.log('🚀 MEMULAI PROSES AUTO DEPLOY FJB OPERATIONS CONTROL SYSTEM');

// 1. Sync index.html ke index.txt
console.log('\n[1/4] Menyinkronkan index.html -> index.txt...');
try {
  fs.copyFileSync('index.html', 'index.txt');
  console.log('✅ File index.html dan index.txt tersinkronisasi.');
} catch (e) {
  console.error('❌ Gagal menyalin file:', e.message);
}

// 2. Clasp Push ke Google Apps Script
console.log('\n[2/4] Mengirim kode ke Google Apps Script...');
run('npx @google/clasp push --force', 'Clasp Push ke Apps Script');

// 3. Clasp Deploy ke Deployment Web App
console.log('\n[3/4] Menerapkan versi Web App di Google Apps Script...');
run('npx @google/clasp version "Auto-deploy update"', 'Buat Versi Baru Apps Script');
run('npx @google/clasp deploy -i AKfycbwFHFYH4trkXYzrLR6Fl0sLH6fF2gToTBJJb9BDcjGLPcPqqAXb5vnbooaYw5_-QrED -d "Auto-deployed Web App"', 'Perbarui Deployment Web App');

// 4. Git Push ke GitHub (memicu Vercel Auto Deploy)
console.log('\n[4/4] Commit dan Push ke GitHub untuk Auto Deploy Vercel...');
try {
  execSync('git add .');
  try {
    execSync('git diff --cached --quiet');
    console.log('ℹ️ Tidak ada perubahan file baru untuk di-commit.');
  } catch (diffErr) {
    const now = new Date().toISOString();
    execSync(`git commit -m "chore(deploy): auto-deploy update ${now}"`);
    console.log('✅ Perubahan berhasil di-commit.');
  }
  execSync('git push origin main', { stdio: 'inherit' });
  console.log('✅ Git push berhasil! Vercel otomatis mendeploy aplikasi.');
} catch (e) {
  console.error('❌ Terjadi kendala saat git push:', e.message);
}

console.log('\n🎉 PROSES AUTO DEPLOY SELESAI!');
console.log('👉 WebApp Vercel : https://fjb01-operation-001.vercel.app');
console.log('👉 Google Script : https://script.google.com/macros/s/AKfycbwFHFYH4trkXYzrLR6Fl0sLH6fF2gToTBJJb9BDcjGLPcPqqAXb5vnbooaYw5_-QrED/exec');
