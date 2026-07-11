const fs = require('fs');

async function test() {
  try {
    console.log('Memulai verifikasi API backend...');

    // 1. Health Check
    const healthRes = await fetch('http://localhost:5000/api/health');
    console.log('1. Health Check:', await healthRes.json());

    // 2. Upload File Excel
    console.log('2. Mengunggah file sample_data.xlsx...');
    const fileBuffer = fs.readFileSync('../sample_data.xlsx');
    const blob = new Blob([fileBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const formData = new FormData();
    formData.append('file', blob, 'sample_data.xlsx');

    const uploadRes = await fetch('http://localhost:5000/api/upload', {
      method: 'POST',
      body: formData
    });
    const uploadData = await uploadRes.json();
    console.log('   Hasil Unggah:', uploadData);

    // 3. Daftar File
    const filesRes = await fetch('http://localhost:5000/api/files');
    console.log('3. Daftar File Terdaftar:', await filesRes.json());

    // 4. Pencarian "Keychron"
    console.log('4. Mencari kata kunci "Keychron"...');
    const searchRes1 = await fetch('http://localhost:5000/api/search?q=Keychron');
    const searchData1 = await searchRes1.json();
    console.log('   Hasil pencarian Keychron:', JSON.stringify(searchData1, null, 2));

    // 5. Pencarian "Asus"
    console.log('5. Mencari kata kunci "Asus"...');
    const searchRes2 = await fetch('http://localhost:5000/api/search?q=Asus');
    const searchData2 = await searchRes2.json();
    console.log('   Hasil pencarian Asus:', JSON.stringify(searchData2, null, 2));

    // 6. Riwayat Log Aktivitas
    console.log('6. Riwayat Log Aktivitas...');
    const logsRes = await fetch('http://localhost:5000/api/logs');
    console.log('   Daftar Log:', await logsRes.json());

  } catch (err) {
    console.error('Tes API gagal:', err);
  }
}

test();
