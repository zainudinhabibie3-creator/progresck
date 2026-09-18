/**
 * ============================================================================
 * APLIKASI WEB MONITORING PROGRES FISIK & KEUANGAN PEKERJAAN
 * File: Code.gs (Server-Side Logic)
 * Engine: Google Apps Script V8
 * ============================================================================
 */

// Global Config & Spreadsheet Connection
function getDb() {
  const prop = PropertiesService.getScriptProperties();
  const ssId = prop.getProperty('SPREADSHEET_ID');
  if (ssId) {
    return SpreadsheetApp.openById(ssId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Main Web App Handler
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Monitoring Progres Pekerjaan')

    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Utility: Password Hashing SHA-256
 */
function hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return rawHash.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * Utility: Generate Unique UUID
 */
function generateUUID(prefix = 'ID') {
  const dateStr = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd');
  const rand = Math.floor(1000 + Math.random() * 9000);
  return `${prefix}-${dateStr}-${rand}`;
}

/**
 * Inisialisasi Database Spreadsheet Otomatis
 */
function initializeDatabase() {
  const ss = getDb();
  const sheetsDef = [
    {
      name: 'USERS',
      headers: ['user_id', 'username', 'password_hash', 'nama', 'email', 'role', 'status', 'unit_kerja', 'created_at', 'updated_at', 'last_login']
    },
    {
      name: 'PEKERJAAN',
      headers: ['pekerjaan_id', 'kode_pekerjaan', 'nama_pekerjaan', 'paket_pekerjaan', 'penyedia', 'nomor_kontrak', 'tanggal_kontrak', 'nilai_kontrak', 'tanggal_mulai', 'masa_pelaksanaan_hari', 'tanggal_selesai', 'status', 'progres_fisik_terakhir', 'progres_keuangan_terakhir', 'target_fisik', 'deviasi_fisik', 'sisa_hari', 'persentase_waktu_terpakai', 'status_risiko', 'keterangan', 'created_at', 'created_by', 'updated_at', 'updated_by']
    },
    {
      name: 'PROGRES_FISIK',
      headers: ['progres_id', 'pekerjaan_id', 'tanggal_monitoring', 'target_fisik', 'realisasi_fisik', 'deviasi_fisik', 'keterangan', 'url_dokumen', 'created_at', 'created_by']
    },
    {
      name: 'PROGRES_KEUANGAN',
      headers: ['keuangan_id', 'pekerjaan_id', 'tanggal_pembayaran', 'nomor_pembayaran', 'nilai_pembayaran', 'persentase_keuangan', 'keterangan', 'url_dokumen', 'created_at', 'created_by']
    },
    {
      name: 'ADENDUM',
      headers: ['adendum_id', 'pekerjaan_id', 'nomor_adendum', 'tanggal_adendum', 'jenis_adendum', 'nilai_sebelum', 'nilai_setelah', 'tanggal_selesai_sebelum', 'tanggal_selesai_setelah', 'tambah_hari', 'keterangan', 'url_dokumen', 'created_at', 'created_by']
    },
    {
      name: 'NOTIFIKASI',
      headers: ['notif_id', 'pekerjaan_id', 'judul', 'pesan', 'kategori', 'tingkat_risiko', 'status_baca', 'created_at']
    },
    {
      name: 'LOG_AKTIVITAS',
      headers: ['log_id', 'timestamp', 'user_id', 'username', 'action', 'module', 'record_id', 'description']
    },
    {
      name: 'SETTINGS',
      headers: ['key', 'value', 'description']
    }
  ];

  sheetsDef.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
    }
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers])
        .setFontWeight('bold')
        .setBackground('#1a1d29')
        .setFontColor('#00f2fe');
      sheet.setFrozenRows(1);
    }
  });

  // Default Settings
  const setSheet = ss.getSheetByName('SETTINGS');
  if (setSheet.getLastRow() <= 1) {
    const defaultSettings = [
      ['APP_NAME', 'Monitoring Progres Pekerjaan', 'Nama Aplikasi Utama'],
      ['THRESHOLD_DEVIASI_WASPADA', '-5', 'Batas Deviasi Fisik Waspada (%)'],
      ['THRESHOLD_DEVIASI_KRITIS', '-10', 'Batas Deviasi Fisik Kritis (%)'],
      ['THRESHOLD_GAP_KEUANGAN', '15', 'Selisih Maksimal Keuangan vs Fisik (%)'],
      ['PUBLIC_ACCESS', 'TRUE', 'Status Akses Dashboard Publik']
    ];
    setSheet.getRange(2, 1, defaultSettings.length, 3).setValues(defaultSettings);
  }

  return { success: true, message: 'Database dan Sheet berhasil diinisialisasi.' };
}

/**
 * Seed Demo Data Initial
 */
function seedDemoData() {
  initializeDatabase();
  const ss = getDb();

  // Create Default Admin User
  const userSheet = ss.getSheetByName('USERS');
  const userValues = userSheet.getDataRange().getValues();
  if (userValues.length <= 1) {
    const defaultAdmin = [
      generateUUID('USR'),
      'admin',
      hashPassword('admin123'),
      'Administrator Utama',
      'admin@instansi.go.id',
      'ADMIN',
      'AKTIF',
      'Dinas Pekerjaan Umum',
      new Date(),
      new Date(),
      ''
    ];
    const defaultUser = [
      generateUUID('USR'),
      'user',
      hashPassword('user123'),
      'Pengawas Lapangan',
      'user@instansi.go.id',
      'USER',
      'AKTIF',
      'Bidang Cipta Karya',
      new Date(),
      new Date(),
      ''
    ];
    userSheet.getRange(2, 1, 2, defaultAdmin.length).setValues([defaultAdmin, defaultUser]);
  }

  // Create Demo Projects
  const pkjSheet = ss.getSheetByName('PEKERJAAN');
  if (pkjSheet.getLastRow() <= 1) {
    const today = new Date();
    const startDate1 = new Date(today.getTime() - (30 * 24 * 60 * 60 * 1000));
    const endDate1 = new Date(startDate1.getTime() + (120 * 24 * 60 * 60 * 1000));

    const demo1 = [
      'PKJ-2026-0001',
      'BM-01',
      'Pembangunan Jembatan Utama Kota',
      'Bina Marga',
      'PT Konstruksi Mandiri Jaya',
      '602/KONTRAK/BM/2026',
      startDate1,
      2500000000,
      startDate1,
      120,
      endDate1,
      'BERJALAN',
      35,
      40,
      40,
      -5,
      90,
      25,
      'PERHATIAN',
      'Pekerjaan berjalan dengan deviasi kecil akibat cuaca.',
      new Date(),
      'admin',
      new Date(),
      'admin'
    ];

    const demo2 = [
      'PKJ-2026-0002',
      'CK-02',
      'Rehabilitasi Gedung Olahraga',
      'Cipta Karya',
      'CV Karya Bersama',
      '603/KONTRAK/CK/2026',
      startDate1,
      850000000,
      startDate1,
      90,
      new Date(startDate1.getTime() + (90 * 24 * 60 * 60 * 1000)),
      'BERJALAN',
      75,
      60,
      70,
      5,
      60,
      33,
      'AMAN',
      'Pekerjaan melampaui target fisik.',
      new Date(),
      'admin',
      new Date(),
      'admin'
    ];

    pkjSheet.getRange(2, 1, 2, demo1.length).setValues([demo1, demo2]);
  }

  return { success: true, message: 'Data demo dan akun default (admin/admin123, user/user123) berhasil dibuat!' };
}

/**
 * Log Helper
 */
function logActivity(userId, username, action, module, recordId, description) {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName('LOG_AKTIVITAS');
    sheet.appendRow([
      generateUUID('LOG'),
      new Date(),
      userId || '',
      username || 'ANONYMOUS',
      action,
      module,
      recordId || '',
      description
    ]);
  } catch (err) {
    Logger.log('Log Error: ' + err.toString());
  }
}

/**
 * Authentication Handler
 */
function loginUser(username, password) {
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName('USERS');
    const data = sheet.getDataRange().getValues();
    const passHash = hashPassword(password);

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username && data[i][2] === passHash) {
        if (data[i][6] !== 'AKTIF') {
          return { success: false, message: 'Akun Anda tidak aktif. Hubungi Admin.' };
        }
        
        // Update last login
        sheet.getRange(i + 1, 11).setValue(new Date());

        const userData = {
          user_id: data[i][0],
          username: data[i][1],
          nama: data[i][3],
          email: data[i][4],
          role: data[i][5],
          unit_kerja: data[i][7]
        };

        logActivity(userData.user_id, userData.username, 'LOGIN', 'AUTH', '', 'User berhasil login');
        return { success: true, user: userData, token: Utilities.base64Encode(JSON.stringify(userData)) };
      }
    }
    return { success: false, message: 'Username atau password salah!' };
  } catch (err) {
    return { success: false, message: 'Error Server: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Dynamic Risk & Status Engine Calculation
 */
function calculateRiskAndStatus(tanggalMulai, masaPelaksanaanHari, targetFisik, realisasiFisik, realisasiKeuangan, adendumHari = 0) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const start = new Date(tanggalMulai);
  start.setHours(0, 0, 0, 0);

  const totalHari = Number(masaPelaksanaanHari) + Number(adendumHari);
  const end = new Date(start.getTime() + (totalHari * 24 * 60 * 60 * 1000));
  end.setHours(0, 0, 0, 0);

  const diffTime = end.getTime() - today.getTime();
  const sisaHari = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  const totalDuration = end.getTime() - start.getTime();
  const elapsed = today.getTime() - start.getTime();
  let persentaseWaktu = totalDuration > 0 ? Math.min(100, Math.max(0, (elapsed / totalDuration) * 100)) : 0;

  const deviasi = Number(realisasiFisik) - Number(targetFisik);

  let status = 'BERJALAN';
  if (today < start) {
    status = 'BELUM MULAI';
  } else if (realisasiFisik >= 100) {
    status = 'SELESAI';
  } else if (sisaHari < 0) {
    status = 'TERLAMBAT';
  }

  // Risk Score Engine
  let riskLevel = 'AMAN';
  const gapKeuanganFisik = Number(realisasiKeuangan) - Number(realisasiFisik);

  if (sisaHari < 0 || deviasi <= -10 || gapKeuanganFisik > 20) {
    riskLevel = 'RISIKO TINGGI';
  } else if ((sisaHari <= 14 && sisaHari > 0) || (deviasi <= -5 && deviasi > -10) || gapKeuanganFisik > 10) {
    riskLevel = 'WASPADA';
  } else if ((sisaHari <= 30 && sisaHari > 14) || (deviasi < 0 && deviasi > -5)) {
    riskLevel = 'PERHATIAN';
  }

  return {
    tanggal_selesai: Utilities.formatDate(end, 'GMT+7', 'yyyy-MM-dd'),
    sisa_hari: sisaHari,
    persentase_waktu: Math.round(persentaseWaktu * 10) / 10,
    deviasi: Math.round(deviasi * 10) / 10,
    status: status,
    status_risiko: riskLevel
  };
}

/**
 * Public Dashboard Data fetcher
 */
function getPublicDashboardData() {
  const ss = getDb();
  const sheet = ss.getSheetByName('PEKERJAAN');
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return {
      stats: { total: 0, berjalan: 0, selesai: 0, terlambat: 0, total_nilai: 0 },
      pekerjaan: []
    };
  }

  let total = 0, berjalan = 0, selesai = 0, terlambat = 0, totalNilai = 0;
  const listPekerjaan = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    total++;
    const nilai = Number(row[7]) || 0;
    totalNilai += nilai;

    const st = row[11];
    if (st === 'BERJALAN') berjalan++;
    else if (st === 'SELESAI') selesai++;
    else if (st === 'TERLAMBAT') terlambat++;

    // Tanggal formatting
    const tglMulai = row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+7', 'dd/MM/yyyy') : row[8];
    const tglSelesai = row[10] instanceof Date ? Utilities.formatDate(row[10], 'GMT+7', 'dd/MM/yyyy') : row[10];

    listPekerjaan.push({
      pekerjaan_id: row[0],
      kode_pekerjaan: row[1],
      nama_pekerjaan: row[2],
      paket_pekerjaan: row[3],
      penyedia: row[4],
      nilai_kontrak: nilai,
      tanggal_mulai: tglMulai,
      tanggal_selesai: tglSelesai,
      status: st,
      progres_fisik: row[12],
      progres_keuangan: row[13],
      deviasi: row[15],
      sisa_hari: row[16],
      status_risiko: row[18]
    });
  }

  return {
    stats: { total, berjalan, selesai, terlambat, total_nilai: totalNilai },
    pekerjaan: listPekerjaan
  };
}

/**
 * Full Dashboard Data (Authenticated)
 */
function getAdminDashboardData() {
  const publicData = getPublicDashboardData();
  const ss = getDb();
  
  // High Risk Jobs
  const highRisk = publicData.pekerjaan.filter(p => p.status_risiko === 'RISIKO TINGGI' || p.status_risiko === 'WASPADA');
  
  // Recent Notifications
  const notifSheet = ss.getSheetByName('NOTIFIKASI');
  const notifData = notifSheet.getDataRange().getValues();
  const notifikasi = [];
  for (let i = notifData.length - 1; i >= 1 && notifikasi.length < 5; i--) {
    notifikasi.push({
      notif_id: notifData[i][0],
      judul: notifData[i][2],
      pesan: notifData[i][3],
      kategori: notifData[i][4],
      tingkat_risiko: notifData[i][5],
      created_at: notifData[i][7] instanceof Date ? Utilities.formatDate(notifData[i][7], 'GMT+7', 'dd/MM/yyyy HH:mm') : notifData[i][7]
    });
  }

  return {
    stats: publicData.stats,
    pekerjaan: publicData.pekerjaan,
    risk_list: highRisk,
    notifications: notifikasi
  };
}

/**
 * CRUD Pekerjaan
 */
function createPekerjaan(formData, username) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName('PEKERJAAN');
    const pekerjaanId = generateUUID('PKJ');

    const calc = calculateRiskAndStatus(
      formData.tanggal_mulai,
      formData.masa_pelaksanaan_hari,
      0,
      0,
      0
    );

    const newRow = [
      pekerjaanId,
      formData.kode_pekerjaan,
      formData.nama_pekerjaan,
      formData.paket_pekerjaan,
      formData.penyedia,
      formData.nomor_kontrak,
      new Date(formData.tanggal_kontrak),
      Number(formData.nilai_kontrak),
      new Date(formData.tanggal_mulai),
      Number(formData.masa_pelaksanaan_hari),
      new Date(calc.tanggal_selesai),
      calc.status,
      0, // progres fisik
      0, // progres keuangan
      0, // target fisik
      0, // deviasi
      calc.sisa_hari,
      calc.persentase_waktu,
      calc.status_risiko,
      formData.keterangan || '',
      new Date(),
      username,
      new Date(),
      username
    ];

    sheet.appendRow(newRow);
    logActivity(username, username, 'CREATE', 'PEKERJAAN', pekerjaanId, `Menambah pekerjaan: ${formData.nama_pekerjaan}`);
    return { success: true, message: 'Pekerjaan berhasil disimpan!', pekerjaan_id: pekerjaanId };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Get Detail Pekerjaan Complete
 */
function getDetailPekerjaan(pekerjaanId) {
  const ss = getDb();
  
  // Pekerjaan Master Data
  const sheetPkj = ss.getSheetByName('PEKERJAAN');
  const dataPkj = sheetPkj.getDataRange().getValues();
  let master = null;

  for (let i = 1; i < dataPkj.length; i++) {
    if (dataPkj[i][0] === pekerjaanId) {
      const row = dataPkj[i];
      master = {
        pekerjaan_id: row[0],
        kode_pekerjaan: row[1],
        nama_pekerjaan: row[2],
        paket_pekerjaan: row[3],
        penyedia: row[4],
        nomor_kontrak: row[5],
        tanggal_kontrak: row[6] instanceof Date ? Utilities.formatDate(row[6], 'GMT+7', 'yyyy-MM-dd') : row[6],
        nilai_kontrak: row[7],
        tanggal_mulai: row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+7', 'yyyy-MM-dd') : row[8],
        masa_pelaksanaan_hari: row[9],
        tanggal_selesai: row[10] instanceof Date ? Utilities.formatDate(row[10], 'GMT+7', 'yyyy-MM-dd') : row[10],
        status: row[11],
        progres_fisik_terakhir: row[12],
        progres_keuangan_terakhir: row[13],
        target_fisik: row[14],
        deviasi_fisik: row[15],
        sisa_hari: row[16],
        persentase_waktu_terpakai: row[17],
        status_risiko: row[18],
        keterangan: row[19]
      };
      break;
    }
  }

  if (!master) return { success: false, message: 'Data pekerjaan tidak ditemukan.' };

  // Histori Fisik
  const sheetFisik = ss.getSheetByName('PROGRES_FISIK');
  const dataFisik = sheetFisik.getDataRange().getValues();
  const listFisik = [];
  for (let i = 1; i < dataFisik.length; i++) {
    if (dataFisik[i][1] === pekerjaanId) {
      listFisik.push({
        progres_id: dataFisik[i][0],
        tanggal: dataFisik[i][2] instanceof Date ? Utilities.formatDate(dataFisik[i][2], 'GMT+7', 'dd/MM/yyyy') : dataFisik[i][2],
        target: dataFisik[i][3],
        realisasi: dataFisik[i][4],
        deviasi: dataFisik[i][5],
        keterangan: dataFisik[i][6],
        url_dokumen: dataFisik[i][7]
      });
    }
  }

  // Histori Keuangan
  const sheetKeu = ss.getSheetByName('PROGRES_KEUANGAN');
  const dataKeu = sheetKeu.getDataRange().getValues();
  const listKeuangan = [];
  for (let i = 1; i < dataKeu.length; i++) {
    if (dataKeu[i][1] === pekerjaanId) {
      listKeuangan.push({
        keuangan_id: dataKeu[i][0],
        tanggal: dataKeu[i][2] instanceof Date ? Utilities.formatDate(dataKeu[i][2], 'GMT+7', 'dd/MM/yyyy') : dataKeu[i][2],
        nomor_pembayaran: dataKeu[i][3],
        nilai: dataKeu[i][4],
        persentase: dataKeu[i][5],
        keterangan: dataKeu[i][6],
        url_dokumen: dataKeu[i][7]
      });
    }
  }

  // Histori Adendum
  const sheetAd = ss.getSheetByName('ADENDUM');
  const dataAd = sheetAd.getDataRange().getValues();
  const listAdendum = [];
  for (let i = 1; i < dataAd.length; i++) {
    if (dataAd[i][1] === pekerjaanId) {
      listAdendum.push({
        adendum_id: dataAd[i][0],
        nomor_adendum: dataAd[i][2],
        tanggal: dataAd[i][3] instanceof Date ? Utilities.formatDate(dataAd[i][3], 'GMT+7', 'dd/MM/yyyy') : dataAd[i][3],
        jenis: dataAd[i][4],
        nilai_setelah: dataAd[i][6],
        tambah_hari: dataAd[i][9],
        keterangan: dataAd[i][10]
      });
    }
  }

  return {
    success: true,
    master: master,
    histori_fisik: listFisik,
    histori_keuangan: listKeuangan,
    histori_adendum: listAdendum
  };
}

/**
 * Save Input Progres Fisik
 */
function saveProgresFisik(formData, username) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = getDb();
    const sheetFisik = ss.getSheetByName('PROGRES_FISIK');
    const progresId = generateUUID('PRF');

    const deviasi = Number(formData.realisasi_fisik) - Number(formData.target_fisik);

    sheetFisik.appendRow([
      progresId,
      formData.pekerjaan_id,
      new Date(formData.tanggal_monitoring),
      Number(formData.target_fisik),
      Number(formData.realisasi_fisik),
      deviasi,
      formData.keterangan || '',
      formData.url_dokumen || '',
      new Date(),
      username
    ]);

    // Recalculate Master Pekerjaan
    recalculateMasterPekerjaan(formData.pekerjaan_id, username);

    logActivity(username, username, 'CREATE', 'PROGRES_FISIK', progresId, `Input progres fisik: ${formData.realisasi_fisik}%`);
    return { success: true, message: 'Progres fisik berhasil diperbarui!' };
  } catch (err) {
    return { success: false, message: 'Gagal update fisik: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Save Input Progres Keuangan
 */
function saveProgresKeuangan(formData, username) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = getDb();
    const sheetKeu = ss.getSheetByName('PROGRES_KEUANGAN');
    const keuanganId = generateUUID('PRK');

    sheetKeu.appendRow([
      keuanganId,
      formData.pekerjaan_id,
      new Date(formData.tanggal_pembayaran),
      formData.nomor_pembayaran,
      Number(formData.nilai_pembayaran),
      Number(formData.persentase_keuangan),
      formData.keterangan || '',
      formData.url_dokumen || '',
      new Date(),
      username
    ]);

    // Recalculate Master
    recalculateMasterPekerjaan(formData.pekerjaan_id, username);

    logActivity(username, username, 'CREATE', 'PROGRES_KEUANGAN', keuanganId, `Input progres keuangan: ${formData.persentase_keuangan}%`);
    return { success: true, message: 'Progres keuangan berhasil disimpan!' };
  } catch (err) {
    return { success: false, message: 'Gagal update keuangan: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Save Adendum
 */
function saveAdendum(formData, username) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = getDb();
    const sheetAd = ss.getSheetByName('ADENDUM');
    const adendumId = generateUUID('ADN');

    sheetAd.appendRow([
      adendumId,
      formData.pekerjaan_id,
      formData.nomor_adendum,
      new Date(formData.tanggal_adendum),
      formData.jenis_adendum,
      Number(formData.nilai_sebelum),
      Number(formData.nilai_setelah),
      new Date(formData.tanggal_selesai_sebelum),
      new Date(formData.tanggal_selesai_setelah),
      Number(formData.tambah_hari || 0),
      formData.keterangan || '',
      formData.url_dokumen || '',
      new Date(),
      username
    ]);

    // Update Master Pekerjaan Nilai Kontrak / Selesai
    const sheetPkj = ss.getSheetByName('PEKERJAAN');
    const dataPkj = sheetPkj.getDataRange().getValues();

    for (let i = 1; i < dataPkj.length; i++) {
      if (dataPkj[i][0] === formData.pekerjaan_id) {
        if (Number(formData.nilai_setelah) > 0) {
          sheetPkj.getRange(i + 1, 8).setValue(Number(formData.nilai_setelah));
        }
        if (Number(formData.tambah_hari) > 0) {
          const newDurasi = Number(dataPkj[i][9]) + Number(formData.tambah_hari);
          sheetPkj.getRange(i + 1, 10).setValue(newDurasi);
        }
        break;
      }
    }

    recalculateMasterPekerjaan(formData.pekerjaan_id, username);

    logActivity(username, username, 'CREATE', 'ADENDUM', adendumId, `Tambah Adendum: ${formData.nomor_adendum}`);
    return { success: true, message: 'Adendum berhasil ditambahkan!' };
  } catch (err) {
    return { success: false, message: 'Gagal simpan adendum: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper Update Master Sheet
 */
function recalculateMasterPekerjaan(pekerjaanId, username) {
  const ss = getDb();
  
  // Get Latest Fisik
  const sheetFisik = ss.getSheetByName('PROGRES_FISIK');
  const dataFisik = sheetFisik.getDataRange().getValues();
  let latestTarget = 0, latestFisik = 0, latestDeviasi = 0;
  for (let i = dataFisik.length - 1; i >= 1; i--) {
    if (dataFisik[i][1] === pekerjaanId) {
      latestTarget = dataFisik[i][3];
      latestFisik = dataFisik[i][4];
      latestDeviasi = dataFisik[i][5];
      break;
    }
  }

  // Get Latest Keuangan
  const sheetKeu = ss.getSheetByName('PROGRES_KEUANGAN');
  const dataKeu = sheetKeu.getDataRange().getValues();
  let latestKeuangan = 0;
  for (let i = dataKeu.length - 1; i >= 1; i--) {
    if (dataKeu[i][1] === pekerjaanId) {
      latestKeuangan = dataKeu[i][5];
      break;
    }
  }

  // Update Master Row
  const sheetPkj = ss.getSheetByName('PEKERJAAN');
  const dataPkj = sheetPkj.getDataRange().getValues();

  for (let i = 1; i < dataPkj.length; i++) {
    if (dataPkj[i][0] === pekerjaanId) {
      const row = dataPkj[i];
      const calc = calculateRiskAndStatus(row[8], row[9], latestTarget, latestFisik, latestKeuangan);

      sheetPkj.getRange(i + 1, 11).setValue(new Date(calc.tanggal_selesai)); // tgl selesai
      sheetPkj.getRange(i + 1, 12).setValue(calc.status);
      sheetPkj.getRange(i + 1, 13).setValue(latestFisik);
      sheetPkj.getRange(i + 1, 14).setValue(latestKeuangan);
      sheetPkj.getRange(i + 1, 15).setValue(latestTarget);
      sheetPkj.getRange(i + 1, 16).setValue(calc.deviasi);
      sheetPkj.getRange(i + 1, 17).setValue(calc.sisa_hari);
      sheetPkj.getRange(i + 1, 18).setValue(calc.persentase_waktu);
      sheetPkj.getRange(i + 1, 19).setValue(calc.status_risiko);
      sheetPkj.getRange(i + 1, 23).setValue(new Date());
      sheetPkj.getRange(i + 1, 24).setValue(username);
      break;
    }
  }
}

/**
 * User Management Services (ADMIN ONLY)
 */
function getUsersList() {
  const ss = getDb();
  const sheet = ss.getSheetByName('USERS');
  const data = sheet.getDataRange().getValues();

  const users = [];
  for (let i = 1; i < data.length; i++) {
    users.push({
      user_id: data[i][0],
      username: data[i][1],
      nama: data[i][3],
      email: data[i][4],
      role: data[i][5],
      status: data[i][6],
      unit_kerja: data[i][7],
      last_login: data[i][10] instanceof Date ? Utilities.formatDate(data[i][10], 'GMT+7', 'dd/MM/yyyy HH:mm') : '-'
    });
  }
  return users;
}

function createUser(userData, adminUser) {
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName('USERS');
    const data = sheet.getDataRange().getValues();

    // Check duplicate
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === userData.username) {
        return { success: false, message: 'Username sudah digunakan!' };
      }
    }

    const userId = generateUUID('USR');
    sheet.appendRow([
      userId,
      userData.username,
      hashPassword(userData.password),
      userData.nama,
      userData.email,
      userData.role,
      'AKTIF',
      userData.unit_kerja,
      new Date(),
      new Date(),
      ''
    ]);

    logActivity(adminUser, adminUser, 'CREATE', 'USERS', userId, `Tambah user baru: ${userData.username}`);
    return { success: true, message: 'User berhasil dibuat!' };
  } catch (err) {
    return { success: false, message: 'Gagal membuat user: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function toggleUserStatus(userId, adminUser) {
  const ss = getDb();
  const sheet = ss.getSheetByName('USERS');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      const newStatus = data[i][6] === 'AKTIF' ? 'NONAKTIF' : 'AKTIF';
      sheet.getRange(i + 1, 7).setValue(newStatus);
      logActivity(adminUser, adminUser, 'UPDATE', 'USERS', userId, `Ubah status user ke ${newStatus}`);
      return { success: true, message: `Status user diubah menjadi ${newStatus}` };
    }
  }
  return { success: false, message: 'User tidak ditemukan.' };
}

/**
 * Automation Daily Monitoring Trigger
 */
function dailyMonitoringTrigger() {
  const ss = getDb();
  const sheetPkj = ss.getSheetByName('PEKERJAAN');
  const dataPkj = sheetPkj.getDataRange().getValues();
  const notifSheet = ss.getSheetByName('NOTIFIKASI');

  const today = new Date();

  for (let i = 1; i < dataPkj.length; i++) {
    const row = dataPkj[i];
    const pekerjaanId = row[0];
    const namaPekerjaan = row[2];
    const sisaHari = Number(row[16]);
    const deviasi = Number(row[15]);
    const statusRisiko = row[18];

    // Notification Trigger Cases
    if (sisaHari <= 7 && sisaHari > 0 && row[11] === 'BERJALAN') {
      notifSheet.appendRow([
        generateUUID('NTF'),
        pekerjaanId,
        `🔴 H-${sisaHari} BATAS AKHIR PEKERJAAN`,
        `Pekerjaan ${namaPekerjaan} tersisa ${sisaHari} hari lagi dengan progres ${row[12]}% (Target ${row[14]}%).`,
        'DEADLINE',
        'KRITIS',
        'UNREAD',
        today
      ]);
    } else if (deviasi <= -10 && row[11] === 'BERJALAN') {
      notifSheet.appendRow([
        generateUUID('NTF'),
        pekerjaanId,
        `⚠️ DEVIASI SANGAT TINGGI (${deviasi}%)`,
        `Pekerjaan ${namaPekerjaan} mengalami keterlambatan fisik deviasi ${deviasi}%. Perlu pembuktian evaluasi KONTRAK (SCM).`,
        'DEVIASI',
        'KRITIS',
        'UNREAD',
        today
      ]);
    }
  }
}
