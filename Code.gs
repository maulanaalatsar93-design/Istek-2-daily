/**
 * SISTEM MONITORING INSPEKSI TEKNIK 2 (DAILY ACTIVITY)
 * FINAL OPTIMIZED VERSION - Precision durasi & fast loading
 */

const SS_ID = "1bi3-sHZX2ij6-Jaza8M1h0ldQMS6zo2IhutpfIAQOh0";
const GID = {
  BAGIAN: "26433367",
  PIC: "1846268070",
  STATUS: "687780390",
  PABRIK_AREA: "1368866809",
  PROGRAM: "390206398",
  CRITICALITY: "1890488090",
  EQUIPMENT: "597683604", 
  DAILY_NAME: "Daily_Activity" 
};

// Cache untuk Master Data
let MASTER_DATA_CACHE = null;
let CACHE_TIME = 0;
const CACHE_DURATION = 10 * 60 * 1000; // 10 menit

/**
 * Normalize number - support . dan , sebagai decimal separator
 */
function normalizeNumber(value) {
  if (typeof value === 'number') {
    return isNaN(value) ? 0 : parseFloat(value.toFixed(2));
  }
  if (!value) return 0;
  
  let str = String(value).trim();
  
  // Jika sudah format "1.5 jam" atau "2.5 h", extract angka saja
  str = str.replace(/\s*(jam|h|hour|hours)?$/i, '');
  
  // Ganti koma dengan titik untuk decimal
  str = str.replace(',', '.');
  
  // Remove semua karakter selain angka dan titik
  str = str.replace(/[^0-9.]/g, '');
  
  // Handle multiple dots - keep hanya first dot
  const parts = str.split('.');
  if (parts.length > 2) {
    str = parts[0] + '.' + parts.slice(1).join('');
  }
  
  const num = parseFloat(str);
  if (isNaN(num)) return 0;
  
  // Validasi range waktu (0 - 24 jam)
  if (num < 0 || num > 24) return 0;
  
  return parseFloat(num.toFixed(2));
}

/**
 * Extract durasi dari row dengan multiple fallback strategies
 * Struktur: M:Lama(12), N:Start(13), O:Finish(14)
 */
function extractDuration(row) {
  let jam = 0;
  
  // STRATEGY 1: Kolom M (Index 12) - Lama Pekerjaan (PRIMARY)
  if (row[12]) {
    jam = normalizeNumber(row[12]);
    if (jam > 0) {
      return jam;
    }
  }
  
  // STRATEGY 2: Calculate dari N (Start, Index 13) dan O (Finish, Index 14)
  if (row[13] && row[14]) {
    try {
      const startStr = String(row[13]).trim();
      const endStr = String(row[14]).trim();
      
      // Check format HH:MM
      if (startStr.match(/^\d{1,2}:\d{2}/) && endStr.match(/^\d{1,2}:\d{2}/)) {
        const [startH, startM] = startStr.split(':').map(x => parseInt(x));
        const [endH, endM] = endStr.split(':').map(x => parseInt(x));
        
        if (!isNaN(startH) && !isNaN(startM) && !isNaN(endH) && !isNaN(endM)) {
          const startMins = startH * 60 + startM;
          let endMins = endH * 60 + endM;
          
          // Handle midnight crossing (jika end < start)
          if (endMins < startMins) {
            endMins += 24 * 60;
          }
          
          jam = (endMins - startMins) / 60;
          jam = parseFloat(jam.toFixed(2));
          
          if (jam > 0 && jam <= 24) {
            return jam;
          }
        }
      }
    } catch (e) {
      // Silent fail
    }
  }
  
  return 0;
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Inspeksi Teknik 2 (Daily Activity)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getLiveStats(rangeFilter = 'today', picFilter = 'all') {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const dailySheet = ss.getSheetByName(GID.DAILY_NAME);
    
    if (!dailySheet) return createEmptyStats();
    
    const dailyData = dailySheet.getDataRange().getValues();
    if (dailyData.length <= 1) return createEmptyStats();
    
    const timezone = "GMT+8"; 
    const now = new Date();
    const todayStr = Utilities.formatDate(now, timezone, "yyyy-MM-dd");
    let cutoffDate = new Date(now.getTime());
    
    if (rangeFilter === '7d') {
      cutoffDate.setDate(cutoffDate.getDate() - 6);
    } else if (rangeFilter === '30d') {
      cutoffDate.setDate(cutoffDate.getDate() - 29);
    } else if (rangeFilter === 'all') {
      cutoffDate = new Date(2000, 0, 1);
    }
    
    const cutoffStr = Utilities.formatDate(cutoffDate, timezone, "yyyy-MM-dd");
    const stats = {}; 
    const dailyAgg = {};
    const summary = { input: 0, jam: 0 };
    const recentData = [];

    for (let i = 1; i < dailyData.length; i++) {
      const row = dailyData[i];
      
      try {
        let tglRaw = row[1];
        if (!(tglRaw instanceof Date)) {
          if (typeof tglRaw === 'string' && tglRaw.trim()) {
            tglRaw = new Date(tglRaw);
          } else {
            continue;
          }
        }
        
        if (!(tglRaw instanceof Date) || isNaN(tglRaw.getTime())) continue;
        
        const tglStr = Utilities.formatDate(tglRaw, timezone, "yyyy-MM-dd");
        
        if (!dailyAgg[tglStr]) dailyAgg[tglStr] = { input: 0 };
        dailyAgg[tglStr].input += 1;
        
        if (tglStr >= cutoffStr && tglStr <= todayStr) {
          const nama = String(row[3] || "").trim();
          let jam = extractDuration(row);
          
          if (nama && nama !== "" && nama !== "undefined") {
            if (!stats[nama]) stats[nama] = { input: 0, jam: 0 };
            stats[nama].input += 1;
            stats[nama].jam += jam;
            
            recentData.push({
              nama: nama,
              pekerjaan: (row[9] || "").toString().substring(0, 35)
            });
          }
          
          if (picFilter === 'all' || nama === picFilter) {
            summary.input += 1;
            summary.jam += jam;
          }
        }
      } catch (e) {
        continue;
      }
    }

    return {
      lbInput: Object.entries(stats)
        .map(([nama, val]) => ({ nama, val: val.input }))
        .sort((a, b) => b.val - a.val)
        .slice(0, 10),
      lbJam: Object.entries(stats)
        .map(([nama, val]) => ({ nama, val: parseFloat(val.jam.toFixed(1)) }))
        .sort((a, b) => b.val - a.val)
        .slice(0, 10),
      summary: { input: summary.input, jam: parseFloat(summary.jam.toFixed(1)) },
      recentEntries: recentData.slice(-5).reverse(),
      chart: {
        labels: Object.keys(dailyAgg).sort().slice(-30).map(d => d.split('-').reverse().slice(0,2).join('/')),
        data: Object.keys(dailyAgg).sort().slice(-30).map(d => dailyAgg[d].input)
      }
    };
  } catch (error) {
    Logger.log("getLiveStats Error: " + error.toString());
    return createEmptyStats();
  }
}

function createEmptyStats() {
  return {
    lbInput: [],
    lbJam: [],
    summary: { input: 0, jam: 0 },
    recentEntries: [],
    chart: { labels: [], data: [] }
  };
}

function getMasterData() {
  try {
    if (MASTER_DATA_CACHE && (Date.now() - CACHE_TIME) < CACHE_DURATION) {
      return MASTER_DATA_CACHE;
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    
    // 1. Data Lokasi (Pabrik & Area)
    const locMap = {};
    const areaSheet = ss.getSheets().find(s => s.getSheetId() == GID.PABRIK_AREA);
    if (areaSheet) {
      const areaData = areaSheet.getDataRange().getValues().slice(1);
      areaData.forEach(r => {
        if (r[0]) {
          const pabrik = String(r[0]).trim();
          const area = String(r[1]).trim();
          if (!locMap[pabrik]) locMap[pabrik] = [];
          if (area && !locMap[pabrik].includes(area)) locMap[pabrik].push(area);
        }
      });
    }

    // 2. Data Equipment (by Area)
    const equipMap = {};
    const allEquipMap = new Map();
    const equipSheet = ss.getSheets().find(s => s.getSheetId() == GID.EQUIPMENT);
    if (equipSheet) {
      const equipData = equipSheet.getDataRange().getValues().slice(1);
      equipData.forEach(r => {
        const areaKey = String(r[1] || "").trim();
        const equipVal = String(r[2] || "").trim();
        const equipDesc = String(r[3] || "").trim();
        
        if (equipVal) {
          if (areaKey) {
            if (!equipMap[areaKey]) equipMap[areaKey] = [];
            equipMap[areaKey].push({ name: equipVal, desc: equipDesc });
          }
          if (!allEquipMap.has(equipVal)) allEquipMap.set(equipVal, equipDesc);
        }
      });
    }

    // 3. Data PIC & Bagian
    const bagianSheet = ss.getSheets().find(s => s.getSheetId() == GID.BAGIAN);
    const bagianList = bagianSheet.getRange(2, 1, bagianSheet.getLastRow()-1, 1).getValues().flat().filter(String);
    const picValues = ss.getSheets().find(s => s.getSheetId() == GID.PIC).getDataRange().getValues().slice(1);
    const picMap = {};
    const allNames = [];
    
    bagianList.forEach(b => picMap[b] = []);
    picValues.forEach(r => {
      const match = bagianList.find(b => String(r[1]).toLowerCase().includes(b.toLowerCase()));
      if (match) picMap[match].push({ nama: r[0], nik: r[3] });
      if (r[0]) allNames.push(r[0]);
    });

    const data = {
      bagianList: bagianList.sort(),
      allNames: [...new Set(allNames)].sort(),
      picMap,
      locMap,
      equipMap,
      allEquipList: Array.from(allEquipMap).map(([name, desc]) => ({ name, desc })).sort((a,b) => a.name.localeCompare(b.name)),
      status: ss.getSheets().find(s => s.getSheetId() == GID.STATUS).getDataRange().getValues().flat().filter(String).sort(),
      programs: ss.getSheets().find(s => s.getSheetId() == GID.PROGRAM).getDataRange().getValues().flat().filter(String).sort(),
      criticality: ss.getSheets().find(s => s.getSheetId() == GID.CRITICALITY).getDataRange().getValues().flat().filter(String)
    };

    MASTER_DATA_CACHE = data;
    CACHE_TIME = Date.now();

    return data;
  } catch (error) {
    Logger.log("getMasterData Error: " + error.toString());
    return { 
      bagianList: [], allNames: [], picMap: {}, locMap: {}, equipMap: {}, 
      allEquipList: [], status: [], programs: [], criticality: [] 
    };
  }
}

function submitData(payload) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(GID.DAILY_NAME);
    const generatedCodes = [];
    const rows = [];
    
    payload.entries.forEach(item => {
      const shortCode = Math.random().toString(36).substring(2, 9).toUpperCase();
      generatedCodes.push(shortCode);
      
      // Calculate durasi dari start_t dan end_t
      let lamaNumerik = 0;
      if (item.start_t && item.end_t) {
        try {
          const [startH, startM] = item.start_t.split(':').map(x => parseInt(x));
          const [endH, endM] = item.end_t.split(':').map(x => parseInt(x));
          
          if (!isNaN(startH) && !isNaN(startM) && !isNaN(endH) && !isNaN(endM)) {
            const startMins = startH * 60 + startM;
            let endMins = endH * 60 + endM;
            if (endMins < startMins) endMins += 24 * 60;
            lamaNumerik = parseFloat(((endMins - startMins) / 60).toFixed(2));
          }
        } catch (e) {
          lamaNumerik = 0;
        }
      }
      
      rows.push([
        shortCode, payload.common.tanggal, payload.common.bagian, payload.common.pic_nama, payload.common.pic_nik,
        item.pabrik, item.area, item.equipment, item.program, item.pekerjaan,
        item.status, item.wo, lamaNumerik, item.start_t, item.end_t, item.criticality, new Date()
      ]);
    });
    
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    
    MASTER_DATA_CACHE = null;
    return generatedCodes;
  } catch (error) {
    Logger.log("submitData Error: " + error.toString());
    throw new Error("Gagal menyimpan data: " + error.toString());
  }
}
