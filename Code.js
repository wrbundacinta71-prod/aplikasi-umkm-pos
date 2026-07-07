function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Aplikasi UMKM - Kasir & Utang')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Mengambil data produk dan pelanggan untuk form dropdown
function getMasterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetProduk = ss.getSheetByName("Produk");
  const dataProduk = sheetProduk.getDataRange().getValues();
  const listProduk = [];
  for(let i = 1; i < dataProduk.length; i++) {
    listProduk.push({ nama: dataProduk[i][0], tunai: dataProduk[i][1], utang: dataProduk[i][2] });
  }
  
  const sheetPelanggan = ss.getSheetByName("Pelanggan");
  const dataPelanggan = sheetPelanggan.getDataRange().getValues();
  const listPelanggan = [];
  for(let i = 1; i < dataPelanggan.length; i++) {
    listPelanggan.push({ grup: dataPelanggan[i][0], nama: dataPelanggan[i][1] });
  }
  
  return { produk: listProduk, pelanggan: listPelanggan };
}

// Menyimpan transaksi baru
function simpanTransaksi(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transaksi");
  const idTransaksi = "TRX-" + new Date().getTime();
  
  sheet.appendRow([
    idTransaksi,
    data.tanggal,
    data.grup,
    data.pelanggan,
    data.namaTransaksi,
    Number(data.utang),
    Number(data.bayar),
    data.metode
  ]);
  
  return idTransaksi;
}

// Mengambil riwayat transaksi dan menghitung saldo utang
function getRiwayat(grup, nama) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transaksi");
  const data = sheet.getDataRange().getValues();
  
  let riwayat = [];
  let totalUtang = 0;
  let totalBayar = 0;
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][2] === grup && data[i][3] === nama) {
      let utang = Number(data[i][5]) || 0;
      let bayar = Number(data[i][6]) || 0;
      
      totalUtang += utang;
      totalBayar += bayar;
      
      riwayat.push({
        id: data[i][0],
        tanggal: Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        transaksi: data[i][4],
        utang: utang,
        bayar: bayar
      });
    }
  }
  
  return {
    riwayat: riwayat,
    saldoUtang: totalUtang - totalBayar
  };
}
