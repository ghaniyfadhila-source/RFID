var idDatabaseInduk = "1ECx7xjKDL7iZw8uDmOtDDUrDHRlUQDv_oa4joP3YxH8"
var idDevice = SpreadsheetApp.openById(idDatabaseInduk).getSheetByName("id_whacenter").getRange(1,1).getValue()

function coba(){
  kirimPesanWa("08122516355", "test")
}

function doPost(e){
   var ss = SpreadsheetApp.openByUrl(e.parameter.url);
   var action = e.parameter.action;
  
  switch(action){
    case "login":
      return loginAdmin(e,ss);
      break;  
    case "presensiSiswa":
      return presensiSiswa(e,ss);
      break;  
    case "registerDevice":
      return registerDevice(e,ss);
      break;  
    case "insertJamKe0":
      return insertJamKe0(e,ss);
      break;
    case "insertSholat":
      return insertSholat(e,ss);
      break;
    case "insertBerhalanganSholat":
      return insertBerhalanganSholat(e,ss);
      break;
    case "presensiMandiri":
      return presensiMandiri(e,ss);
      break;
    case "notifMasuk" :
      return notifMasuk(e);
      break;  
     case "gantiPassword" :
      return gantiPassword(e,ss);
      break;   
      
  }
}


function doGet(e) {
  var ss = SpreadsheetApp.openByUrl(e.parameter.url);
  var action = e.parameter.action;
  
  
  switch(action){
    case "readPresensi":
      return readPresensi(e,ss);
      break;
    case "readPresensiBulanan":
      return readPresensiBulanan(e,ss);
      break;
    case "readPresensiSholatBulanan":
      return readPresensiSholatBulanan(e,ss);
      break;  
    case "info":
      return getContent(e,ss);
      break;
    case "readSPP":
      return readSPP(e,ss);
      break;
    case "readStatusRegisterPresensi":
      return readStatusRegisterPresensi(e,ss);
      break;
    case "readStatusPresensi":
      return readStatusPresensi(e,ss);
      break;  
    case "readDaftarNilai":
      return readDaftarNilai(e,ss);
      break;
    case "generateQr":
      return generateQr(e);
      break;
    case "readJurnalKelas":
      return jurnalKelas(e,ss);
      break;
    case "readStatusRekamWajah":
      return statusRekamWajah(e,ss);
      break;
    case "readJadwalHarian":
      return readJadwalHarian(e,ss);
      break;
    case "readLokasiAcuan":
      return readLokasiAcuan(e,ss);
      break;
    case "readMateri":
      return readMateri(e,ss);
      break;
    case "readStatusPengumuman":
      return statusPengumumanKelulusan(ss);
      break;  
    case "readKelulusan":
      return readKelulusan(e,ss);
      break;
    case "rekapPresensi":
      return rekapPresensi(e,ss);
      break;  
  }
  
}


function rekapPresensi(request,ss){
  var bulan = request.parameter.bulan
  var kode = request.parameter.kode
  var sheet = ss.getSheetByName(bulan)
  var guru = kode
  
  var data = sheet.getDataRange().getValues()
  
  var filterGuru = data.filter(function(pegawai){
    if (pegawai[2] == guru){
      return pegawai
    }
  
  })
  
  var output = JSON.stringify({"rekap": filterGuru[0].splice(3,93).toString()});

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
  
//  Logger.log(filterGuru[0][3])

}

function loginAdmin(e,ss){
  var username = e.parameter.username
  var pass = e.parameter.pass
  var flag = 0
  var data = ss.getSheetByName("Wali Kelas").getDataRange().getValues()
  var filterData = data.filter(function(admin){
    if(admin[0] == username && admin[1] == pass){
      flag = 1
      return admin
      
    }
  })
  
  
  if(flag == 0){
    var hasil = {"hasil":"gagal"}
  }else{
    var hasil = {"hasil":"sukses"}
  }
  
  var output = JSON.stringify(hasil)
  
  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function gantiPassword(e,ss) {
var nis = e.parameter.nis;
  var tingkatan = e.parameter.tingkatan;
  var kelas = e.parameter.kelas;
  var passLama =e.parameter.passLama;
  var passBaru = e.parameter.passBaru;

//var linkDbSiswa = sheet.getRange(2,3).getValue();

 // var ssServerSiswa = SpreadsheetApp.openByUrl(linkDbSiswa);
  var sheetServerSiswa = ss.getSheetByName("Kelas "+tingkatan);
  var dataServerSiswa = sheetServerSiswa.getDataRange().getValues();
 
  var filterServerSiswa = dataServerSiswa.filter(function(server,index){
    if(server[0].toString() == nis && server[3].toString()==passLama){
      
      sheetServerSiswa.getRange(index+1, 4).setValue(passBaru)
      return server;
    
    }
  })
     
  
  if(filterServerSiswa.length > 0){
    var hasil = {"hasil": "sukses"} 
    }else{
    var hasil = {"hasil": "failed"}
    }
  var output = JSON.stringify(hasil)
  
return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);  
    
}
                                               
function presensiSiswa(request,ss){
  var tanggal = request.parameter.tanggal;
  var bulan = request.parameter.bulan;
  var nis = request.parameter.nis;
  var androidId = request.parameter.androidId;
  var namaPegawai = request.parameter.nama;
  var jenisPresensi = request.parameter.jenisPresensi;
  var lokasiPresensi = request.parameter.lokasiPresensi;
  var noWa = request.parameter.noWa;
  var sheetPresensi = ss.getSheetByName(bulan);
  var sheetWaktuAcuan = ss.getSheetByName("Detail Sekolah")
  var kolom = kolomPresensiKehadiran(tanggal);
//  var waktuPresensiGuru = waktuPresensi();
  var date = new Date()
  var jam = getJam(date)
  var menit = getMenit(date)
  
  //ini untuk detect pinjam hp
  var sheetDataSiswa = ss.getSheetByName("Data Siswa")
  var dataDevice = sheetDataSiswa.getDataRange().getValues()
  var filterDevice = dataDevice.filter(function(device){
    if(device[3] == androidId){
      return device
    }

  })
  
  if(filterDevice.length>0){
    var column = 0
    var jamAcuan = 0
    var jamDatangMin = 0
    var jamPulangMax = 0
    var menitAcuan = 0
    var pesan = ""
    var flag = 0
    if(jenisPresensi == "hadir"){
      column = kolom
      jamAcuan = sheetWaktuAcuan.getRange(2,5).getValue()
      menitAcuan = sheetWaktuAcuan.getRange(2,6).getValue()
      jamDatangMin = sheetWaktuAcuan.getRange(2,9).getValue()
    
      if(jam<jamAcuan){
        if(jam<jamDatangMin){
          flag = 1
          pesan = "🛎️INFORMASI DARI PUSAT DATA SMK NEGERI 9 SEMARANG🛎️ \n"
          +namaPegawai+" *gagal* melakukan presensi *hadir* yaitu *"+jam+":"+menit+"* minimal presensi hadir adalah "+jamDatangMin+":00"+"⚠️Hubungi nomor dibawah ini ini jika pengguna mengalami kendala atau ketidaksesuaian pada aplikasi⚠️ \n" + "\n" + "- 087767930440 Nanda📞 \n" + "- 083842372975 Keyza 📞 \n" + "\n" + "Sekian informasi dari kami Terima Kasih🙏🏻"
//        }else if(jam == 6 && menit < 50){
//          flag = 2
//          pesan = " 🛎️INFORMASI DARI PUSAT DATA SMK NEGERI 9 SEMARANG🛎️ \n"+ namaPegawai+" *gagal* melakukan presensi *hadir* yaitu *"+jam+":"+menit+"* minimal presensi hadir adalah 6:50"+"⚠️Hubungi nomor dibawah ini ini jika pengguna mengalami kendala atau ketidaksesuaian pada aplikasi⚠️ \n" + "\n" + "- 087767930440 Nanda📞 \n" + "- 083842372975 Keyza 📞 \n" + "\n" + "Sekian informasi dari kami Terima Kasih🙏🏻"
        }else if(jam == jamDatangMin){
          flag = 5
          pesan = "🛎️INFORMASI DARI PUSAT DATA SMK NEGERI 9 SEMARANG🛎️ \n" 
          +namaPegawai+" presensi hadir *terpenuhi* yaitu *"+jam+":"+menit+"*"+"⚠️Hubungi nomor dibawah ini ini jika pengguna mengalami kendala atau ketidaksesuaian pada aplikasi⚠️ \n" + "\n" + "- 087767930440 Nanda📞 \n" + "- 083842372975 Keyza 📞 \n" + "\n" + "Sekian informasi dari kami Terima Kasih🙏🏻"
        }else{
          flag = 6
          pesan = "🛎️INFORMASI DARI PUSAT DATA SMK NEGERI 9 SEMARANG🛎️ \n"
          +namaPegawai+" presensi hadir *terpenuhi* yaitu *"+jam+":"+menit+"*"+"⚠️Hubungi nomor dibawah ini ini jika pengguna mengalami kendala atau ketidaksesuaian pada aplikasi⚠️ \n" + "\n" + "- 087767930440 Nanda📞 \n" + "- 083842372975 Keyza 📞 \n" + "\n" + "Sekian informasi dari kami Terima Kasih🙏🏻"
        } 
        
      }else if(jam==jamAcuan && menit <= menitAcuan ){
        flag = 7
        pesan = "🛎️INFORMASI DARI PUSAT DATA SMK NEGERI 9 SEMARANG🛎️ \n"
        +namaPegawai+" presensi hadir *terpenuhi* yaitu *"+jam+":"+menit+"*"+"⚠️Hubungi nomor dibawah ini ini jika pengguna mengalami kendala atau ketidaksesuaian pada aplikasi⚠️ \n" + "\n" + "- 087767930440 Nanda📞 \n" + "- 083842372975 Keyza 📞 \n" + "\n" + "Sekian informasi dari kami Terima Kasih🙏🏻"
      }else{
        flag = 8
        pesan = "🛎️INFORMASI DARI PUSAT DATA SMK NEGERI 9 SEMARANG🛎️ \n"
        +namaPegawai+" tercatat melakukan presensi hadir *melebihi waktu* yaitu *"+jam+":"+menit+"*"+"⚠️Hubungi nomor dibawah ini ini jika pengguna mengalami kendala atau ketidaksesuaian pada aplikasi⚠️ \n" + "\n" + "- 087767930440 Nanda📞 \n" + "- 083842372975 Keyza 📞 \n" + "\n" + "Sekian informasi dari kami Terima Kasih🙏🏻"
      }
  
    
    }else{
      jamAcuan = sheetWaktuAcuan.getRange(2,7).getValue()
      menitAcuan = sheetWaktuAcuan.getRange(2,8).getValue()
      column = kolom+1
      jamPulangMax = sheetWaktuAcuan.getRange(2,10).getValue()
      //maksimal presensi 16.30
      if(jam<jamAcuan){
        flag = 9
        pesan = "🛎️INFORMASI DARI PUSAT DATA SMK NEGERI 9 SEMARANG🛎️ \n"
        +namaPegawai+" tercatat melakukan presensi pulang *sebelum waktunya* yaitu *"+jam+":"+menit+"*" +"⚠️Hubungi nomor dibawah ini ini jika pengguna mengalami kendala atau ketidaksesuaian pada aplikasi⚠️ \n" + "\n" + "- 087767930440 Nanda📞 \n" + "- 083842372975 Keyza 📞 \n" + "\n" + "Sekian informasi dari kami Terima Kasih🙏🏻"
      }else if(jam>=jamAcuan){
        
        if(jam<jamPulangMax){
          flag = 10 
          pesan = "🛎️INFORMASI DARI PUSAT DATA SMK NEGERI 9 SEMARANG🛎️ \n"
          +namaPegawai+" presensi pulang *terpenuhi* yaitu *"+jam+":"+menit+"*"+"⚠️Hubungi nomor dibawah ini ini jika pengguna mengalami kendala atau ketidaksesuaian pada aplikasi⚠️ \n" + "\n" + "- 087767930440 Nanda📞 \n" + "- 083842372975 Keyza 📞 \n" + "\n" + "Sekian informasi dari kami Terima Kasih🙏🏻"
        }else if(jam == jamPulangMax && menit<=menitAcuan){
          flag = 11
          pesan = "🛎️INFORMASI DARI PUSAT DATA SMK NEGERI 9 SEMARANG🛎️ \n"
          +namaPegawai+" presensi pulang *terpenuhi* yaitu *"+jam+":"+menit+"*"
        }else if(jam == jamPulangMax && menit>menitAcuan){
          flag = 3
          pesan = "🛎️INFORMASI DARI PUSAT DATA SMK NEGERI 9 SEMARANG🛎️ \n"
          +namaPegawai+" *gagal* presensi pulang yaitu *"+jam+":"+menit+"* presensi pulang maksimal "+jamPulangMax+":"+menitAcuan+"0"+"⚠️Hubungi nomor dibawah ini ini jika pengguna mengalami kendala atau ketidaksesuaian pada aplikasi⚠️ \n" + "\n" + "- 087767930440 Nanda📞 \n" + "- 083842372975 Keyza 📞 \n" + "\n" + "Sekian informasi dari kami Terima Kasih🙏🏻"
        }else{
          flag = 4
          pesan = "🛎️INFORMASI DARI PUSAT DATA SMK NEGERI 9 SEMARANG🛎️ \n"
          +namaPegawai+" *gagal* presensi pulang yaitu *"+jam+":"+menit+"* presensi pulang maksimal "+jamPulangMax+":"+menitAcuan+"0"+"⚠️Hubungi nomor dibawah ini ini jika pengguna mengalami kendala atau ketidaksesuaian pada aplikasi⚠️ \n" + "\n" + "- 087767930440 Nanda📞 \n" + "- 083842372975 Keyza 📞 \n" + "\n" + "Sekian informasi dari kami Terima Kasih🙏🏻"
        }
//        flag = 9
//        pesan = namaPegawai+" presensi pulang *terpenuhi* yaitu *"+jam+":"+menit+"*" +"⚠️Hubungi nomor dibawah ini ini jika pengguna mengalami kendala atau ketidaksesuaian pada aplikasi⚠️ \n" + "\n" + "- 087767930440 Nanda📞 \n" + "- 083842372975 Keyza 📞 \n" + "\n" + "Sekian informasi dari kami Terima Kasih🙏🏻"
      }
    }
  
  
    var dataPresensi = sheetPresensi.getDataRange().getValues()
    var filterSiswa = dataPresensi.filter(function(siswa,row){
      if(siswa[2] == nis){
        var statusLokasi = sheetDataSiswa.getRange(row,7).getValue()
//        sheetPresensi.getRange(row+1, column).setValue(jam+":"+menit)
        if(flag >= 5){
        
          var hari = getDay() 
          if(hari == 6 || hari ==0){
          
          }else{
            if((typeof lokasiPresensi == 'undefined') || (typeof lokasiPresensi == null)){
              //membuat gagal yang belum update ke app terbaru
              flag = 0
            }else if(lokasiPresensi =="rumah"){
              //jika statusLokasi nya 1, berarti harusnya presensi di sekolah
              if(statusLokasi == "1"){
                pesan = namaPegawai+" *Gagal* presensi. Sistem mencatat pada hari ini Anda seharusnya tidak presensi di Rumah melainkan di Sekolah"
                flag = -1
                kirimPesanWa(noWa,pesan)
//                kirimPesan(grupIdMyFamily,pesan,"markdown")
              }else if(statusLokasi == "0"){
                sheetPresensi.getRange(row+1, column).setValue(jam+":"+menit).setBackground("yellow")
                kirimPesanWa(noWa,pesan)
//                kirimPesan(grupIdMyFamily,pesan+" *Lokasi Presensi di Rumah*","markdown")
              }else{
                pesan = namaPegawai+" *Gagal* presensi. Sistem mencatat pada hari ini Anda seharusnya tidak presensi di Rumah melainkan di tempat Magang"
                flag = -3
                kirimPesanWa(noWa,pesan)
              }
              
              
              
            }else if(lokasiPresensi =="sekolah"){
//              ini presensi di sekolah
              if(statusLokasi == "1"){
                sheetPresensi.getRange(row+1, column).setValue(jam+":"+menit).setBackground("green")
                kirimPesanWa(noWa,pesan)
//                kirimPesan(grupIdMyFamily,pesan+" *Lokasi Presensi di Sekolah*","markdown")  
              }else if(statusLokasi == "0"){
                pesan = namaPegawai+" *Gagal* presensi. Sistem mencatat pada hari ini Anda seharusnya tidak presensi di Sekolah melainkan di Rumah"
                flag = -2
                kirimPesanWa(noWa,pesan)
//                kirimPesan(grupIdMyFamily,pesan,"markdown")
              }else{
                pesan = namaPegawai+" *Gagal* presensi. Sistem mencatat pada hari ini Anda seharusnya tidak presensi di Sekolah melainkan di tempat Magang"
                flag = -3
                kirimPesanWa(noWa,pesan)
              }
//              sheetPresensi.getRange(row+1, column).setValue(jam+":"+menit).setBackground("green")
//              kirimPesan(grupIdMyFamily,pesan+" *Lokasi Presensi di Sekolah*","markdown")  
            }else if(lokasiPresensi =="magang"){
            //ini presensi magang
              if(statusLokasi == "2"){
                sheetPresensi.getRange(row+1, column).setValue(jam+":"+menit).setBackground("brown")
                kirimPesanWa(noWa,pesan)
//                kirimPesan(grupIdMyFamily,pesan+" *Lokasi Presensi di Sekolah*","markdown")  
              }else if(statusLokasi == "0"){
                pesan = namaPegawai+" *Gagal* presensi. Sistem mencatat pada hari ini Anda seharusnya tidak presensi di tempat Magang melainkan di Rumah"
                flag = -2
                kirimPesanWa(noWa,pesan)
//                kirimPesan(grupIdMyFamily,pesan,"markdown")
              }else{
                pesan = namaPegawai+" *Gagal* presensi. Sistem mencatat pada hari ini Anda seharusnya tidak presensi di tempat Magang melainkan di Sekolah"
                flag = -1
                kirimPesanWa(noWa,pesan)
              }
            }
            
          } 
        
        }else{
          //ini presensi di luar jam 6.50 - 16.30 tidak dimasukan ke database
          var hari = getDay()
          if(hari ==0){
          
          }else{  
//            kirimPesan(grupIdMyFamily,pesan,"markdown")
          }
        }
      return siswa
    }
  })
  
//  if(filterGuru.length>0){
//      var hasil ="succes"
//    }else{
//      var hasil = "failed"
//    }
    
    // if(flag>=5){
    //   var hasil ="succes"
    // }else if (flag>=0 && flag<5){
    //   var hasil = "failed"
    //   } else if(flag == -1){
    //     var hasil = "sekolah"
    //     }else if(flag == -2){
    //       var hasil = "rumah"
    //       }else{
    //        var hasil = "magang"
    //       }

    var hasil = pesan
  
    var output = JSON.stringify({"hasil": hasil});

    return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
  }else{
    var output = JSON.stringify({"hasil": "denied"});

    return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
  
  }
}

function readStatusPresensi(request,ss){
  var tanggal = request.parameter.tanggal;
  var bulan = request.parameter.bulan;
  var nis = request.parameter.nis;
  var jenisPresensi = request.parameter.jenisPresensi;
  var sheetPresensi = ss.getSheetByName(bulan);
  var kolom = kolomPresensiKehadiran(tanggal);
  var waktuPresensiSiswa = waktuPresensi();
  
  
  var column = 0
  if(jenisPresensi == "hadir"){
    column = kolom
    }else{
      column = kolom+1
    }
  var status = ""
  var dataPresensi = sheetPresensi.getDataRange().getValues()
  var filterSiswa = dataPresensi.filter(function(siswa,row){
    if(siswa[2] == nis){
      status = sheetPresensi.getRange(row+1, column).getValue()
      return siswa
    }
  })
  
  if(status == ""){
    var hasil = "belum"
    }else{
    var hasil = "sudah"
    }
  
  var output = JSON.stringify({"status": hasil});

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);

}


function registerDevice(request,ss){
  var sheetSiswa = ss.getSheetByName("Data Siswa")
  var nis = request.parameter.nis
  var androidId = request.parameter.androidId
  var flag = 0
  var dataSiswa = sheetSiswa.getDataRange().getValues()
  var filterSiswa = dataSiswa.filter(function(siswa,row){
    if(siswa[2] == nis && siswa[3]==""){
      sheetSiswa.getRange(row+1,4).setValue(androidId)
      flag = 1
      return siswa
    }
  
  })
  
  if(flag == 0){
    var hasil = "gagal"
    }else{
    var hasil = "sukses"
    }
  var output = JSON.stringify({"status": hasil});

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}


function readStatusRegisterPresensi(request,ss){

  var sheetSiswa = ss.getSheetByName("Data Siswa")
  var nis = request.parameter.nis
  var flag = 0
  var dataSiswa = sheetSiswa.getDataRange().getValues()
  var filterSiswa = dataSiswa.filter(function(siswa,row){
    if(siswa[2] == nis && siswa[3]==" "){
      flag = 1
      return siswa
    }
  
  })
  
  if(flag == 0){
    var hasil = "sudah"
    }else{
    var hasil = "belum"
    }
  var output = JSON.stringify({"status": hasil});

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
  
  
}


function notifMasuk(request){
  var nis = request.parameter.nis
  var nama = request.parameter.nama
  var pesan = request.parameter.pesan
  var noHp = request.parameter.noHp
  
//  var noHp = 

}
  
function filterHpOrtu(nis,ss){
  var sheetKelas = ss.getSheetByName("Kelas "+tingkatan)
  var dataOrtu = sheetKelas.getDataRange().getValues()
  
  var filterOrtu = dataOrtu.filter(function(ortu){
    if(ortu[0] == nis){
      return ortu
    }
  })
  
  if(filterOrtu.length>0){
    return filterOrtu[0][6]
  }else{
    return 0
  }
  
}

function statusPengumumanKelulusan(ss){
  var sheetKelulusan = ss.getSheetByName("status_pengumuman")
  var statusPengumuman = sheetKelulusan.getRange(2,1).getValue()
  var output = JSON.stringify({"status":statusPengumuman});

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function readKelulusan(request,ss){
  var nis = request.parameter.nis
  var sheetKelulusan = ss.getSheetByName("kelulusan")
  var dataSiswa = sheetKelulusan.getDataRange().getValues()
  var flag = 0
  
  var filterSiswa = dataSiswa.filter(function(siswa){
    if(siswa[1] == nis){
    flag = 1
    return siswa
    }
  })
  
  if(flag == 0){
    var hasil = "gagal"
    }else{
      var hasil = filterSiswa[0][4]
    }
  
  var output = JSON.stringify({"hasil": hasil})
  
 return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JSON)

}

function readMateri(request,ss){
  var bulan = request.parameter.bulan;
  var sheetJurnal = ss.getSheetByName("Jurnal "+bulan)
  var tanggal = request.parameter.tanggal;
  var jamKe = request.parameter.jamKe;
 
  var column = 0;
  switch(Number(tanggal)){
    case 1 :
      column = 9
      break;
    case 2 :
      column = 21
      break;
    case 3 :
      column = 33
      break;
    case 4 :
      column = 45
      break;
    case 5 :
      column = 57
      break;
    case 6 :
      column = 69
      break;
    case 7 :
      column = 81
      break;
    case 8 :
      column = 93
      break;
    case 9 :
      column = 105
      break;
    case 10 :
      column = 117
      break;
    case 11 :
      column = 129
      break;
    case 12 :
      column = 141
      break;
    case 13 :
      column = 153
      break;
    case 14 :
      column = 165
      break;
    case 15 :
      column = 177
      break;
    case 16 :
      column = 189
      break;
    case 17 :
      column = 201
      break;
    case 18 :
      column = 213
      break;
    case 19 :
      column = 225
      break;
    case 20 :
      column = 237
      break;
    case 21 :
      column = 249
      break;
    case 22 :
      column = 261
      break;
    case 23 :
      column = 273
      break;
    case 24 :
      column = 285
      break;
    case 25 :
      column = 297
      break;
    case 26 :
      column = 309
      break;
    case 27 :
      column = 321
      break;
    case 28 :
      column = 333
      break;
    case 29 :
      column = 345
      break;
    case 30 :
      column = 357
      break;  
    case 31 :
      column = 369
      break;   
  }
  var isiMateri = sheetJurnal.getRange(sheetJurnal.getLastRow()-1, column+Number(jamKe)-1).getValue();
  
  var output = JSON.stringify({"materi":isiMateri});

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function readLokasiAcuan(request,ss){
  var posisi = request.parameter.posisi
  var sheetDataSekolah = ss.getSheetByName("Detail Sekolah")
  var jarakMax =  sheetDataSekolah.getRange(2,4).getValue()
  var filterSiswa = filterDataSiswa(ss,request)
  if(posisi == 'sekolah'){
     
     var lat = sheetDataSekolah.getRange(2,2).getValue()
     var lng = sheetDataSekolah.getRange(2,3).getValue()
     
  
     if(lat == ""){
       var hasil = {
         "lokasi" : "error",
         "latAcuan" : "error",
         "longAcuan": "error",
         "jarakMax" : 1,
         "noWa" : "0"
       }
     }else{
       var hasil = {
         "lokasi" : "success",
         "latAcuan" : lat,
         "longAcuan": lng,
         "jarakMax" : jarakMax,
         "noWa" : filterSiswa[0][16]
       }  
       }
  }else if(posisi == "rumah"){
    
    if(filterSiswa.length>0){
      var hasil = {
         "lokasi" : "success",
         "latAcuan" : filterSiswa[0][4],
         "longAcuan": filterSiswa[0][5],
         "jarakMax" : filterSiswa[0][15],
         "noWa" : filterSiswa[0][16]
      }
     }else{
        var hasil = {
         "lokasi" : "error",
         "latAcuan" : "error",
         "longAcuan": "error",
         "jarakMax" : 1,
         "noWa" : "0"
       }
     }
      
      
  
  }else{
    if(filterSiswa.length>0){
      var hasil = {
         "lokasi" : "success",
         "latAcuan" : filterSiswa[0][18],
         "longAcuan": filterSiswa[0][19],
         "jarakMax" : filterSiswa[0][15],
         "noWa" : filterSiswa[0][16]
       }
      }else{
        var hasil = {
         "lokasi" : "error",
         "latAcuan" : "error",
         "longAcuan": "error",
         "jarakMax" : 1,
         "noWa" : "0" 
       }
      
      }
  }
//  else{
//    var lat = sheetDataSekolah.getRange(2,2).getValue()
//     var lng = sheetDataSekolah.getRange(2,3).getValue()
//     
//  
//     if(lat == ""){
//       var hasil = {
//         "lokasi" : "error",
//         "latAcuan" : "error",
//         "longAcuan": "error",
//         "jarakMax" : 1
//       }
//     }else{
//       var hasil = {
//         "lokasi" : "success",
//         "latAcuan" : lat,
//         "longAcuan": lng,
//         "jarakMax" : jarakMax
//       }  
//     }
//  }
  
  
  var output = JSON.stringify(hasil);

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function filterDataSiswa(ss,request){
  var sheetSiswa = ss.getSheetByName("Data Siswa")
    var nis = request.parameter.nis
    var flag = 0
    var dataSiswa = sheetSiswa.getDataRange().getValues()
    var filterSiswa = dataSiswa.filter(function(siswa){
      if(siswa[2] == nis){
        flag = 1
        return siswa
      }  
    })
    
    return filterSiswa
}

function readJadwalHarian(request,ss){
  var sheetJadwal = ss.getSheetByName("Jadwal Pelajaran")
  var date = new Date();
  var hari = date.getDay()
  var column = 0
  switch(hari){
    case 0:
      column = 9
      break
    case 1:
      column = 3
      break
    case 2:
      column = 4
      break
    case 3:
      column = 5
      break
    case 4:
      column = 6
      break
    case 5:
      column = 7
      break
    case 6:
      column = 8
      break 
      
  }
  var listMateri = sheetJadwal.getRange(2, column,sheetJadwal.getLastRow()-1, 1).getValues()
  var listWaktu = sheetJadwal.getRange(2, 2,sheetJadwal.getLastRow()-1, 1).getValues()
  var waktu = listWaktu.map(function(r){return r[0]})
  var materi = []
//  var waktu = ["7.00-7.45","7.45-8.30","8.30-9.15","9.30-10.15","10.15-11.00","11.00-11.45","11.45-12.30","13.00-13.45","13.45-14.30","14.30-15.15","15.15-16.00"]
  for(var i=0;i<listMateri.length;i++){
    var mapel = {
      "jamKe":i+1,
      "durasi":waktu[i],
      "mapel":listMateri[i][0] 
    }
    materi.push(mapel)
      }
  
  var hasil = {"listMapel":materi}
  var output = JSON.stringify(hasil);

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function statusRekamWajah(request,ss){
  var sheetLokasi = ss.getSheetByName("Data Siswa")
  var nis = request.parameter.nis
  var dataSiswa = sheetLokasi.getDataRange().getValues()
  var flag = 0
  var filterSiswa = dataSiswa.filter(function(siswa){
    if(siswa[2] == nis){
      flag = 1
      return siswa
    }
  
  })
  
  if(flag == 0){
    var hasil = {
      "hasil":"error",
      "status":filterSiswa[0][9]
    }
  }else{
    var hasil = {
      "hasil":"success",
      "status":filterSiswa[0][9]
    }
  }
  
  var output = JSON.stringify(hasil);

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function presensiMandiri(request,ss){
  var nis = request.parameter.nis;
  var bulan = request.parameter.bulan;
  var sheetJurnal = ss.getSheetByName("Jurnal "+bulan)
  var tanggal = request.parameter.tanggal;
  var jamKeUser = request.parameter.jamKeUser;
  var materi = request.parameter.materi;
  var terikatWaktu = request.parameter.terikatWaktu;
  var date = new Date();
  var waktu = Utilities.formatDate(date, 'Asia/Jakarta', 'HH:mm');
  var splitWaktu = waktu.split(":")
  var jam = splitWaktu[0]
  var menit = splitWaktu[1]
  
  //mohon menyesuaikan dengan interval waktu di masing masing sekolah
  var jamKe = 0
  switch(Number(jam)){
    
    case 8 :
      if(menit>=0){
        //8.00 - 8.59
        jamKe = 1
      }
      break
    case 9 :
      if(menit>=0){
        //8.00 - 8.59
        jamKe = 2
      }
      break
    case 10 :
      if(menit>=0){
        //8.00 - 8.59
        jamKe = 3
      }
      break
  
    case 11 :
      if(menit>=0){
        //8.00 - 8.59
        jamKe = 4
      }
      break
  }
  
  
  
  var flag = 0
  var x = 0
  
  //di sini melakukan filtering untuk presensi sesuai jam atau tidak
  if(terikatWaktu == "on"){
    if(Number(jamKeUser)<jamKe){
      flag = 1
    }else if(Number(jamKeUser)==jamKe){
      flag = 2
      var dataSiswa = sheetJurnal.getDataRange().getValues();
      var filterSiswa = dataSiswa.filter(function(siswa,row){
        if(siswa[2]==nis){
          var kolom = getKolomKehadiran(tanggal)
          sheetJurnal.getRange(row+1, kolom+jamKe-1).setValue("H")
          sheetJurnal.getRange(row+2, kolom+jamKe-1).setValue(materi)
          return siswa
        }
       })
      }
  }else{
    flag = 2
      var dataSiswa = sheetJurnal.getDataRange().getValues();
      var filterSiswa = dataSiswa.filter(function(siswa,row){
        if(siswa[2]==nis){
          var kolom = getKolomKehadiran(tanggal)
          sheetJurnal.getRange(row+1, kolom+Number(jamKeUser)-1).setValue("H")
          sheetJurnal.getRange(row+2, kolom+Number(jamKeUser)-1).setValue(materi)
          return siswa
        }
      })
  }
  
  
  
  
  if(flag == 0){
    var status = {"status":"Presensi gagal,waktu presensi diluar jam pembelajaran"}
    }else if(flag == 1){
    var status = {"status":"Presensi gagal,waktu presensi telah lewat"}
    }else if(flag == 2){
    var status = {"status":"success"}
    }
  
  var output = JSON.stringify(status);

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
  
  
}

function getKolomKehadiran(tanggal){
  var column = 0;
  switch(Number(tanggal)){
    case 1 :
      column = 9
      break;
    case 2 :
      column = 21
      break;
    case 3 :
      column = 33
      break;
    case 4 :
      column = 45
      break;
    case 5 :
      column = 57
      break;
    case 6 :
      column = 69
      break;
    case 7 :
      column = 81
      break;
    case 8 :
      column = 93
      break;
    case 9 :
      column = 105
      break;
    case 10 :
      column = 117
      break;
    case 11 :
      column = 129
      break;
    case 12 :
      column = 141
      break;
    case 13 :
      column = 153
      break;
    case 14 :
      column = 165
      break;
    case 15 :
      column = 177
      break;
    case 16 :
      column = 189
      break;
    case 17 :
      column = 201
      break;
    case 18 :
      column = 213
      break;
    case 19 :
      column = 225
      break;
    case 20 :
      column = 237
      break;
    case 21 :
      column = 249
      break;
    case 22 :
      column = 261
      break;
    case 23 :
      column = 273
      break;
    case 24 :
      column = 285
      break;
    case 25 :
      column = 297
      break;
    case 26 :
      column = 309
      break;
    case 27 :
      column = 321
      break;
    case 28 :
      column = 333
      break;
    case 29 :
      column = 345
      break;
    case 30 :
      column = 357
      break;  
    case 31 :
      column = 369
      break;   
  }
  
  return column
}

function jurnalKelas(request,ss){
  var tingkatan = request.parameter.tingkatan;
  var kelas = request.parameter.kelas;
  var sheet = ss.getSheetByName("Kelas "+tingkatan);
  var listKelas = sheet.getDataRange().getValues();
  var flag = 0;
  
  var filterKelas = listKelas.filter(function(linkKelas){
    if(linkKelas[1] == kelas){
      flag = 1
      return linkKelas
    }
  })
  
  if(flag == 0){
    var data = {"linkKelas":"error"}
    }else{
      var data = {"linkKelas": filterKelas[0][3]}  
    }
  
  var hasil = JSON.stringify(data);
  
  return ContentService
         .createTextOutput(hasil)
         .setMimeType(ContentService.MimeType.JSON);
}

function readDaftarNilai(request,ss){
  var tingkatan = request.parameter.tingkatan;
  var sheet = ss.getSheetByName("Kelas "+tingkatan);
  var kelas = request.parameter.kelas;
  var nis = request.parameter.nis;
//  var kelas = "11 AP 1";
  
  
  
  var data = sheet.getDataRange().getValues();
  
  var filterKelas = data.filter(function(kelasSiswa){
    if(kelasSiswa[1] == kelas){
      return kelasSiswa;
    }
  });
  
  var linkKelas = filterKelas[0][3];
  var ssKelas = SpreadsheetApp.openByUrl(linkKelas);
  var sheetKelas = ssKelas.getSheetByName("nilai");
  var flag = 0;
//  var data = sheetKelas.getDataRange().getValues();
//  var lastColumn = sheetKelas.getLastColumn()
//  var rowSiswa = 0;
  var dataNilai = {};
  dataNilai["nilai"] = contentNilai(nis,sheetKelas);
//  var filterSiswa = data.filter(function(siswa,row){
//    if(siswa[0] == nis){
//      flag = 1;
//      rowSiswa = row;
//      return siswa;  
//    }
//  
//  });
//  
//  var objek = {};
//    
//  for (var column = 0; column<lastColumn;column++){
//        objek[data[0][column]] = data[rowSiswa][column];  
//  }
  
//  if(flag == 0){
//    var hasil = {"nilai":"gagal"}
//    }else{
//    var hasil = dataNilai
//    }
    
  var output = JSON.stringify(dataNilai);

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function contentNilai(nis,sheet){
var allData = [];
  
  var rangeData = sheet.getDataRange().getValues();
  var rowSiswa = 0;
  var filterPresensi = rangeData.filter(function(siswa,row){
    if(siswa[0]==nis){
    rowSiswa = row;
    return siswa;  
    }
  });
  

    var objek = {};
    for (var column = 0; column<21;column++){
      if(column<=1){   
      
        objek[rangeData[0][column]] = rangeData[rowSiswa][column];
      
      }else if(column == 2){
        objek[rangeData[0][2]]= {
              "PTS Gasal" : rangeData[rowSiswa][column],
              "PAS Gasal" : rangeData[rowSiswa][column+1],
              "PTS Genap" : rangeData[rowSiswa][column+2],
              "PAS Genap" : rangeData[rowSiswa][column+3]
              
        }
      }else{
        var a = column+4+3*(column-3)-1
        objek[rangeData[0][a]]= {
              "PTS Gasal" : rangeData[rowSiswa][a],
              "PAS Gasal" : rangeData[rowSiswa][a+1],
              "PTS Genap" : rangeData[rowSiswa][a+2],
              "PAS Genap" : rangeData[rowSiswa][a+3]
              
        }
      }
       
      
    }   
    allData.push(objek);     
  return allData;
}


function readPresensiSholatBulanan(request,ss){
  var tingkatan = request.parameter.tingkatan;
  var sheet = ss.getSheetByName("Kelas "+tingkatan);
  var nis = request.parameter.nis;
  var kelas = request.parameter.kelas;
  var bulan = request.parameter.bulan;
  
  
  var data = sheet.getDataRange().getValues();
  
  var filterKelas = data.filter(function(kelasSiswa){
    if(kelasSiswa[1] == kelas){
      return kelasSiswa;
    }
  });
  
//  var jumlahSiswa = filterKelas[0][2];
  var linkKelas = filterKelas[0][3];
  var ssKelas = SpreadsheetApp.openByUrl(linkKelas);
  var sheetKelas = ssKelas.getSheetByName("Absen "+bulan);
  
//  var ss = SpreadsheetApp.openById("1kyWZ0X9oiAVtR9ZeuwbZvbApx9btRXfoUgSZS9zvDr0");
//  var sheet = ss.getSheetByName("Absen JULI");
  
  
  var data = {};
  data["results"] = contentPresensiSholat(nis,sheetKelas, bulan);
  
  var output = JSON.stringify(data);

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function contentPresensiSholat(nis,sheet,bulan){
  var allData = [];
  
  var rangeData = sheet.getDataRange().getValues();
  var rowSiswa = 0;
  var filterPresensi = rangeData.filter(function(siswa,row){
    if(siswa[2]==nis){
    rowSiswa = row;
    return siswa;  
    }
  });
  
    var objek = {};
    
    for (var column = 2; column<38;column++){
      if(column<=6){   
      
        objek[rangeData[8][column]] = rangeData[rowSiswa][column];
      
      }else{
        objek[""+column-6] = rangeData[rowSiswa][column];
      }
       
      
         }
    objek["bulan"] = bulan;
    objek["denda"] = rangeData[rowSiswa][38]
    allData.push(objek);
      
  return allData;

}

function insertBerhalanganSholat(request,ss){
  var tingkatan = request.parameter.tingkatan;
  var sheet = ss.getSheetByName("Kelas "+tingkatan);
  var nis = request.parameter.nis;
  var kelas = request.parameter.kelas;
  var bulan = request.parameter.bulan;
  var tanggal = request.parameter.tanggal;
  var kodeRuang = request.parameter.kodeRuang;
  
  
  var data = sheet.getDataRange().getValues();
  
  var filterKelas = data.filter(function(kelasSiswa){
    if(kelasSiswa[1] == kelas){
      return kelasSiswa;
    }
  });
  
//  var jumlahSiswa = filterKelas[0][2];
  var linkKelas = filterKelas[0][3];
  var ssKelas = SpreadsheetApp.openByUrl(linkKelas);
  var sheetKelas = ssKelas.getSheetByName("Absen "+bulan);
  
  var kodeRuangdiDatabaseAsli = "BERHALANGAN"+""+tanggal+""+bulan;
  var  kodeRuangdiDatabase = Utilities.base64Encode(kodeRuangdiDatabaseAsli);
  
  
  if(kodeRuangdiDatabase !==kodeRuang){
       var output = JSON.stringify({"hasil":"kode ruangan salah"});
  
  }else if(kodeRuangdiDatabase ==kodeRuang){  
  var column = 0;
  switch(Number(tanggal)){
    case 1 :
      column = 8
      break;
    case 2 :
      column = 9
      break;
    case 3 :
      column = 10
      break;
    case 4 :
      column = 11
      break;
    case 5 :
      column = 12
      break;
    case 6 :
      column = 13
      break;
    case 7 :
      column = 14
      break;
    case 8 :
      column = 15
      break;
    case 9 :
      column = 16
      break;
    case 10 :
      column = 17
      break;
    case 11 :
      column = 18
      break;
    case 12 :
      column = 19
      break;
    case 13 :
      column = 20
      break;
    case 14 :
      column = 21
      break;
    case 15 :
      column = 22
      break;
    case 16 :
      column = 23
      break;
    case 17 :
      column = 24
      break;
    case 18 :
      column = 25
      break;
    case 19 :
      column = 26
      break;
    case 20 :
      column = 27
      break;
    case 21 :
      column = 28
      break;
    case 22 :
      column = 29
      break;
    case 23 :
      column = 30
      break;
    case 24 :
      column = 31
      break;
    case 25 :
      column = 32
      break;
    case 26 :
      column = 33
      break;
    case 27 :
      column = 34
      break;
    case 28 :
      column = 35
      break;
    case 29 :
      column = 36
      break;
    case 30 :
      column = 37
      break;  
    case 31 :
      column = 38
      break;      
  }
  
  var rangeData = sheetKelas.getDataRange().getValues();
  var rowSiswa = 0;
  var filterPresensi = rangeData.filter(function(siswa,row){
    if(siswa[2]==nis){
    rowSiswa = row;
    return siswa;  
    }
  });
  
  if(filterPresensi.length>0){
    sheetKelas.getRange(rowSiswa+1, column).setValue("M")
    var hasil = "sukses"
    }else{
    var hasil = "gagal"
    }
  
    var output = JSON.stringify({"hasil":hasil});
  }
  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function insertSholat(request,ss){
  var tingkatan = request.parameter.tingkatan;
  var sheet = ss.getSheetByName("Kelas "+tingkatan);
  var nis = request.parameter.nis;
  var kelas = request.parameter.kelas;
  var bulan = request.parameter.bulan;
  var tanggal = request.parameter.tanggal;
  var kodeRuang = request.parameter.kodeRuang;
  
  
  var data = sheet.getDataRange().getValues();
  
  var filterKelas = data.filter(function(kelasSiswa){
    if(kelasSiswa[1] == kelas){
      return kelasSiswa;
    }
  });
  
//  var jumlahSiswa = filterKelas[0][2];
  var linkKelas = filterKelas[0][3];
  var ssKelas = SpreadsheetApp.openByUrl(linkKelas);
  var sheetKelas = ssKelas.getSheetByName("Absen "+bulan);
  
  var kodeRuangdiDatabaseAsli = "DZUHUR"+""+tanggal+""+bulan;
  var  kodeRuangdiDatabase = Utilities.base64Encode(kodeRuangdiDatabaseAsli);
  
  
  if(kodeRuangdiDatabase !==kodeRuang){
       var output = JSON.stringify({"hasil":"kode ruangan salah"});
  
  }else if(kodeRuangdiDatabase ==kodeRuang){  
  var column = 0;
  switch(Number(tanggal)){
    case 1 :
      column = 8
      break;
    case 2 :
      column = 9
      break;
    case 3 :
      column = 10
      break;
    case 4 :
      column = 11
      break;
    case 5 :
      column = 12
      break;
    case 6 :
      column = 13
      break;
    case 7 :
      column = 14
      break;
    case 8 :
      column = 15
      break;
    case 9 :
      column = 16
      break;
    case 10 :
      column = 17
      break;
    case 11 :
      column = 18
      break;
    case 12 :
      column = 19
      break;
    case 13 :
      column = 20
      break;
    case 14 :
      column = 21
      break;
    case 15 :
      column = 22
      break;
    case 16 :
      column = 23
      break;
    case 17 :
      column = 24
      break;
    case 18 :
      column = 25
      break;
    case 19 :
      column = 26
      break;
    case 20 :
      column = 27
      break;
    case 21 :
      column = 28
      break;
    case 22 :
      column = 29
      break;
    case 23 :
      column = 30
      break;
    case 24 :
      column = 31
      break;
    case 25 :
      column = 32
      break;
    case 26 :
      column = 33
      break;
    case 27 :
      column = 34
      break;
    case 28 :
      column = 35
      break;
    case 29 :
      column = 36
      break;
    case 30 :
      column = 37
      break;  
    case 31 :
      column = 38
      break;      
  }
  
  var rangeData = sheetKelas.getDataRange().getValues();
  var rowSiswa = 0;
  var filterPresensi = rangeData.filter(function(siswa,row){
    if(siswa[2]==nis){
    rowSiswa = row;
    return siswa;  
    }
  });
  
  if(filterPresensi.length>0){
    sheetKelas.getRange(rowSiswa+1, column).setValue("0")
    var hasil = "sukses"
    }else{
    var hasil = "gagal"
    }
  
    var output = JSON.stringify({"hasil":hasil});
  }
  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);

}

function insertJamKe0(request,ss){
var tingkatan = request.parameter.tingkatan;
  var sheet = ss.getSheetByName("Kelas "+tingkatan);
  var nis = request.parameter.nis;
  var kelas = request.parameter.kelas;
  var bulan = request.parameter.bulan;
  var tanggal = request.parameter.tanggal;
  var kodeRuang = request.parameter.kodeRuang;
  var date = new Date();
  var menit = 0
  var minute = date.getMinutes();
  if(minute<10){
    menit = "0"+minute
  }else{
    menit = minute
  }
  var waktuPresensi = date.getHours()+":"+menit;
  
  var data = sheet.getDataRange().getValues();
  
  var filterKelas = data.filter(function(kelasSiswa){
    if(kelasSiswa[1] == kelas){
      return kelasSiswa;
    }
  });
  
//  var jumlahSiswa = filterKelas[0][2];
  var linkKelas = filterKelas[0][3];
  var ssKelas = SpreadsheetApp.openByUrl(linkKelas);
  var sheetKelas = ssKelas.getSheetByName("Absen "+bulan);
  
  var kodeRuangdiDatabaseAsli = "SKANSA"+""+tanggal+""+bulan;
  var  kodeRuangdiDatabase = Utilities.base64Encode(kodeRuangdiDatabaseAsli);
  
  
  if(kodeRuangdiDatabase !==kodeRuang){
       var output = JSON.stringify({"hasil":"kode ruangan salah"});
  
  }else if(kodeRuangdiDatabase ==kodeRuang){  
  var column = 0;
  switch(Number(tanggal)){
    case 1 :
      column = 8
      break;
    case 2 :
      column = 20
      break;
    case 3 :
      column = 32
      break;
    case 4 :
      column = 44
      break;
    case 5 :
      column = 56
      break;
    case 6 :
      column = 68
      break;
    case 7 :
      column = 80
      break;
    case 8 :
      column = 92
      break;
    case 9 :
      column = 104
      break;
    case 10 :
      column = 116
      break;
    case 11 :
      column = 128
      break;
    case 12 :
      column = 140
      break;
    case 13 :
      column = 152
      break;
    case 14 :
      column = 164
      break;
    case 15 :
      column = 176
      break;
    case 16 :
      column = 188
      break;
    case 17 :
      column = 200
      break;
    case 18 :
      column = 212
      break;
    case 19 :
      column = 224
      break;
    case 20 :
      column = 236
      break;
    case 21 :
      column = 248
      break;
    case 22 :
      column = 260
      break;
    case 23 :
      column = 272
      break;
    case 24 :
      column = 284
      break;
    case 25 :
      column = 296
      break;
    case 26 :
      column = 308
      break;
    case 27 :
      column = 220
      break;
    case 28 :
      column = 232
      break;
    case 29 :
      column = 244
      break;
    case 30 :
      column = 256
      break;  
    case 31 :
      column = 268
      break;   
  }
  
  var rangeData = sheetKelas.getDataRange().getValues();
  var rowSiswa = 0;
  var filterPresensi = rangeData.filter(function(siswa,row){
    if(siswa[2]==nis){
    rowSiswa = row;
    return siswa;  
    }
  });
  
  if(filterPresensi.length>0){
    sheetKelas.getRange(rowSiswa+1, column).setValue(waktuPresensi)
    var hasil = "sukses"
    }else{
    var hasil = "gagal"
    }
  
    var output = JSON.stringify({"hasil":hasil});
  }
  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function generateQr(request){
  var kodeRuang = request.parameter.kodeRuang;
  var output = HtmlService.createTemplateFromFile("qrcode");
  output.url = {first:"https://api.qrserver.com/v1/create-qr-code/?size=150x150&data="+kodeRuang};
  var htmlOutput = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle('QRCode Instansi');
    return htmlOutput
}

function readSPP(request,ss){
  var tingkatan = request.parameter.tingkatan;
  var sheet = ss.getSheetByName("Kelas "+tingkatan);
  var kelas = request.parameter.kelas;
  var nis = request.parameter.nis;
//  var kelas = "11 AP 1";
  
  
  
  var data = sheet.getDataRange().getValues();
  
  var filterKelas = data.filter(function(kelasSiswa){
    if(kelasSiswa[1] == kelas){
      return kelasSiswa;
    }
  });
  
  var linkKelas = filterKelas[0][2];
  var ssKelas = SpreadsheetApp.openByUrl(linkKelas);
  var sheetKelas = ssKelas.getSheetByName("spp");
  var flag = 0;
  var dataSPP = sheetKelas.getDataRange().getValues();
  
  var filterSPP = dataSPP.filter(function(spp){
    if(spp[0] == nis){
      flag = 1;
      return spp;
    }
  
  });
  
  if(flag == 1){
    var hasil = {
      "hasil": "succes",
      "juli" : filterSPP[0][2],
      "agustus": filterSPP[0][3],
      "september" : filterSPP[0][4],
      "oktober": filterSPP[0][5],
      "november" : filterSPP[0][6],
      "desember": filterSPP[0][7],
      "januari" : filterSPP[0][8],
      "februari": filterSPP[0][9],
      "maret" : filterSPP[0][10],
      "april": filterSPP[0][11],
      "mei" : filterSPP[0][12],
      "juni": filterSPP[0][13]
    }
  }else{
    var hasil = { "hasil" :"failed" }
  }
  
  var output = JSON.stringify(hasil);
  
  return ContentService
          .createTextOutput(output)
          .setMimeType(ContentService.MimeType.JSON);
  
}

function getContent(request, ss){
  var sheetName = request.parameter.sheetName;
  var sheet = ss.getSheetByName(sheetName);
  
  var data = {};
  data["results"] = content(sheet);
  
  var hasil = JSON.stringify(data);
  
  return ContentService
         .createTextOutput(hasil)
         .setMimeType(ContentService.MimeType.JSON);
}

function content(sheet){
 var allData = [];
  
  var rangeData = sheet.getDataRange().getValues();
  var jumlahKolom = sheet.getLastColumn();
  
  for(var row = 1; row<rangeData.length;row++){
    var objek = {};
    
    for (var column = 1; column<jumlahKolom;column++){
         objek[rangeData[0][column]] = rangeData[row][column];
         }
    allData.push(objek);
    
      }
      
  return allData;
}

function readPresensi(request, ss){
  
  var tingkatan = request.parameter.tingkatan;
  var sheet = ss.getSheetByName("Kelas "+tingkatan);
  var kelas = request.parameter.kelas;
//  var kelas = "11 AP 1";
  var tanggal = request.parameter.tanggal;
  var bulan = request.parameter.bulan;
  
  
  var data = sheet.getDataRange().getValues();
  
  var filterKelas = data.filter(function(kelasSiswa){
    if(kelasSiswa[1] == kelas){
      return kelasSiswa;
    }
  });
  
  var jumlahSiswa = filterKelas[0][2];
  var linkKelas = filterKelas[0][3];
  var ssKelas = SpreadsheetApp.openByUrl(linkKelas);
  var sheetKelas = ssKelas.getSheetByName("Absen "+bulan);
  
  var column = 0;
  switch(Number(tanggal)){
    case 1 :
      column = 8
      break;
    case 2 :
      column = 9
      break;
    case 3 :
      column = 10
      break;
    case 4 :
      column = 11
      break;
    case 5 :
      column = 12
      break;
    case 6 :
      column = 13
      break;
    case 7 :
      column = 14
      break;
    case 8 :
      column = 15
      break;
    case 9 :
      column = 16
      break;
    case 10 :
      column = 17
      break;
    case 11 :
      column = 18
      break;
    case 12 :
      column = 19
      break;
    case 13 :
      column = 20
      break;
    case 14 :
      column = 21
      break;
    case 15 :
      column = 22
      break;
    case 16 :
      column = 23
      break;
    case 17 :
      column = 24
      break;
    case 18 :
      column = 25
      break;
    case 19 :
      column = 26
      break;
    case 20 :
      column = 27
      break;
    case 21 :
      column = 28
      break;
    case 22 :
      column = 29
      break;
    case 23 :
      column = 30
      break;
    case 24 :
      column = 31
      break;
    case 25 :
      column = 32
      break;
    case 26 :
      column = 33
      break;
    case 27 :
      column = 34
      break;
    case 28 :
      column = 35
      break;
    case 29 :
      column = 36
      break;
    case 30 :
      column = 37
      break;  
    case 31 :
      column = 38
      break;   
  }
  
  var dataPresensiAll = sheetKelas.getRange(13, column, jumlahSiswa,1).getValues();
  var presensi = [];
  for(var i=0;i<jumlahSiswa;i++){
      var dataPresensi = dataPresensiAll[i][0];
      presensi.push(dataPresensi);
  }
  
  var output = JSON.stringify({"hasil":presensi});

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function readPresensiBulanan(request, ss){
  
  var tingkatan = request.parameter.tingkatan;
  var sheet = ss.getSheetByName("Kelas "+tingkatan);
  var nis = request.parameter.nis;
  var kelas = request.parameter.kelas;
  var bulan = request.parameter.bulan;
  
  
  var data = sheet.getDataRange().getValues();
  
  var filterKelas = data.filter(function(kelasSiswa){
    if(kelasSiswa[1] == kelas){
      return kelasSiswa;
    }
  });
  
//  var jumlahSiswa = filterKelas[0][2];
  var linkKelas = filterKelas[0][3];
  var ssKelas = SpreadsheetApp.openByUrl(linkKelas);
  var sheetKelas = ssKelas.getSheetByName("Absen "+bulan);
  
//  var ss = SpreadsheetApp.openById("1kyWZ0X9oiAVtR9ZeuwbZvbApx9btRXfoUgSZS9zvDr0");
//  var sheet = ss.getSheetByName("Absen JULI");
  
  
  var data = {};
  data["results"] = contentAbsensi(nis,sheetKelas, bulan);
  
  var output = JSON.stringify(data);

  return ContentService
       .createTextOutput(output)
       .setMimeType(ContentService.MimeType.JSON);
}

function contentAbsensi(nis,sheet,bulan){
 var allData = [];
  
  var rangeData = sheet.getDataRange().getValues();
  var rowSiswa = 0;
  var filterPresensi = rangeData.filter(function(siswa,row){
    if(siswa[2]==nis){
    rowSiswa = row;
    return siswa;  
    }
  });
  
//  for(var row = 12; row<12+jumlahSiswa;row++){
    var objek = {};
    
    for (var column = 2; column<38;column++){
      if(column<=6){   
      
        objek[rangeData[8][column]] = rangeData[rowSiswa][column];
      
      }else{
        objek[""+column-6] = rangeData[rowSiswa][column+372];
      }
       
      
         }
    objek["bulan"] = bulan;
    allData.push(objek);
    
//      }
      
  return allData;
}


function pesanWa(noHp,message){
  
  var myHeaders = {"Authorization":"69TI9tiTuRVglCNOzMTCmVB7diT9hWOf1X14PYJH3NvkUtBdz1gKm45tcRzvfPXN" }
  
  var formdata = {
    "phone" : ""+noHp,
    "message" : ""+message
  }

  var requestOptions = {
    method: 'POST',
    headers: myHeaders,
    'contentType':'application/json',
    'payload':JSON.stringify(formdata),
    redirect: 'follow'
};

  var response = UrlFetchApp.fetch("https://console.wablas.com/api/send-message", requestOptions);
  
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT)

}


function kolomPresensiKehadiran(tanggal){
  var column = 0;
  switch(Number(tanggal)){
    case 1 :
      column = 4
      break;
    case 2 :
      column = 7
      break;
    case 3 :
      column = 10
      break;
    case 4 :
      column = 13
      break;
    case 5 :
      column = 16
      break;
    case 6 :
      column = 19
      break;
    case 7 :
      column = 22
      break;
    case 8 :
      column = 25
      break;
    case 9 :
      column = 28
      break;
    case 10 :
      column = 31
      break;
    case 11 :
      column = 34
      break;
    case 12 :
      column = 37
      break;
    case 13 :
      column = 40
      break;
    case 14 :
      column = 43
      break;
    case 15 :
      column = 46
      break;
    case 16 :
      column = 49
      break;
    case 17 :
      column = 52
      break;
    case 18 :
      column = 55
      break;
    case 19 :
      column = 58
      break;
    case 20 :
      column = 61
      break;
    case 21 :
      column = 64
      break;
    case 22 :
      column = 67
      break;
    case 23 :
      column = 70
      break;
    case 24 :
      column = 73
      break;
    case 25 :
      column = 76
      break;
    case 26 :
      column = 79
      break;
    case 27 :
      column = 82
      break;
    case 28 :
      column = 85
      break;
    case 29 :
      column = 88
      break;
    case 30 :
      column = 91
      break;  
    case 31 :
      column = 94
      break;   
  }
  
  return column
}


function getMenit(date){
  var menit = 0
  var minute = date.getMinutes();
  if(minute<10){
    menit = "0"+minute
  }else{
    menit = minute
  }
  
  return menit  
}

function getJam(date){
  return date.getHours()

}

function waktuPresensi(){
  var date = new Date();
  var menit = 0
  var minute = date.getMinutes();
  if(minute<10){
    menit = "0"+minute
  }else{
    menit = minute
  }
  var waktuPresensi = date.getHours()+":"+menit;
return waktuPresensi
}

function getDay(){
  var date = new Date().getDay()
  return date
}


function kirimPesanWa(noHp,pesan){
    var formdata = {
    "device_id" : ""+idDevice,
    "number" : ""+noHp,
    "message" : ""+pesan
  }

  var requestOptions = {
    method: 'POST',
    'contentType':'application/json',
    'payload':JSON.stringify(formdata),
    redirect: 'follow'
};

  var response = UrlFetchApp.fetch("https://app.whacenter.com/api/send", requestOptions);
  Logger.log(response)

}
