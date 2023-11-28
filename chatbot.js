var token = 
"xxx";
var telegramUrl =  "https://api.telegram.org/bot"+token;
var webAppUrl = "xxx";
var ssid = "xxx";

const sheetSet = SpreadsheetApp.openById(ssid).getSheetByName("Sheet1")

function setWebhook(){
var url = telegramUrl + "/setWebhook?url="  +webAppUrl;
var resp = UrlFetchApp.fetch(url);
Logger.log(resp)  
}

function cariPanjangKolomDenganNilai() {
  var spreadsheet = SpreadsheetApp.openById(ssid);
  var sheet = spreadsheet.getSheetByName("Sheet1");
  var dataRange = sheet.getRange(1, 1, sheet.getLastRow(), 1);
  var values = dataRange.getValues();
  var panjangKolomDenganNilai = 0;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] !== "") {
      panjangKolomDenganNilai++;
    }
  }
  
  return panjangKolomDenganNilai;
}

function sendMessage(id, text, replymarkup) {
var data = {
    method : "post",
    payload : {
        method : "sendMessage",
        chat_id : String(id),
        text : text,
        parse_mode : "HTML",
        reply_markup : JSON.stringify(replymarkup)
    }
}    
UrlFetchApp.fetch("https://api.telegram.org/bot"+token+"/",data);
}

function doPost(e){
    var updates = JSON.parse(e.postData.contents);
    var id = updates.message.chat.id;
    var username = updates.message.from.username;
    var text = updates.message.text;
    var nama = updates.message.chat.first_name;

    if(text=='/start'){
    var answer = "ID ANDA : " +id+"\nNama Anda :" + nama+"\nUsername : "+username;
    sendMessage(id,answer);
}

if(text.indexOf("/template_input")==0){
sendMessage(id, "ID LOP=\
                \nSEGMEN=\
                \nWITEL=\
                \nNAMA CC=\
                \nJUDUL PROJECT=\
                \nNILAI PROJECT=\
                \nMETODE PENGADAAN=\
                \nMEKANISME=\
                \nANPRUS/ MITRA=\
                \nPRODUCT CATEGORY=\
                \nPORTOFOLIO=\
                \nSUB PROJECT=\
                \nDIGITAL 7+3/ DIGITAL NON 7+3)/ NON DIGITAL=\
                \nKLASIFIKASI DIGITAL/ NON DIGITAL=\
                \nSTATUS PROJECT=\
                \nSALES FUNNEL=\
                \nKET SALES FUNNEL=\
                \nSBR/TIDAK SBR=\
                \nBESARAN DISKON (SBR)=\
                \nPOSISI SBR=\
                \nTGL_KET SALES FUNNEL=\
                ")
}

if(text.indexOf("/template_update")==0){
sendMessage(id, "NOMOR ANTRIAN=\
                \nUpdate=\
                \nValue=\
                ")
}  

function searchRow(antrian) {
  const spreadsheet = SpreadsheetApp.openById(ssid);
  const sheet = spreadsheet.getSheetByName("Sheet1");
  const dataRange = sheet.getRange(1, 27, sheet.getLastRow(), 1);
  const values = dataRange.getValues();
  let targetRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == antrian) {
      targetRow = i + 1;
      break;
    }
  }

  return targetRow;
}

function dataFound(argRow, argCol, argVal) {
  const worksheet = SpreadsheetApp.openById(ssid).getSheetByName("Sheet1")
  worksheet.getRange(argRow, argCol).setValue(argVal);

  return "Data berhasil diperbarui"
}

if(text.indexOf("/update_status_project")==0) {
  const command = text.split("\n").slice(0,1).join("")
  const noAntrian = text.split("\n").slice(1,2).join("").split("Nomor Antrian=").slice(1,2).join("")
  const isValue = text.split("\n").slice(2,3).join("").split("Status Project=").slice(1,2).join("")

  const row = searchRow(noAntrian.replace(/\s/g, ''))

  const balikan = row === -1 ? "Nomor antrian tidak ditemukan" : dataFound(row, 18, isValue);

  sendMessage(id, balikan)
}

if(text.indexOf("/update_sales_funnel")==0) {
  const command = text.split("\n").slice(0,1).join("")
  const noAntrian = text.split("\n").slice(1,2).join("").split("Nomor Antrian=").slice(1,2).join("")
  const isValue = text.split("\n").slice(2,3).join("").split("Sales Funnel=").slice(1,2).join("")

  const row = searchRow(noAntrian.replace(/\s/g, ''))

  const balikan = row === -1 ? "Nomor antrian tidak ditemukan" : dataFound(row, 19, isValue);

  sendMessage(id,balikan)
}

if(text.indexOf("/update_Ket_Sales_funnel")==0) {
  const command = text.split("\n").slice(0,1).join("")
  const noAntrian = text.split("\n").slice(1,2).join("").split("Nomor Antrian=").slice(1,2).join("")
  const isValue = text.split("\n").slice(2,3).join("").split("Ket Sales Funnel=").slice(1,2).join("")

  const row = searchRow(noAntrian.replace(/\s/g, ''))

  const balikan = row === -1 ? "Nomor antrian tidak ditemukan" : dataFound(row, 20, isValue);

  sendMessage(id, balikan)
}


if(text.indexOf("/input")==0){
var aa = text.split("\n").slice(0,1).join("")
var bb = text.split("\n").slice(1,2).join("").split("ID LOP=").slice(1,2).join("")
var cc = text.split("\n").slice(2,3).join("").split("SEGMEN=").slice(1,2).join("")
var dd = text.split("\n").slice(3,4).join("").split("WITEL=").slice(1,2).join("")
var ee = text.split("\n").slice(4,5).join("").split("NAMA CC=").slice(1,2).join("")
var ff = text.split("\n").slice(5,6).join("").split("JUDUL PROJECT=").slice(1,2).join("")
var gg = text.split("\n").slice(6,7).join("").split("NILAI PROJECT=").slice(1,2).join("")
var hh = text.split("\n").slice(7,8).join("").split("METODE PENGADAAN=").slice(1,2).join("")
var ii = text.split("\n").slice(8,9).join("").split("MEKANISME=").slice(1,2).join("")
var jj = text.split("\n").slice(9,10).join("").split("ANPRUS/ MITRA=").slice(1,2).join("")
var kk = text.split("\n").slice(10,11).join("").split("PRODUCT CATEGORY=").slice(1,2).join("")
var ll = text.split("\n").slice(11,12).join("").split("PORTOFOLIO=").slice(1,2).join("")
var mm = text.split("\n").slice(12,13).join("").split("SUB PROJECT=").slice(1,2).join("")
var nn = text.split("\n").slice(13,14).join("").split("DIGITAL 7+3/ DIGITAL NON 7+3)/ NON DIGITAL=").slice(1,2).join("")
var oo = text.split("\n").slice(14,15).join("").split("KLASIFIKASI DIGITAL/ NON DIGITAL=").slice(1,2).join("")
var pp = text.split("\n").slice(15,16).join("").split("STATUS PROJECT=").slice(1,2).join("")
var qq = text.split("\n").slice(16,17).join("").split("SALES FUNNEL=").slice(1,2).join("")
var rr = text.split("\n").slice(17,18).join("").split("KET SALES FUNNEL=").slice(1,2).join("")
var ss = text.split("\n").slice(18,19).join("").split("SBR/TIDAK SBR=").slice(1,2).join("")
var tt = text.split("\n").slice(19,20).join("").split("BESARAN DISKON (SBR)=").slice(1,2).join("")
var uu = text.split("\n").slice(20,21).join("").split("POSISI SBR=").slice(1,2).join("")
const antrian = "TIC" + cariPanjangKolomDenganNilai()

var now = new Date()
var waktu = Utilities.formatDate(now, "Asia/Jakarta","dd/MM/yyyy hh:mm:ss")
//var ans = aa + bb + cc + dd + ee

sendMessage(id,"data anda pada "+waktu+" sudah berhasil ditambahkan dengan nomor tiket: " + antrian)
SpreadsheetApp.openById(ssid).getSheetByName("Sheet1").appendRow([waktu, username, nama, bb,cc,dd,ee, ff, gg, hh, ii, jj, kk, ll, mm, nn, oo, pp, qq,rr,ss,tt,uu,"","","", antrian])
}

}
