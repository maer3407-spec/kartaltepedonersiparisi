/*
 * KARTALTEPE DÖNER ORGANİZASYON MERKEZİ — V2.2
 * Mevcut Kod.gs işlevleri korunarak genişletilmiştir.
 * Mevcut Siparisler ve Talebe_Ikram sütunları silinmez/yeniden sıralanmaz.
 * Kurulum: tablolariHazirlaV21() çalıştırın. Ardından admin.html dosyasını
 * Apps Script projesine HTML dosyası olarak ekleyin (dosya adı: admin).
 */

// ---------- MEVCUT V1 FONKSİYONLARI (korunmuş) ----------
function tablolariHazirla() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("HATA: Script bir E-Tabloya bağlı değil! E-Tablo içinden Uzantılar > Apps Script açın.");
  var satisSheet = ss.getSheetByName('Siparisler');
  if (!satisSheet) {
    satisSheet = ss.insertSheet('Siparisler');
    satisSheet.appendRow(['Kayıt Tarihi','1 Kg Paket Adedi','Yarım Kg Paket Adedi','Toplam Kilogram','Müşteri Ad Soyad','Müşteri Telefon','Teslimat Şekli','Teslimat Tarihi','Referans İsim Soyisim','Referans Telefon No','Teslimat İlçe','Teslimat Mahalle','Teslimat Açık Adres','Notlar']);
    satisSheet.getRange(1,1,1,14).setFontWeight('bold').setBackground('#4A5D3F').setFontColor('#ffffff');
  }
  var ikramSheet = ss.getSheetByName('Talebe_Ikram');
  if (!ikramSheet) {
    ikramSheet = ss.insertSheet('Talebe_Ikram');
    ikramSheet.appendRow(['Kayıt Tarihi','İkram Edilen Ürün','Hayırsever Ad Soyad','Telefon','Adet / Paket','Referans İsim Soyisim','Referans Telefon No','İkram / Dua Notu']);
    ikramSheet.getRange(1,1,1,8).setFontWeight('bold').setBackground('#2B4C65').setFontColor('#ffffff');
  }
}

function doGet(e) {
  var p = e && e.parameter ? e.parameter : {};
  if (p.action === 'publicConfig') {
    var cfg = getPublicConfig_();
    var json = JSON.stringify(cfg).replace(/</g,'\\u003c');
    if (p.callback && /^[A-Za-z_$][0-9A-Za-z_$\.]{0,80}$/.test(p.callback)) {
      return ContentService.createTextOutput(p.callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }
  if (p.action === 'admin') return HtmlService.createHtmlOutputFromFile('admin').setTitle('Döner Organizasyon Merkezi — Yönetim').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (p.kategori) {
    var sonuc = kaydiTabloyaYaz(p);
    return ContentService.createTextOutput(JSON.stringify(sonuc)).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Döner Sipariş & Talebe İkram Sistemi').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport','width=device-width, initial-scale=1, maximum-scale=1');
}

function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents;
    if (!raw) throw new Error('Veri paketi boş geldi.');
    var data = JSON.parse(raw);
    var result;
    if (data.action === 'create') result = kaydiTabloyaYaz(data.data || {});
    else if (data.action === 'trackOrder') result = publicTrackOrder_(data.id);
    else if (data.kategori) result = kaydiTabloyaYaz(data);
    else throw new Error('Geçersiz işlem.');
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:String(err.message || err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function kaydiTabloyaYaz(data) {
  if (!data || !data.kategori) throw new Error('Kayıt kategorisi eksik.');
  tablolariHazirlaV21();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tarih = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Istanbul', 'dd.MM.yyyy HH:mm:ss');
    var cfg = getSettings_();
    if (String(cfg.orderOpen) === 'false') throw new Error('Sipariş ve ikram alımı şu anda kapalıdır.');
    var kg = data.kategori === 'satis' ? (Number(data.adet1Kg)||0) + (Number(data.adetYarimKg)||0)*0.5 : ikramKg_(data);
    if (!(kg > 0)) throw new Error('Geçerli bir kilogram/paket miktarı girilmelidir.');
    var method = String(data.paymentMethod || '');
    if (method !== 'Nakit' && method !== 'Havale/EFT') throw new Error('Lütfen geçerli bir ödeme yöntemi seçin.');
    if (method === 'Nakit' && String(cfg.cashEnabled) === 'false') throw new Error('Nakit ödeme şu anda kapalıdır.');
    if (method === 'Havale/EFT' && String(cfg.bankEnabled) === 'false') throw new Error('Havale/EFT şu anda kapalıdır.');
    var used = getCommittedKg_();
    var capacity = Number(cfg.capacityKg) || 400;
    if (used + kg > capacity + 0.0001) throw new Error('Kapasite dolmuştur. Kalan miktar: ' + Math.max(0,capacity-used).toFixed(1) + ' kg');
    var id = makeId_(data.kategori === 'satis' ? 'S' : 'I');
    var price = Number(cfg.pricePerKg) || 0;
    var amount = Math.round(kg * price * 100) / 100;
    if (data.kategori === 'satis') {
      var sh = ss.getSheetByName(V21.orders);
      var row = [tarih,Number(data.adet1Kg)||0,Number(data.adetYarimKg)||0,kg,data.adSoyad||'',"'"+(data.telefon||''),data.teslimatSekli||'',data.teslimatTarihi||'',data.referansAd||'',"'"+(data.referansTel||''),data.ilce||'',data.mahalle||'',data.adres||'',data.notlar||''];
      row = row.concat([id,price,amount,'Yeni','Bekliyor',method]);
      sh.appendRow(row);
      appendLog_('Yeni sipariş',id,data.adSoyad||'');
    } else if (data.kategori === 'ikram') {
      var ish = ss.getSheetByName(V21.gifts);
      var ir = [tarih,data.urun||'',data.adSoyad||'',"'"+(data.telefon||''),Number(data.adet)||1,data.referansAd||'',"'"+(data.referansTel||''),data.not||''];
      ir = ir.concat([id,kg,'Yeni',price,amount,'Bekliyor',method]);
      ish.appendRow(ir);
      appendLog_('Yeni ücretli ikram',id,data.adSoyad||'');
    } else throw new Error('Geçersiz kayıt kategorisi.');
    return {status:'success',id:id,kg:kg,unitPrice:price,total:amount,paymentStatus:'Bekliyor',paymentMethod:method};
  } finally { lock.releaseLock(); }
}

// ---------- V2.2 KURULUM VE VERİ MODELİ ----------
var V21 = {
  settings:'Ayarlar', payments:'Odemeler', cuts:'Kesim_Plani', packing:'Paketleme',
  delivery:'Teslimat', staff:'Personel', logs:'Islem_Kaydi', sessions:'Oturumlar',
  orders:'Siparisler', gifts:'Talebe_Ikram'
};

function tablolariHazirlaV21() {
  tablolariHazirla();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureHeaders_(ss.getSheetByName(V21.orders), ['Sipariş ID','Kg Birim Fiyatı','Toplam Tutar','Sipariş Durumu','Ödeme Durumu','Ödeme Yöntemi']);
  ensureHeaders_(ss.getSheetByName(V21.gifts), ['İkram ID','Toplam Kg','İkram Durumu','Kg Birim Fiyatı','Toplam Tutar','Ödeme Durumu','Ödeme Yöntemi']);
  ensureSheet_(ss,V21.settings,['Ayar','Değer','Açıklama']);
  ensureSheet_(ss,V21.payments,['Ödeme ID','Sipariş ID','Tarih','Tutar','Yöntem','Durum','Kaydeden','Not']);
  ensureHeaders_(ss.getSheetByName(V21.payments), ['Kayıt Türü','İlgili Kayıt ID']);
  ensureSheet_(ss,V21.cuts,['Kesim ID','Kesim Tarihi','Planlanan Kg','Kesim Durumu','Not']);
  ensureHeaders_(ss.getSheetByName(V21.cuts), ['Oluşturma Tarihi']);
  ensureSheet_(ss,V21.packing,['Paket ID','Sipariş/İkram ID','Paket Tipi','Paket Adedi','Paketlenen Kg','Durum','Not']);
  ensureHeaders_(ss.getSheetByName(V21.packing), ['Kayıt Türü','Oluşturma Tarihi']);
  ensureSheet_(ss,V21.delivery,['Teslimat ID','Sipariş ID','Teslimat Tarihi','Teslimat Şekli','Adres/Alıcı','Personel','Durum','Not']);
  ensureHeaders_(ss.getSheetByName(V21.delivery), ['Kayıt Türü']);
  ensureSheet_(ss,V21.staff,['Personel ID','Ad Soyad','Kullanıcı Adı','Şifre Özeti','Rol','Aktif','Oluşturma Tarihi']);
  ensureSheet_(ss,V21.logs,['Tarih','İşlem','Kayıt ID','Detay','Kullanıcı']);
  ensureSheet_(ss,V21.sessions,['Oturum Özeti','Personel ID','Bitiş Zamanı']);
  seedSettings_();
  seedFirstAdmin_();
  ensureSheet_(ss,'Siparis_Durum_Gecmisi',['Tarih','Kayıt ID','Kayıt Türü','Eski Durum','Yeni Durum','İşlemi Yapan']);
  return 'V2.2 tabloları hazır. Mevcut kayıtlar korunmuştur.';
}

function ensureSheet_(ss,name,headers) {
  var sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastRow()===0) sh.appendRow(headers);
  sh.setFrozenRows(1); sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#4A5D3F').setFontColor('#ffffff');
  return sh;
}
function ensureHeaders_(sh,headers) {
  var last = Math.max(sh.getLastColumn(),1); var existing = sh.getRange(1,1,1,last).getValues()[0];
  headers.forEach(function(h){ if(existing.indexOf(h)===-1){ last++; sh.getRange(1,last).setValue(h); existing.push(h); } });
}
function appendByHeaders_(sh,obj){var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];var row=h.map(function(k){return Object.prototype.hasOwnProperty.call(obj,k)?obj[k]:'';});sh.appendRow(row);}
function seedSettings_() {
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.settings);
  var defaults = [
    ['pricePerKg','850','Döner kilogram satış fiyatı (TL)'],['capacityKg','400','Toplam fiziksel kapasite (kg)'],
    ['orderOpen','true','Yeni sipariş alımı açık mı?'],['accountName','KARTALTEPE KURS VE OKUL TALEBELERİNE YARDIM DERNEĞİ','Hesap sahibi'],
    ['iban','TR8100205000008770097800001','İkram IBAN bilgisi'],['pickupAddress','Kartaltepe Mah. Sefalı Sok. No:2, Bakırköy / İstanbul','Gel-al adresi'],
    ['contactPhone','05345709807','İletişim telefonu'],['deliveryEnabled','true','Adrese teslim seçeneği'],['pickupEnabled','true','Gel-al seçeneği'],
    ['cashEnabled','true','Nakit ödeme'],['bankEnabled','true','Havale/EFT ödeme'],['salesCapacityKg','400','Satış için hedef kapasite (bilgilendirme)'],
    ['publicNotice','Sipariş ve ikram kayıtları alınmaktadır.','Müşteri ekranı duyurusu']
  ];
  var vals=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,2).getValues():[]; var keys=vals.map(function(r){return String(r[0]);});
  defaults.forEach(function(r){if(keys.indexOf(r[0])===-1) sh.appendRow(r);});
}
function seedFirstAdmin_() {
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.staff);
  if(sh.getLastRow()>1) return;
  var username='yonetici'; var password=Utilities.getUuid().replace(/-/g,'').slice(0,12);
  sh.appendRow([makeId_('P'),'Sistem Yöneticisi',username,hashPassword_(username,password),'admin',true,new Date()]);
  Logger.log('V2.1 ilk yönetici kullanıcı adı: '+username+' | İlk şifre: '+password+' | Giriş: Web uygulaması URL sonuna ?action=admin ekleyin.');
}

// ---------- GENEL VERİ / KAPASİTE ----------
function getSettings_() {
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.settings); var out={};
  if(!sh || sh.getLastRow()<2) return out;
  sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(function(r){out[String(r[0])]=r[1];}); return out;
}
function getPublicConfig_() {
  tablolariHazirlaV21(); var s=getSettings_(); var cap=Number(s.capacityKg)||400; var used=getCommittedKg_();
  return {pricePerKg:Number(s.pricePerKg)||0,capacityKg:cap,usedKg:used,remainingKg:Math.max(0,cap-used),orderOpen:String(s.orderOpen)!=='false',accountName:s.accountName||'',iban:s.iban||'',pickupAddress:s.pickupAddress||'',contactPhone:s.contactPhone||'',deliveryEnabled:String(s.deliveryEnabled)!=='false',pickupEnabled:String(s.pickupEnabled)!=='false',cashEnabled:String(s.cashEnabled)!=='false',bankEnabled:String(s.bankEnabled)!=='false',publicNotice:s.publicNotice||''};
}
function ikramKg_(d) { var n=Number(d.adet)||1; var u=String(d.urun||''); return n*(u.indexOf('Yarım')!==-1?0.5:1); }
function getCommittedKg_() {
  var ss=SpreadsheetApp.getActiveSpreadsheet(), total=0;
  var sh=ss.getSheetByName(V21.orders); if(sh && sh.getLastRow()>1){var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];var c=h.indexOf('Toplam Kilogram');var st=h.indexOf('Sipariş Durumu');var rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();rows.forEach(function(r){var status=st>=0?String(r[st]):'';if(status!=='İptal Edildi')total+=Number(r[c])||0;});}
  var ig=ss.getSheetByName(V21.gifts); if(ig && ig.getLastRow()>1){var ih=ig.getRange(1,1,1,ig.getLastColumn()).getValues()[0];var kg=ih.indexOf('Toplam Kg');var st2=ih.indexOf('İkram Durumu');var prod=ih.indexOf('İkram Edilen Ürün');var qty=ih.indexOf('Adet / Paket');var rr=ig.getRange(2,1,ig.getLastRow()-1,ig.getLastColumn()).getValues();rr.forEach(function(r){if(st2>=0&&String(r[st2])==='İptal Edildi')return;var n=Number(r[kg])||0;if(!n&&prod>=0&&qty>=0)n=(Number(r[qty])||1)*(String(r[prod]).indexOf('Yarım')!==-1?0.5:1);total+=n;});}
  return total;
}
function makeId_(prefix){return prefix+'-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Europe/Istanbul','yyMMdd-HHmmss')+'-'+Math.floor(100000+Math.random()*900000);}
function appendLog_(action,id,detail,user){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.logs);if(sh)sh.appendRow([new Date(),action,id,detail||'',user||'Sistem']);}

// ---------- OTURUM / YETKİLENDİRME ----------
function hashPassword_(user,password){var bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(user)+'|doner-v21|' + String(password),Utilities.Charset.UTF_8);return bytes.map(function(b){return ('0'+((b+256)%256).toString(16)).slice(-2);}).join('');}
function adminLogin(username,password){tablolariHazirlaV21();var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.staff);var rows=sh.getDataRange().getValues();for(var i=1;i<rows.length;i++){var r=rows[i];if(String(r[2])===String(username)&&String(r[3])===hashPassword_(username,password)&&String(r[5])!=='false'&&r[5]!==false){var token=Utilities.getUuid()+Utilities.getUuid();var exp=Date.now()+6*60*60*1000;CacheService.getScriptCache().put('v21_'+token,JSON.stringify({id:r[0],name:r[1],role:r[4],username:r[2],exp:exp}),21600);return {ok:true,token:token,name:r[1],role:r[4]};}}throw new Error('Kullanıcı adı veya şifre hatalı.');}
function requireAdmin_(token,allowed){var raw=CacheService.getScriptCache().get('v21_'+String(token||''));if(!raw)throw new Error('Oturum süresi doldu. Yeniden giriş yapın.');var u=JSON.parse(raw);if(u.exp<Date.now())throw new Error('Oturum süresi doldu. Yeniden giriş yapın.');if(allowed&&allowed.indexOf(u.role)===-1)throw new Error('Bu işlem için yetkiniz bulunmuyor.');return u;}
function adminLogout(token){CacheService.getScriptCache().remove('v21_'+String(token||''));return {ok:true};}

// ---------- YÖNETİM PANELİ API ----------
function adminGetDashboard(token){requireAdmin_(token,['admin','manager','finance','operations']);var ss=SpreadsheetApp.getActiveSpreadsheet();var cfg=getPublicConfig_();var orders=readObjects_(ss.getSheetByName(V21.orders));var gifts=readObjects_(ss.getSheetByName(V21.gifts));var pay=readObjects_(ss.getSheetByName(V21.payments));var totalSales=orders.filter(function(r){return r['Sipariş Durumu']!=='İptal Edildi';}).reduce(function(a,r){return a+(Number(r['Toplam Kilogram'])||0);},0);var giftKg=gifts.filter(function(r){return r['İkram Durumu']!=='İptal Edildi';}).reduce(function(a,r){return a+(Number(r['Toplam Kg'])||0);},0);var collected=pay.filter(function(r){return String(r['Durum'])!=='İptal';}).reduce(function(a,r){return a+(Number(r['Tutar'])||0);},0);return {capacity:cfg.capacityKg,used:cfg.usedKg,remaining:cfg.remainingKg,salesKg:totalSales,giftKg:giftKg,orderCount:orders.length,giftCount:gifts.length,collected:collected,pricePerKg:cfg.pricePerKg,orderOpen:cfg.orderOpen,recentOrders:orders.slice(-8).reverse(),recentGifts:gifts.slice(-5).reverse()};}
function adminGetSettings(token){requireAdmin_(token,['admin','manager']);return getSettings_();}
function adminSaveSettings(token,values){var u=requireAdmin_(token,['admin','manager']);if(!values||typeof values!=='object')throw new Error('Ayar verisi geçersiz.');var allowed=['pricePerKg','capacityKg','orderOpen','accountName','iban','pickupAddress','contactPhone','deliveryEnabled','pickupEnabled','cashEnabled','bankEnabled','publicNotice','salesCapacityKg'];var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.settings);var map={};if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(function(r,i){map[String(r[0])]=i+2;});allowed.forEach(function(k){if(Object.prototype.hasOwnProperty.call(values,k)){var row=map[k];if(row)sh.getRange(row,2).setValue(values[k]);else sh.appendRow([k,values[k],'']);}});appendLog_('Ayarlar güncellendi','',Object.keys(values).join(', '),u.name);return {ok:true,settings:getSettings_()};}
function adminListOrders(token){requireAdmin_(token,['admin','manager','finance','operations']);return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.orders)).reverse();}
function adminListIkram(token){requireAdmin_(token,['admin','manager','operations','finance']);return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.gifts)).reverse();}
function adminUpdateOrder(token,id,patch){var u=requireAdmin_(token,['admin','manager','operations']);var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.orders);var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];var idCol=h.indexOf('Sipariş ID');if(idCol<0)throw new Error('Sipariş ID sütunu bulunamadı.');var rows=sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),sh.getLastColumn()).getValues();for(var i=0;i<rows.length;i++){if(String(rows[i][idCol])===String(id)){var statusCol=h.indexOf('Sipariş Durumu');var oldStatus=statusCol>=0?rows[i][statusCol]:'';Object.keys(patch||{}).forEach(function(k){var c=h.indexOf(k);if(c>=0)sh.getRange(i+2,c+1).setValue(patch[k]);});if(patch&&patch['Sipariş Durumu']&&String(oldStatus)!==String(patch['Sipariş Durumu']))SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Siparis_Durum_Gecmisi').appendRow([new Date(),id,'Satış',oldStatus,patch['Sipariş Durumu'],u.name]);appendLog_('Sipariş güncellendi',id,JSON.stringify(patch),u.name);return {ok:true};}}throw new Error('Sipariş bulunamadı.');}
function adminAddPayment(token,p){var u=requireAdmin_(token,['admin','manager','finance']);if(!p||!p.orderId||!(Number(p.amount)>0))throw new Error('Sipariş/ikram ID ve geçerli ödeme tutarı gerekli.');var ss=SpreadsheetApp.getActiveSpreadsheet();var id=String(p.orderId).trim().toUpperCase();var isGift=id.indexOf('I-')===0;var target=ss.getSheetByName(isGift?V21.gifts:V21.orders);var headers=target.getRange(1,1,1,target.getLastColumn()).getValues()[0];var idCol=headers.indexOf(isGift?'İkram ID':'Sipariş ID');var amountCol=headers.indexOf('Toplam Tutar');var statusCol=headers.indexOf('Ödeme Durumu');var rows=target.getDataRange().getValues();var rowNo=-1,total=0;for(var i=1;i<rows.length;i++){if(String(rows[i][idCol])===id){rowNo=i+1;total=Number(rows[i][amountCol])||0;break;}}if(rowNo<0)throw new Error('Sipariş/ikram kaydı bulunamadı.');var paySh=ss.getSheetByName(V21.payments);var payData=paySh.getDataRange().getValues();var paid=0;for(var j=1;j<payData.length;j++){var pid=String(payData[j][1]||'');var st=String(payData[j][5]||'');if(pid===id&&st!=='İptal')paid+=Number(payData[j][3])||0;}var nextPaid=paid+Number(p.amount);if(total>0&&nextPaid>total+0.01)throw new Error('Girilen ödeme kalan borçtan fazla. Kalan: '+Math.max(0,total-paid).toFixed(2)+' TL');var row=[makeId_('OD'),id,new Date(),Number(p.amount),p.method||'Havale/EFT','Alındı',u.name,p.note||'',isGift?'İkram':'Satış',id];paySh.appendRow(row);target.getRange(rowNo,statusCol+1).setValue(nextPaid>=total?'Ödendi':'Kısmi Ödendi');appendLog_('Ödeme alındı',id,String(p.amount)+' TL · '+(isGift?'İkram':'Satış'),u.name);return {ok:true,paid:nextPaid,remaining:Math.max(0,total-nextPaid)};}
function adminListStaff(token){requireAdmin_(token,['admin']);var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.staff);var rows=sh.getDataRange().getValues();return rows.slice(1).map(function(r){return {id:r[0],name:r[1],username:r[2],role:r[4],active:r[5]};});}
function adminAddStaff(token,p){var u=requireAdmin_(token,['admin']);if(!p||!p.name||!p.username||!p.password)throw new Error('Ad, kullanıcı adı ve ilk şifre zorunludur.');if(String(p.password).length<8)throw new Error('Şifre en az 8 karakter olmalıdır.');var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.staff);var vals=sh.getDataRange().getValues();if(vals.some(function(r,i){return i>0&&String(r[2])===String(p.username);}))throw new Error('Bu kullanıcı adı zaten kullanılıyor.');sh.appendRow([makeId_('P'),p.name,p.username,hashPassword_(p.username,p.password),p.role||'operations',true,new Date()]);appendLog_('Personel eklendi','',p.username,u.name);return {ok:true};}
function adminSetStaffActive(token,id,active){var u=requireAdmin_(token,['admin']);var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.staff);var d=sh.getDataRange().getValues();for(var i=1;i<d.length;i++){if(String(d[i][0])===String(id)){sh.getRange(i+1,6).setValue(!!active);appendLog_('Personel durumu',id,String(active),u.name);return {ok:true};}}throw new Error('Personel bulunamadı.');}
function adminListPayments(token){requireAdmin_(token,['admin','manager','finance']);return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.payments)).reverse();}
function readObjects_(sh){if(!sh||sh.getLastRow()<2)return [];var v=sh.getDataRange().getValues(),h=v[0];return v.slice(1).filter(function(r){return r.some(function(x){return x!==''&&x!==null;});}).map(function(r){var o={};h.forEach(function(k,i){o[String(k)]=r[i] instanceof Date?Utilities.formatDate(r[i],Session.getScriptTimeZone()||'Europe/Istanbul','dd.MM.yyyy HH:mm'):r[i];});return o;});}


// ================= V2.2: TAHSİLAT, SİPARİŞ TAKİBİ VE OPERASYON =================
function publicTrackOrder_(rawId) {
  var id=String(rawId||'').trim().toUpperCase();
  if(!/^([SI]-[0-9]{6}-[0-9]{6}-(?:[0-9]{3}|[0-9]{6}))$/.test(id)) throw new Error('Sipariş numarası biçimi geçersiz.');
  tablolariHazirlaV21();
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var order=readObjects_(ss.getSheetByName(V21.orders)).filter(function(r){return String(r['Sipariş ID']||'').toUpperCase()===id;})[0];
  var gift=readObjects_(ss.getSheetByName(V21.gifts)).filter(function(r){return String(r['İkram ID']||'').toUpperCase()===id;})[0];
  var r=order||gift; if(!r) throw new Error('Bu numarayla eşleşen kayıt bulunamadı. Numaranızı kontrol edin.');
  var isGift=!!gift;
  return {ok:true,id:id,type:isGift?'Talebe ikramı':'Döner siparişi',createdAt:r['Kayıt Tarihi']||'',kg:Number(isGift?r['Toplam Kg']:r['Toplam Kilogram'])||0,total:Number(r['Toplam Tutar'])||0,paymentStatus:r['Ödeme Durumu']||'Bekliyor',paymentMethod:r['Ödeme Yöntemi']||'',status:isGift?(r['İkram Durumu']||'Yeni'):(r['Sipariş Durumu']||'Yeni'),deliveryDate:isGift?'':(r['Teslimat Tarihi']||''),deliveryType:isGift?'':(r['Teslimat Şekli']||'')};
}
function adminUpdateGift(token,id,patch){var u=requireAdmin_(token,['admin','manager','operations']);var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.gifts);var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];var idc=h.indexOf('İkram ID');var rows=sh.getDataRange().getValues();for(var i=1;i<rows.length;i++){if(String(rows[i][idc])===String(id)){var old=h.indexOf('İkram Durumu')>=0?rows[i][h.indexOf('İkram Durumu')]:'';Object.keys(patch||{}).forEach(function(k){var c=h.indexOf(k);if(c>=0)sh.getRange(i+1,c+1).setValue(patch[k]);});if(patch&&patch['İkram Durumu']&&String(old)!==String(patch['İkram Durumu']))SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Siparis_Durum_Gecmisi').appendRow([new Date(),id,'İkram',old,patch['İkram Durumu'],u.name]);appendLog_('İkram durumu güncellendi',id,JSON.stringify(patch),u.name);return {ok:true};}}throw new Error('İkram kaydı bulunamadı.');}
function adminListCuts(token){requireAdmin_(token,['admin','manager','operations']);return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.cuts)).reverse();}
function adminAddCut(token,p){var u=requireAdmin_(token,['admin','manager','operations']);if(!p||!p.date||!(Number(p.kg)>0))throw new Error('Kesim tarihi ve planlanan kg zorunludur.');var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.cuts);var id=makeId_('K');appendByHeaders_(sh,{'Kesim ID':id,'Kesim Tarihi':p.date,'Planlanan Kg':Number(p.kg),'Kesim Durumu':'Planlandı','Not':p.note||'','Oluşturma Tarihi':new Date()});appendLog_('Kesim planı eklendi',id,p.date+' · '+p.kg+' kg',u.name);return {ok:true,id:id};}
function adminUpdateCut(token,id,status,note){var u=requireAdmin_(token,['admin','manager','operations']);var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.cuts);var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],d=sh.getDataRange().getValues(),c=h.indexOf('Kesim ID');for(var i=1;i<d.length;i++){if(String(d[i][c])===String(id)){sh.getRange(i+1,h.indexOf('Kesim Durumu')+1).setValue(status);if(note!==undefined)sh.getRange(i+1,h.indexOf('Not')+1).setValue(note);appendLog_('Kesim planı güncellendi',id,status,u.name);return {ok:true};}}throw new Error('Kesim planı bulunamadı.');}
function adminListPacking(token){requireAdmin_(token,['admin','manager','operations']);return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.packing)).reverse();}
function adminAddPacking(token,p){var u=requireAdmin_(token,['admin','manager','operations']);if(!p||!p.recordId||!(Number(p.count)>0)||!(Number(p.kg)>0))throw new Error('Kayıt ID, paket adedi ve kg zorunludur.');var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.packing);var id=makeId_('PK');appendByHeaders_(sh,{'Paket ID':id,'Sipariş/İkram ID':p.recordId,'Kayıt Türü':p.recordType||'Satış','Paket Tipi':p.packType||'1 Kg','Paket Adedi':Number(p.count),'Paketlenen Kg':Number(p.kg),'Durum':'Paketlendi','Not':p.note||'','Oluşturma Tarihi':new Date()});appendLog_('Paketleme kaydı',id,p.recordId+' · '+p.kg+' kg',u.name);return {ok:true,id:id};}
function adminUpdatePacking(token,id,status,note){var u=requireAdmin_(token,['admin','manager','operations']);var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.packing);var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],d=sh.getDataRange().getValues(),c=h.indexOf('Paket ID');for(var i=1;i<d.length;i++){if(String(d[i][c])===String(id)){sh.getRange(i+1,h.indexOf('Durum')+1).setValue(status);if(note!==undefined)sh.getRange(i+1,h.indexOf('Not')+1).setValue(note);appendLog_('Paketleme durumu',id,status,u.name);return {ok:true};}}throw new Error('Paketleme kaydı bulunamadı.');}
function adminListDelivery(token){requireAdmin_(token,['admin','manager','operations']);return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.delivery)).reverse();}
function adminAddDelivery(token,p){var u=requireAdmin_(token,['admin','manager','operations']);if(!p||!p.recordId||!p.date||!p.type)throw new Error('Kayıt ID, tarih ve teslimat şekli zorunludur.');var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.delivery);var id=makeId_('TS');appendByHeaders_(sh,{'Teslimat ID':id,'Sipariş ID':p.recordId,'Kayıt Türü':p.recordType||'Satış','Teslimat Tarihi':p.date,'Teslimat Şekli':p.type,'Adres/Alıcı':p.address||'','Personel':p.staff||'','Durum':'Planlandı','Not':p.note||''});appendLog_('Teslimat planlandı',id,p.recordId+' · '+p.date,u.name);return {ok:true,id:id};}
function adminUpdateDelivery(token,id,status){var u=requireAdmin_(token,['admin','manager','operations']);var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(V21.delivery);var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],d=sh.getDataRange().getValues(),c=h.indexOf('Teslimat ID');for(var i=1;i<d.length;i++){if(String(d[i][c])===String(id)){sh.getRange(i+1,h.indexOf('Durum')+1).setValue(status);appendLog_('Teslimat durumu',id,status,u.name);return {ok:true};}}throw new Error('Teslimat kaydı bulunamadı.');}
