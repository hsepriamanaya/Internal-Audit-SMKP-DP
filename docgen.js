// Generator Word (.docx) di browser/node, meniru layout draft Lampiran. Butuh library `docx` (global di browser, require di node).
(function(root){
const D=typeof require!=='undefined'&&typeof module!=='undefined'?require('docx'):root.docx;
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,AlignmentType,BorderStyle,PageBreak,VerticalAlign,ShadingType}=D;
const FONT='Times New Roman',SZ=24; // 12pt
const LETTER={width:12240,height:15840}; // 21.59 x 27.94 cm
const MARG={top:1440,bottom:1440,left:1440,right:1440}; // 2.54 cm
const CW=12240-2880; // lebar konten (DXA)
const B={style:BorderStyle.SINGLE,size:4,color:'000000'}, BR={top:B,bottom:B,left:B,right:B};
const NB={style:BorderStyle.NONE,size:0,color:'FFFFFF'}, NBR={top:NB,bottom:NB,left:NB,right:NB};
function run(t,o){return new TextRun(Object.assign({text:String(t==null?'':t),font:FONT,size:SZ},o||{}))}
function para(t,o){o=o||{};const lines=String(t==null?'':t).split('\n');
  return new Paragraph({alignment:o.al||AlignmentType.LEFT,spacing:{after:o.after==null?60:o.after,before:o.before||0},keepNext:o.keepNext,
    children:lines.flatMap((l,i)=>[run(l,{bold:o.b,italics:o.i,size:o.size}),...(i<lines.length-1?[new TextRun({break:1})]:[])])})}
function cell(content,w,o){o=o||{};const kids=(Array.isArray(content)?content:[content]).map(c=>typeof c==='string'||typeof c==='number'?para(c,{b:o.b,al:o.al}):c);
  return new TableCell({children:kids,width:{size:w,type:WidthType.DXA},borders:o.nb?NBR:BR,columnSpan:o.span,verticalAlign:o.va||VerticalAlign.TOP,margins:{top:50,bottom:50,left:90,right:90},
    shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined})}
function table(rows,widths){return new Table({width:{size:widths.reduce((a,b)=>a+b,0),type:WidthType.DXA},columnWidths:widths,rows:rows.map(r=>new TableRow({children:r,cantSplit:true}))})}
function head(txt){return para(txt,{b:true,al:AlignmentType.CENTER,after:200})}
function kv(k,v){return new Paragraph({spacing:{after:40},tabStops:[{type:'left',position:2700}],children:[run(k,{bold:true}),run('\t: '+(v||''))]})}
function checkRow(kat){const m={KRITIKAL:'Kritikal',MAYOR:'Major',MINOR:'Minor'};
  return ['Kritikal','Major','Minor'].map(x=>para((m[kat]===x?'☒ ':'☐ ')+x,{after:40}))}
function bullets(a){return (a||[]).length?a.map(x=>para('- '+x,{after:30})):[para('')]}
function fmtTgl(s){if(!s)return '';const m=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];const d=new Date(s);return isNaN(d)?s:d.getDate()+' '+m[d.getMonth()]+' '+d.getFullYear()}
function ttd(label,nama,tgl,w){const c=[Math.round(w*.2),Math.round(w*.3),Math.round(w*.25),w-Math.round(w*.2)-Math.round(w*.3)-Math.round(w*.25)];
  return table([[cell(label,c[0],{b:true}),cell(nama||'',c[1]),cell('Tanda Tangan',c[2]),cell('Tanggal\n'+fmtTgl(tgl),c[3])]],c)}

// ---- Lampiran 6: satu halaman per ketidaksesuaian ----
function ncPage(d,nc,last){
  const w=[2400,4500,2460];
  const out=[head('FORMULIR KETIDAKSESUAIAN DAN TINDAK LANJUT KETIDAKSESUAIAN AUDIT PENERAPAN SISTEM MANAJEMEN KESELAMATAN PERTAMBANGAN'),
   kv('Nama Perusahaan',d.company),kv('Tanggal Audit',d.tglAudit),kv('Nomor Ketidaksesuaian',nc.no),para('Uraian Ketidaksesuaian',{b:true,before:80}),
   table([
    [cell([para('Uraian Masalah',{b:true}),para('(Problem)')],w[0]),cell(nc.masalah||'',w[1]),cell([para('Kategori',{b:true}),...checkRow(nc.kat)],w[2])],
    [cell([para('Area / Lokasi / Departemen',{b:true}),para('(Location)')],w[0]),cell(nc.lokasi||'',w[1]+w[2],{span:2})],
    [cell([para('Bukti',{b:true}),para('(Objective Evidence)')],w[0]),cell(bullets(nc.bukti),w[1]+w[2],{span:2})],
    [cell([para('Referensi (Reference)',{b:true}),para('Elemen'),para('Sub Elemen')],w[0]),cell(bullets(nc.ref),w[1]+w[2],{span:2})],
    [cell([para('Deskripsi',{b:true}),para(nc.uraian||'',{al:AlignmentType.JUSTIFY})],CW,{span:3})]
   ],w),
   para('Batas Waktu Perbaikan : '+fmtTgl(nc.batas),{b:true,before:100}),
   ttd('Nama Auditor',nc.auditor,nc.tglAudit||d.tglAudit,CW),ttd('Nama Auditi',nc.auditi,'',CW),
   para('Tindak Lanjut Audit',{b:true,before:160,keepNext:true}),
   table([[cell([para('Akar Permasalahan',{b:true}),para(nc.akar||'')],CW)],
    [cell([para('Tindakan Perbaikan',{b:true}),para('Korektif:',{b:true}),para(nc.korektif||''),para('Preventif:',{b:true}),para(nc.preventif||'')],CW)],
    [cell([para('Bukti',{b:true}),para(nc.buktiTL||'')],CW)],
    [cell([para('Tanggal Perbaikan : '+fmtTgl(nc.tglPerbaikan),{b:true})],CW)]],[CW]),
   ttd('Nama Auditi',nc.auditi,'',CW),
   para('Ulasan Tim Auditor',{b:true,before:160,keepNext:true}),para('Nomor Tindak Lanjut : '+(nc.noTL||nc.no)),
   table([[cell([para('Verifikasi',{b:true}),para(nc.verifikasi||'')],CW)]],[CW]),
   ttd('Nama Auditor',nc.auditor,'',CW)];
  if(!last) out.push(new Paragraph({children:[new PageBreak()]}));
  return out}
function lampiran6(d){return new Document({styles:{default:{document:{run:{font:FONT,size:SZ}}}},sections:[{properties:{page:{size:LETTER,margin:MARG}},children:d.ncs.flatMap((n,i)=>ncPage(d,n,i===d.ncs.length-1))}]})}

// ---- Lampiran 5: rekapitulasi ----
function lampiran5(d){const w=[600,1900,4400,1260,1200],hd=['No','NOMOR KETIDAKSESUAIAN','DESKRIPSI KETIDAKSESUAIAN','KATEGORI','BATAS WAKTU PERBAIKAN'];
  const rows=[hd.map((h,i)=>cell(h,w[i],{b:true,al:AlignmentType.CENTER,fill:'D9D9D9'}))].concat(d.ncs.map((n,i)=>[cell(i+1,w[0],{al:AlignmentType.CENTER}),cell(n.no,w[1]),cell(n.uraian||'',w[2]),cell(n.kat,w[3],{al:AlignmentType.CENTER}),cell(fmtTgl(n.batas),w[4])]));
  return new Document({styles:{default:{document:{run:{font:FONT,size:22}}}},sections:[{properties:{page:{size:LETTER,margin:MARG}},children:[head('FORMULIR REKAPITULASI KETIDAKSESUAIAN'),kv('Nama Perusahaan',d.company),kv('Tanggal Audit',d.tglAudit),para(''),table(rows,w)]}]})}


// ---- Lampiran 2: RTL (landscape) ----
const LAND={width:12240,height:15840,orientation:D.PageOrientation.LANDSCAPE};
function lampiran2(d){const CWL=15840-2880,w=[450,1350,3000,2600,3100,1250,1210];const hd=['No','Nomor Ketidaksesuaian','Deskripsi Ketidaksesuaian','Akar Permasalahan','Tindakan Koreksi','Penanggung Jawab','Batas Waktu Perbaikan'];
  const rows=[hd.map((h,i)=>cell(h,w[i],{b:true,al:AlignmentType.CENTER,fill:'D9D9D9'}))].concat(d.ncs.map((n,i)=>[cell(i+1,w[0],{al:AlignmentType.CENTER}),cell(n.no,w[1]),cell(n.uraian||'',w[2]),cell(n.akar||'',w[3]),cell([para('Korektif:',{b:true}),para(n.korektif||''),para('Preventif:',{b:true}),para(n.preventif||'')],w[4]),cell(n.pj||'',w[5]),cell(fmtTgl(n.batas),w[6])]));
  return new Document({styles:{default:{document:{run:{font:FONT,size:20}}}},sections:[{properties:{page:{size:LAND,margin:MARG}},children:[head('FORMULIR RENCANA TINDAK LANJUT KETIDAKSESUAIAN AUDIT\nPENERAPAN SISTEM MANAJEMEN KESELAMATAN PERTAMBANGAN'),kv('Nama Perusahaan',d.company),kv('Tanggal Audit',d.tglAudit),para(''),table(rows,w)]}]})}
// ---- Lampiran 3: Kesesuaian ----
function lampiran3(d){const w=[2200,7160],rows=[];
  d.kes.forEach(k=>{rows.push([cell('Elemen',w[0],{b:true,fill:'F2F2F2'}),cell(k.elemen,w[1],{fill:'F2F2F2'})],[cell('Sub Elemen',w[0],{b:true}),cell(k.sub,w[1])],[cell('Keterangan',w[0],{b:true}),cell(k.ket||'',w[1],{})])});
  return new Document({styles:{default:{document:{run:{font:FONT,size:SZ}}}},sections:[{properties:{page:{size:LETTER,margin:MARG}},children:[head('FORMULIR KESESUAIAN AUDIT\nPENERAPAN SISTEM MANAJEMEN KESELAMATAN PERTAMBANGAN'),kv('Nama Perusahaan',d.company),kv('Tanggal Audit',d.tglAudit),para('Uraian Kesesuaian',{b:true,before:80}),table(rows,w)]}]})}
// ---- Lampiran 4: Rekomendasi & OFI ----
function lampiran4(d){const w=[3000,6360],rows=[];
  d.ofi.forEach(k=>{rows.push([cell('Elemen',w[0],{b:true,fill:'F2F2F2'}),cell(k.elemen,w[1],{fill:'F2F2F2'})],[cell('Sub Elemen',w[0],{b:true}),cell(k.sub,w[1])],[cell('Rekomendasi dan Peluang untuk Perbaikan',w[0],{b:true}),cell(k.ket||'',w[1])])});
  return new Document({styles:{default:{document:{run:{font:FONT,size:SZ}}}},sections:[{properties:{page:{size:LETTER,margin:MARG}},children:[head('FORMULIR REKOMENDASI DAN PELUANG UNTUK PERBAIKAN\nPENERAPAN SISTEM MANAJEMEN KESELAMATAN PERTAMBANGAN *)'),kv('Nama Perusahaan Auditi',d.company),kv('Tanggal Audit',d.tglAudit),para('Uraian Rekomendasi dan Peluang untuk Perbaikan',{b:true,before:80}),table(rows,w)]}]})}
// ---- Lampiran 1: Excel minimal via JSZip ----
function xe(v){return String(v==null?'':v).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}
function colL(n){let s='';n++;while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26)}return s}
async function lampiran1(d,JS){ // d.rows: [{kind:'el'|'sub',kode,nama,max,nilai,kat,ket,no,w}] d.weights
  const Z=JS||(typeof JSZip!=='undefined'?JSZip:require('jszip')),z=new Z();
  const hdr=['Kriteria','Kode','Sub Elemen','Nilai Maks','Nilai Audit','Kategori','Keterangan','No Temuan'];
  let rows=[[{s:d.company}],[{s:'FORMULIR KRITERIA AUDIT SMKP MINERBA'}],[{s:d.tglAudit}],[],hdr.map(h=>({s:h,b:1}))];
  const start=rows.length+1;let r=start,elRows={};
  for(const x of d.rows){
    if(x.kind==='el'){elRows[x.kode]={row:r,sub:[]};rows.push([{s:x.kode,b:1},{s:x.nama,b:1},{},{},{},{},{},{}]);}
    else{const e=elRows[x.el];e.sub.push(r);rows.push([{},{s:x.kode},{s:x.nama},{n:x.max},x.nilai===''||x.nilai==null?{}:(x.nilai==='NA'?{s:'N/A'}:{n:x.nilai}),{s:x.kat||''},{s:x.ket||''},{s:x.no||''}])}
    r++}
  // ringkasan
  rows.push([]);r++;rows.push([{s:'PERHITUNGAN NILAI AUDIT',b:1}]);r++;rows.push([{s:'Elemen',b:1},{s:'Bobot %',b:1},{s:'Nilai',b:1},{s:'Maks',b:1},{s:'Capaian elemen',b:1},{s:'Nilai audit %',b:1}]);r++;
  const sumRows=[];
  Object.keys(elRows).forEach(k=>{const e=elRows[k],a=e.sub[0],b=e.sub[e.sub.length-1];
    rows.push([{s:k},{n:d.weights[k]},{f:`SUM(E${a}:E${b})`},{f:`SUMIF(E${a}:E${b},"<>N/A",D${a}:D${b})`},{f:`IF(D${r}=0,0,C${r}/D${r})`},{f:`E${r}*B${r}`}]);sumRows.push(r);r++});
  rows.push([{s:'Total',b:1},{},{},{},{},{f:'SUM(F'+sumRows[0]+':F'+sumRows[sumRows.length-1]+')'}]);
  // maks: hitung hanya item yang bukan N/A -> gunakan nilai statis hasil hitung agar aman
  const sheetRows=rows.map((row,i)=>`<row r="${i+1}">${row.map((c,j)=>{const ref=colL(j)+(i+1);
    if(c.f)return `<c r="${ref}"><f>${xe(c.f)}</f></c>`;if(c.n!=null)return `<c r="${ref}"><v>${c.n}</v></c>`;if(c.s!=null)return `<c r="${ref}" t="inlineStr" ${c.b?'s="1"':''}><is><t xml:space="preserve">${xe(c.s)}</t></is></c>`;return ''}).join('')}</row>`).join('');
  z.file('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  z.file('_rels/.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  z.file('xl/workbook.xml','<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Form Penilaian" sheetId="1" r:id="rId1"/></sheets><calcPr fullCalcOnLoad="1"/></workbook>');
  z.file('xl/_rels/workbook.xml.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  z.file('xl/styles.xml','<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf fontId="0"/><xf fontId="1" applyFont="1"/></cellXfs></styleSheet>');
  z.file('xl/worksheets/sheet1.xml','<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="10"/><col min="2" max="2" width="10"/><col min="3" max="3" width="45"/><col min="4" max="5" width="12"/><col min="6" max="6" width="12"/><col min="7" max="7" width="60"/><col min="8" max="8" width="14"/></cols><sheetData>'+sheetRows+'</sheetData></worksheet>');
  return z.generateAsync({type:typeof Blob!=='undefined'?'blob':'nodebuffer'})}


// ---- Laporan Audit Internal ----
function H1(t){return new Paragraph({heading:D.HeadingLevel.HEADING_1,spacing:{before:240,after:120},keepNext:true,children:[run(t,{bold:true})]})}
function H2(t){return new Paragraph({heading:D.HeadingLevel.HEADING_2,spacing:{before:160,after:80},keepNext:true,children:[run(t,{bold:true})]})}
function J(t,o){return para(t,Object.assign({al:AlignmentType.JUSTIFY,after:100},o||{}))}
function bl(t){return new Paragraph({bullet:{level:0},spacing:{after:50},alignment:AlignmentType.JUSTIFY,children:[run(t)]})}
function imgRun(im){try{const b64=im.dataUrl.split(',')[1],bin=typeof atob!=='undefined'?Uint8Array.from(atob(b64),c=>c.charCodeAt(0)):Buffer.from(b64,'base64');
  const type=/png/i.test(im.dataUrl.slice(0,20))?'png':'jpg';return new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:80},children:[new D.ImageRun({type,data:bin,transformation:{width:im.w||420,height:im.h||280}})]})}catch(e){return para('[gambar gagal dimuat]')}}
function ncTable(list,startNo){const w=[500,1500,5860,1500];
  const rows=[['No','NOMOR KETIDAKSESUAIAN','DESKRIPSI KETIDAKSESUAIAN','KATEGORI'].map((h,i)=>cell(h,w[i],{b:true,al:AlignmentType.CENTER,fill:'D9D9D9'}))].concat(list.map((n,i)=>[cell(startNo+i,w[0],{al:AlignmentType.CENTER}),cell(n.no,w[1]),cell(n.uraian||'',w[2],{}),cell(n.kat,w[3],{al:AlignmentType.CENTER})]));
  return table(rows,w)}
function laporan(d){
  const kids=[];const L=d.lap||{};
  // sampul
  kids.push(para('',{after:1200}),para('LAPORAN INTERNAL AUDIT',{b:true,al:AlignmentType.CENTER,size:32,after:60}),para('PENERAPAN SISTEM MANAJEMEN KESELAMATAN PERTAMBANGAN (SMKP)',{b:true,al:AlignmentType.CENTER,size:28,after:240}),
    para(d.company||'',{b:true,al:AlignmentType.CENTER,size:28,after:600}),para('ALAMAT KANTOR',{b:true,al:AlignmentType.CENTER}),
    para('HEAD OFFICE : '+(L.kantorPusat||''),{al:AlignmentType.CENTER}),para('SITE OFFICE : '+(L.alamat||''),{al:AlignmentType.CENTER,after:300}),
    para('TANGGAL PELAKSANAAN AUDIT: '+(d.tglAudit||''),{b:true,al:AlignmentType.CENTER}),new Paragraph({children:[new PageBreak()]}),
    para('DAFTAR ISI',{b:true,al:AlignmentType.CENTER,after:160}),new D.TableOfContents('Daftar Isi',{hyperlink:true,headingStyleRange:'1-2'}),new Paragraph({children:[new PageBreak()]}));
  // 1
  kids.push(H1('1. LATAR BELAKANG'),J(L.latar1||'Audit internal SMKP ini merupakan bagian dari penerapan Elemen-elemen SMKP Minerba, sesuai dengan ketentuan Peraturan Menteri Energi dan Sumber Daya Mineral Nomor 26 Tahun 2018 dan Keputusan Direktur Jenderal Mineral dan Batubara Nomor 185.K/37.04/DJB/2019 Lampiran II.'),
    J(L.latar2||('Audit internal SMKP ini bertujuan untuk mendapatkan gambaran tingkat penerapan SMKP di '+d.company+' dalam periode '+(L.periodeLaporan||'')+'.')));
  // 2
  kids.push(H1('2. GAMBARAN UMUM PERUSAHAAN'),H2('2.1 Data Profil Perusahaan'));
  const pf=[['Nama Perusahaan',d.company],['Alamat Kantor Pusat',L.kantorPusat],['Alamat Domisili Perusahaan',L.alamat],['Jenis Perizinan',L.izin],['Nomor Izin',L.nomorIzin],['Jenis Komoditas Tambang',L.komoditas],['Direktur',L.direktur],['KTT',L.ktt],['Luas Area Konsesi',L.luas],['Jumlah Pekerja',L.pekerja]];
  kids.push(table(pf.map(r=>[cell(r[0],3000,{b:true}),cell(r[1]||'',6360)]),[3000,6360]));
  if(L.profilNarasi)kids.push(J(L.profilNarasi,{before:80}));
  kids.push(H2('2.2 Kegiatan Operasional Tambang'));
  (L.operasional||[]).forEach(o=>{kids.push(para(o.judul,{b:true,before:60,keepNext:true}),J(o.teks||''))});
  if((L.produksi||[]).length){kids.push(para('Target produksi tahun berjalan:',{before:80}),table([['No','Target Produksi','Pencapaian'].map((h,i)=>cell(h,[600,4380,4380][i],{b:true,al:AlignmentType.CENTER,fill:'D9D9D9'}))].concat(L.produksi.map((r,i)=>[cell(i+1,600),cell(r.u+' : '+r.t,4380),cell(r.p,4380)])),[600,4380,4380]))}
  if((L.peralatan||[]).length){kids.push(para('Sarana, prasarana, instalasi, dan peralatan:',{before:100}),table([['No','Jenis','Informasi / PT','Kontraktor','Jumlah'].map((h,i)=>cell(h,[600,3600,1900,1900,1360][i],{b:true,al:AlignmentType.CENTER,fill:'D9D9D9'}))].concat(L.peralatan.map((r,i)=>[cell(i+1,600),cell(r.jenis,3600),cell(r.pt,1900),cell(r.kontraktor,1900),cell((Number(r.pt)||0)+(Number(r.kontraktor)||0),1360)])),[600,3600,1900,1900,1360]))}
  kids.push(H2('2.3 Struktur Organisasi '+d.company));(d.images||[]).filter(i=>i.grup==='struktur').forEach(i=>kids.push(imgRun(i)));
  kids.push(H2('2.4 Penerapan SMKP Minerba'));
  const EL=d.elNames;Object.keys(EL).forEach(e=>{kids.push(para('Elemen '+e+' '+EL[e],{b:true,before:60,keepNext:true}));(L.narasi&&L.narasi[e]?String(L.narasi[e]).split('\n').filter(Boolean):['(belum diisi)']).forEach(t=>kids.push(bl(t)))});
  // 3
  kids.push(H1('3. LINGKUP AUDIT'),J('Audit internal SMKP ini merupakan audit berkala yang dilaksanakan secara internal, yang bertujuan untuk mendapatkan gambaran tingkat penerapan SMKP Minerba yang dilaksanakan oleh '+d.company+'.'),
    J('Audit SMKP Minerba ini dilaksanakan dengan mengambil sampel berdasarkan judgement auditor dan risiko tertinggi dari setiap area kegiatan, yang terdiri atas:'));
  const grp={};(d.areas||[]).forEach(a=>{(grp[a.grup||'Area Lainnya']=grp[a.grup||'Area Lainnya']||[]).push(a)});
  Object.keys(grp).forEach((g,i)=>{kids.push(para((i+1)+'. '+g,{b:true,before:60}));grp[g].forEach(a=>kids.push(bl(a.nama+(a.pelaksana&&a.pelaksana!=='IUP'?' '+a.pelaksana:''))))});
  if((d.pengecualian||[]).length){kids.push(para('Pengecualian (N/A):',{b:true,before:100}));d.pengecualian.forEach(x=>kids.push(bl(x)))}
  // 4
  kids.push(H1('4. PELAKSANAAN AUDIT & TIM AUDITOR'),J('Audit Internal SMKP dilakukan oleh Internal Auditor dari '+d.company+' yang dilaksanakan sejak tanggal '+(d.tglAudit||'')+' dengan tim auditor yang terdiri atas:'));
  (d.tim||[]).forEach(t=>kids.push(para((t.jabatan||'Auditor')+' : '+t.nama,{after:30})));
  // 5
  kids.push(H1('5. RINGKASAN LAPORAN & PENILAIAN AUDIT'),H2('5.1 Hasil Audit'),J('Hasil audit SMKP Minerba yang dilaksanakan terhadap '+d.company+' adalah audit internal sejak '+(d.tglAudit||'')+', dengan memperoleh informasi sebagai berikut.'),
    para('Informasi Praktik Terbaik Penerapan SMKP Minerba (Temuan Positif)',{b:true,before:80,keepNext:true}));
  const kr=[];let lastEl='';(d.kes||[]).forEach(k=>{if(k.elemen!==lastEl){kr.push([cell('Elemen',2000,{b:true,fill:'F2F2F2'}),cell(k.elemen,7360,{fill:'F2F2F2'})]);lastEl=k.elemen}kr.push([cell('Sub Elemen',2000,{b:true}),cell(k.sub,7360)],[cell('Keterangan',2000,{b:true}),cell(k.ket||'',7360)])});
  kids.push(kr.length?table(kr,[2000,7360]):para('(belum ada temuan positif)'));
  kids.push(para('Ringkasan Temuan Audit Ketidaksesuaian',{b:true,before:160,keepNext:true}),J('Selama pelaksanaan audit internal SMKP di '+d.company+', tim auditor telah menemukan ketidaksesuaian terhadap persyaratan elemen SMKP sebagai berikut:'));
  let no=1;[['KRITIKAL','Kritikal'],['MAYOR','Mayor'],['MINOR','Minor']].forEach(([k,lbl])=>{const l=(d.ncs||[]).filter(n=>n.kat===k);kids.push(para('Kategori '+lbl,{b:true,before:100,keepNext:true}),para('Jumlah : Terdapat '+l.length+' temuan ketidaksesuaian dengan kategori '+lbl),para('Uraian :'));if(l.length){kids.push(ncTable(l,no));no+=l.length}});
  kids.push(H2('5.2 Tingkat Pencapaian Penerapan SMKP Minerba'));
  const w=[600,4260,1400,1550,1550];kids.push(table([['No','ELEMEN','BOBOT','NILAI AUDIT','CAPAIAN'].map((h,i)=>cell(h,w[i],{b:true,al:AlignmentType.CENTER,fill:'D9D9D9'}))].concat((d.elSum||[]).map((r,i)=>[cell(i+1,w[0],{al:AlignmentType.CENTER}),cell(r.nama,w[1]),cell(r.bobot+'%',w[2],{al:AlignmentType.CENTER}),cell(r.kontribusi.toFixed(2)+'%',w[3],{al:AlignmentType.CENTER}),cell(r.capaian.toFixed(2)+'%',w[4],{al:AlignmentType.CENTER})])).concat([[cell('TOTAL',w[0]+w[1],{b:true,span:2}),cell('100%',w[2],{b:true,al:AlignmentType.CENTER}),cell(d.total.toFixed(2)+'%',w[3],{b:true,al:AlignmentType.CENTER}),cell(d.total.toFixed(2)+'%',w[4],{b:true,al:AlignmentType.CENTER})]]),w));
  // 6
  kids.push(H1('6. LAMPIRAN-LAMPIRAN'));(d.lampiran||[]).forEach((x,i)=>kids.push(para((i+1)+'. '+x,{after:30})));
  // 7
  kids.push(H1('7. HALAMAN PENUTUP'),para((L.kota||'')+', '+fmtTgl(L.tglLaporan),{before:100}),para('Dibuat oleh:',{before:80}));
  (d.tim||[]).forEach((t,i)=>kids.push(para((i+1)+'. '+(t.jabatan||'Auditor')+' : '+t.nama+(t.reg?' (No. Reg : '+t.reg+')':''),{after:40})));
  kids.push(para('Diketahui oleh,',{before:160}),para('',{after:600}),para(L.ktt||'',{b:true,after:0}),para('Kepala Teknik Tambang'));
  const dok=(d.images||[]).filter(i=>i.grup==='dok');if(dok.length){kids.push(new Paragraph({children:[new PageBreak()]}));dok.forEach(i=>{kids.push(para(i.judul||'',{b:true,al:AlignmentType.CENTER}),imgRun(i))})}
  kids.push(new Paragraph({children:[new PageBreak()]}),para('RESPON MANAJEMEN TERHADAP PELAKSANAAN AUDIT',{b:true,al:AlignmentType.CENTER,after:160}),
    J('Sehubungan dengan selesainya kegiatan audit internal Sistem Manajemen Keselamatan Pertambangan Mineral dan Batubara yang dilaksanakan terhadap '+d.company+' pada tanggal '+(d.tglAudit||'')+':'),
    bl('Kami menerima hasil audit internal Sistem Manajemen Keselamatan Pertambangan Mineral dan Batubara di '+d.company+' dengan nilai '+d.total.toFixed(2)+'%.'),
    bl(L.respon1||('Kami berkomitmen meningkatkan penerapan SMKP Minerba sesuai ketentuan Peraturan Menteri ESDM Nomor 26 Tahun 2018.')),
    ...(L.respon2?[bl(L.respon2)]:[]),J(L.respon3||'Kami akan mendukung sepenuhnya proses tindak lanjut dan perbaikan berdasarkan hasil audit internal SMKP ini.',{before:80}),
    para((L.kota||'')+', '+fmtTgl(L.tglRespon||L.tglLaporan),{before:160}),para('',{after:600}),para(L.ktt||'',{b:true,after:0}),para('Kepala Teknik Tambang'));
  const heads={heading1:{run:{font:FONT,size:SZ,bold:true,color:'000000'}},heading2:{run:{font:FONT,size:SZ,bold:true,color:'000000'}}};
  return new Document({features:{updateFields:true},styles:{default:{document:{run:{font:FONT,size:SZ}},heading1:heads.heading1,heading2:heads.heading2}},sections:[{properties:{page:{size:LETTER,margin:MARG}},children:kids}]})}

async function toBlob(doc){return Packer.toBlob(doc)}
async function toBuffer(doc){return Packer.toBuffer(doc)}
const api={laporan,lampiran1,lampiran2,lampiran3,lampiran4,lampiran5,lampiran6,toBlob,toBuffer,fmtTgl};
if(typeof module!=='undefined') module.exports=api; else root.DocGen=api;
})(this);
