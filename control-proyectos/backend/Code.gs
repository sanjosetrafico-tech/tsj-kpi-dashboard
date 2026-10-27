/**
 * Control de Proyectos - Google Apps Script
 * Vincular este proyecto al Google Sheet dedicado mediante Extensiones > Apps Script.
 * Implementar como aplicación web: Ejecutar como propietario.
 * IMPORTANTE: "Cualquier persona" permite modificaciones públicas; limitar acceso a usuarios autorizados.
 * La página está en otro dominio: para escrituras desde navegador se requiere resolver CORS
 * o utilizar una interfaz Apps Script HTMLService servida desde el propio script.
 */
const TAB='Datos';
function getSheet_(){
 const book=SpreadsheetApp.getActiveSpreadsheet();
 let sh=book.getSheetByName(TAB);
 if(!sh){sh=book.insertSheet(TAB);sh.appendRow(['revision','json','fecha']);}
 return sh;
}
function getState(){
 const sh=getSheet_();
 const row=sh.getLastRow();
 if(row<2)return {revision:0,projects:null};
 return {revision:Number(sh.getRange(row,1).getValue())||0,projects:JSON.parse(sh.getRange(row,2).getValue())};
}
function saveState(json,expectedRevision){
 const lock=LockService.getDocumentLock();
 lock.waitLock(30000);
 try{
  const prior=getState();
  if(Number(expectedRevision)!==prior.revision) return {ok:false,conflict:true,revision:prior.revision};
  const projects=JSON.parse(json);
  if(!Array.isArray(projects))throw new Error('Formato incorrecto');
  const next=prior.revision+1;
  getSheet_().appendRow([next,JSON.stringify(projects),new Date()]);
  return {ok:true,revision:next};
 }finally{lock.releaseLock();}
}
function doGet(){return ContentService.createTextOutput(JSON.stringify(getState())).setMimeType(ContentService.MimeType.JSON);}
