/**
 * Control de Proyectos - Google Sheets
 * Sheet dedicado: 1dueVobI_41RkUk1pVNUhh0ehN1CkakAfrmXkIoztaA4
 * Importante: Apps Script debe desplegarse de nuevo después de pegar el código.
 * Acceso "cualquier persona": no equivale a autenticación, quien tenga la URL puede editar.
 */
const CONFIG={ID_PLANILLA:'1dueVobI_41RkUk1pVNUhh0ehN1CkakAfrmXkIoztaA4',HOJA:'Datos',REGISTRO:'Registro'};
function obtenerLibro(){return SpreadsheetApp.openById(CONFIG.ID_PLANILLA);}
function obtenerHoja(){
 const libro=obtenerLibro();let hoja=libro.getSheetByName(CONFIG.HOJA);
 if(!hoja){hoja=libro.insertSheet(CONFIG.HOJA);hoja.appendRow(['Versión','Datos JSON','Última modificación']);}
 return hoja;
}
function obtenerRegistro(){
 const libro=obtenerLibro();let hoja=libro.getSheetByName(CONFIG.REGISTRO);
 if(!hoja){hoja=libro.insertSheet(CONFIG.REGISTRO);hoja.appendRow(['Fecha','Usuario','Versión','Acción']);}
 return hoja;
}
function obtenerDatos(){
 const hoja=obtenerHoja(),n=hoja.getLastRow();
 if(n<2)return {ok:true,version:0,proyectos:[],actualizado:null};
 const row=hoja.getRange(n,1,1,3).getValues()[0];
 const proyectos=JSON.parse(row[1]||'[]');
 if(!Array.isArray(proyectos))throw Error('Datos incorrectos');
 return {ok:true,version:Number(row[0])||0,proyectos,actualizado:row[2]?new Date(row[2]).toISOString():null};
}
function guardarDatos(proyectos,versionEsperada){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  if(!Array.isArray(proyectos))throw Error('Se esperaba una lista de proyectos');
  const actual=obtenerDatos();
  if(Number(versionEsperada)!==actual.version)return {ok:false,conflicto:true,version:actual.version,mensaje:'Datos modificados desde otra sesión; recargá antes de guardar'};
  const next=actual.version+1;
  obtenerHoja().appendRow([next,JSON.stringify(proyectos),new Date()]);
  obtenerRegistro().appendRow([new Date(),Session.getActiveUser().getEmail()||'No identificado',next,'Actualización']);
  return {ok:true,version:next};
 }finally{lock.releaseLock();}
}
function obtenerResumen(){
 const d=obtenerDatos(),a=d.proyectos.filter(p=>p.active===true);
 let presupuesto=0,pagado=0;
 a.forEach(p=>(p.items||[]).forEach(i=>{presupuesto+=Number(i.cost)||0;pagado+=Number(i.paidAmount)||0;}));
 return {ok:true,version:d.version,proyectosActivos:a.length,proyectosTotales:d.proyectos.length,presupuesto,pagado,saldoPendiente:presupuesto-pagado};
}
function probarConexion(){
 const d=obtenerDatos();
 Logger.log('CONEXIÓN CORRECTA: '+obtenerLibro().getName()+'; proyectos='+d.proyectos.length+'; versión='+d.version);
 return d;
}
/**
 * JSONP: lectura desde GitHub Pages, sin CORS.
 * callback solamente acepta identificadores JavaScript sin caracteres especiales.
 */
function doGet(e){
 try{
  const json=JSON.stringify(obtenerDatos());
  const callback=(e&&e.parameter&&e.parameter.callback)||'';
  if(callback && /^[A-Za-z_$][\w$]{0,79}$/.test(callback))
   return ContentService.createTextOutput(callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
 }catch(err){
  return ContentService.createTextOutput(JSON.stringify({ok:false,mensaje:String(err)})).setMimeType(ContentService.MimeType.JSON);
 }
}
/**
 * Escritura mediante formulario POST de un solo campo "payload" con JSON.
 * Para el cliente GitHub Pages utilizar un iframe oculto; comprobar confirmación
 * recargando la versión por JSONP. No tratar un envío como guardado hasta verificarlo.
 */
function doPost(e){
 let response;
 try{
  const body=JSON.parse((e.parameter&&e.parameter.payload)||(e.postData&&e.postData.contents)||'{}');
  response=guardarDatos(body.proyectos,body.version);
 }catch(err){response={ok:false,mensaje:String(err)};}
 return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}
