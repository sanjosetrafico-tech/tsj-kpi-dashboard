
const CH_STORE_KEY='TSJ_C_HORARIO_V1';
let CH_VERSION=0;
let CH_DRAFT_TIME=new Date().toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit',hour12:false});
let CH_PENDING={horario:null,cabecera:null};
let CH_STATE=loadCHState();

function loadCHState(){
  try{
    const raw=localStorage.getItem(CH_STORE_KEY);
    if(!raw)return{activeDate:'',days:{}};
    const parsed=JSON.parse(raw);
    return parsed&&parsed.days?parsed:{activeDate:'',days:{}};
  }catch(err){
    return{activeDate:'',days:{}};
  }
}

function saveCHState(){
  try{localStorage.setItem(CH_STORE_KEY,JSON.stringify(CH_STATE));}catch(err){}
  CH_VERSION++;
  if(typeof LAST_RENDER!=='undefined')LAST_RENDER.chorario='';
}

function chNumber(value){
  if(value===null||value===undefined||value==='')return null;
  if(typeof value==='number')return Number.isFinite(value)?value:null;
  let s=String(value).trim().replace('%','').replace(/\s/g,'');
  if(!s)return null;
  if(s.includes(',')&&!s.includes('.'))s=s.replace(',','.');
  const n=Number(s);
  return Number.isFinite(n)?n:null;
}

function chDateISO(value){
  const raw=String(value||'').trim();
  let m=raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m)return m[1]+'-'+m[2]+'-'+m[3];
  m=raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(m)return m[3]+'-'+String(m[2]).padStart(2,'0')+'-'+String(m[1]).padStart(2,'0');
  return '';
}

function chDateLabel(iso){
  const m=String(iso||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m?m[3]+'/'+m[2]+'/'+m[1]:String(iso||'');
}

function chHeaderKey(value){
  return normalizeText(String(value||'')).replace(/\s+/g,' ').trim();
}

function chFindKey(row,candidates){
  const keys=Object.keys(row||{});
  for(const candidate of candidates){
    const target=chHeaderKey(candidate);
    const found=keys.find(function(key){return chHeaderKey(key)===target;});
    if(found!==undefined)return found;
  }
  return '';
}

function chRowsFromSheet(sheet){
  const rows=XLSX.utils.sheet_to_json(sheet,{defval:'',raw:true});
  return rows.filter(function(row){
    const key=chFindKey(row,['Linea','Línea']);
    return key&&String(row[key]||'').trim();
  });
}

function chParseRows(rows,type){
  if(!rows.length)throw new Error('El archivo no contiene filas de datos.');

  const lineKey=chFindKey(rows[0],['Linea','Línea']);
  const dateKey=chFindKey(rows[0],['Fecha']);
  const execKey=chFindKey(rows[0],[' Viajes Ejecutados','Viajes Ejecutados']);
  const atrasoKey=chFindKey(rows[0],['Porcentaje Atrasos']);
  const adelantoKey=chFindKey(rows[0],['Porcentaje Adelantos']);
  const complianceKey=type==='horario'
    ?chFindKey(rows[0],['C. Horarios','C Horarios'])
    :chFindKey(rows[0],['Cumplimiento']);

  if(!lineKey||!complianceKey||!atrasoKey||!adelantoKey){
    throw new Error(type==='horario'
      ?'No se reconocen las columnas de Cumplimiento Horario.'
      :'No se reconocen las columnas de Cumplimiento Cabecera.');
  }

  const parsed=rows.map(function(row){
    return{
      linea:String(row[lineKey]||'').trim().toUpperCase(),
      cumplimiento:chNumber(row[complianceKey]),
      atraso:chNumber(row[atrasoKey]),
      adelanto:chNumber(row[adelantoKey]),
      viajes:chNumber(execKey?row[execKey]:null)||0,
      fecha:chDateISO(dateKey?row[dateKey]:'')
    };
  }).filter(function(row){
    return row.linea&&row.cumplimiento!==null;
  });

  parsed.sort(function(a,b){
    return a.cumplimiento-b.cumplimiento||a.linea.localeCompare(b.linea,'es',{numeric:true});
  });

  return{
    rows:parsed,
    date:(parsed.find(function(row){return row.fecha;})||{}).fecha||''
  };
}

async function chReadExcel(file,type){
  if(!window.XLSX)throw new Error('No se pudo cargar el lector de Excel.');
  const buffer=await file.arrayBuffer();
  const workbook=XLSX.read(buffer,{type:'array'});
  const sheetName=workbook.SheetNames[0];
  if(!sheetName)throw new Error('El Excel no contiene hojas.');
  const rows=chRowsFromSheet(workbook.Sheets[sheetName]);
  return chParseRows(rows,type);
}

async function handleCHFile(event,type){
  const file=event.target.files&&event.target.files[0];
  if(!file)return;
  try{
    const parsed=await chReadExcel(file,type);
    CH_PENDING[type]=Object.assign({},parsed,{name:file.name});
    renderCHorario();
  }catch(err){
    CH_PENDING[type]=null;
    alert(err.message||'No se pudo leer el Excel.');
    renderCHorario();
  }
}

function chWeighted(rows,key){
  const valid=(rows||[]).filter(function(row){return Number.isFinite(row[key]);});
  if(!valid.length)return null;
  const totalWeight=valid.reduce(function(sum,row){return sum+(Number(row.viajes)||0);},0);
  if(totalWeight>0){
    return valid.reduce(function(sum,row){
      return sum+row[key]*(Number(row.viajes)||0);
    },0)/totalWeight;
  }
  return valid.reduce(function(sum,row){return sum+row[key];},0)/valid.length;
}

function chFmt(value){
  if(!Number.isFinite(value))return '—';
  return value.toLocaleString('es-AR',{minimumFractionDigits:0,maximumFractionDigits:2})+'%';
}

function chClass(value){
  if(!Number.isFinite(value))return '';
  if(value<90)return 'bad';
  if(value<95)return 'warn';
  return 'good';
}

function chRowClass(value){
  const cls=chClass(value);
  return cls==='bad'?'ch-row-bad':cls==='warn'?'ch-row-warn':'';
}

function chTrend(current,previous){
  if(!Number.isFinite(current)||!Number.isFinite(previous)){
    return{arrow:'—',label:'Sin corte anterior',delta:null,cls:'flat'};
  }
  const delta=current-previous;
  if(Math.abs(delta)<0.01){
    return{arrow:'→',label:'Sin cambio',delta:0,cls:'flat'};
  }
  return{
    arrow:delta>0?'↑':'↓',
    label:(delta>0?'+':'')+delta.toLocaleString('es-AR',{minimumFractionDigits:2,maximumFractionDigits:2})+' pts',
    delta:delta,
    cls:delta>0?'up':'down'
  };
}

function chTrendHtml(current,previous,compact){
  const trend=chTrend(current,previous);
  const label=trend.delta===null?'Sin corte anterior':trend.label;
  return '<span class="ch-trend '+trend.cls+(compact?' compact':'')+'" title="'+esc(label)+'">'+
    '<b>'+trend.arrow+'</b>'+
    (compact&&trend.delta!==null?'<small>'+esc(trend.label.replace(' pts',''))+'</small>':'<span>'+esc(label)+'</span>')+
    '</span>';
}

function chPreviousLine(previousRows,linea){
  return (previousRows||[]).find(function(row){return row.linea===linea;})||null;
}

function chMetricCard(label,rows,kind,previousRows){
  const avg=chWeighted(rows,'cumplimiento');
  const prevAvg=previousRows?chWeighted(previousRows,'cumplimiento'):null;
  const atraso=chWeighted(rows,'atraso');
  const adelanto=chWeighted(rows,'adelanto');
  return '<div class="ch-summary-card '+kind+'">'+
    '<div class="ch-summary-main">'+
      '<div>'+
        '<div class="eyebrow">'+esc(label)+'</div>'+
        '<div class="big '+chClass(avg)+'">'+chFmt(avg)+'</div>'+
        '<div class="sub">Atraso '+chFmt(atraso)+' · Adelanto '+chFmt(adelanto)+'</div>'+
      '</div>'+
      chTrendHtml(avg,prevAvg,false)+
    '</div>'+
    '</div>';
}

function chTable(title,rows,kind,previousRows){
  const body=(rows||[]).map(function(row){
    const prev=chPreviousLine(previousRows,row.linea);
    const prevValue=prev?prev.cumplimiento:null;
    return '<tr class="'+chRowClass(row.cumplimiento)+'">'+
      '<td class="ch-line">'+esc(row.linea)+'</td>'+
      '<td><span class="ch-score-wrap"><span class="ch-score '+chClass(row.cumplimiento)+'">'+chFmt(row.cumplimiento)+'</span>'+
      chTrendHtml(row.cumplimiento,prevValue,true)+'</span></td>'+
      '<td>'+chFmt(row.atraso)+'</td>'+
      '<td>'+chFmt(row.adelanto)+'</td>'+
      '</tr>';
  }).join('');

  return '<section class="ch-panel">'+
    '<div class="ch-panel-title '+kind+'"><span>'+esc(title)+'</span><span class="ch-meta">Meta ≥90%</span></div>'+
    '<table class="ch-table"><thead><tr><th>Línea</th><th>Cumplimiento</th><th>Atraso</th><th>Adelanto</th></tr></thead>'+
    '<tbody>'+body+'</tbody></table></section>';
}

function chActiveCuts(){
  const date=CH_STATE.activeDate;
  return date&&CH_STATE.days[date]?CH_STATE.days[date]:[];
}

function chAvailableDates(){
  return Object.keys(CH_STATE.days||{}).sort().reverse();
}

function chSetActiveDate(value){
  CH_STATE.activeDate=value;
  saveCHState();
  renderCHorario();
}

function chSaveCut(){
  if(!CH_PENDING.horario||!CH_PENDING.cabecera){
    alert('Importá los dos Excel: Cumplimiento Horario y Cumplimiento Cabecera.');
    return;
  }

  const dateH=CH_PENDING.horario.date;
  const dateC=CH_PENDING.cabecera.date;

  if(!dateH||!dateC){
    alert('No pude identificar la fecha en uno de los archivos.');
    return;
  }
  if(dateH!==dateC){
    alert('Los dos Excel corresponden a fechas distintas.');
    return;
  }

  const cuts=(CH_STATE.days[dateH]||[]).slice();
  if(cuts.length>=3){
    alert('Ya hay 3 cortes cargados para ese día.');
    return;
  }

  const time=String(CH_DRAFT_TIME||'').trim()||
    new Date().toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit',hour12:false});

  cuts.push({
    id:Date.now(),
    time:time,
    horario:CH_PENDING.horario.rows,
    cabecera:CH_PENDING.cabecera.rows,
    sourceHorario:CH_PENDING.horario.name,
    sourceCabecera:CH_PENDING.cabecera.name
  });

  cuts.sort(function(a,b){return String(a.time).localeCompare(String(b.time));});

  CH_STATE.days[dateH]=cuts;
  CH_STATE.activeDate=dateH;
  CH_PENDING={horario:null,cabecera:null};
  CH_DRAFT_TIME=new Date().toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit',hour12:false});
  saveCHState();
  renderCHorario();
}

function chDeleteCut(id){
  const date=CH_STATE.activeDate;
  if(!date)return;

  CH_STATE.days[date]=(CH_STATE.days[date]||[]).filter(function(cut){
    return String(cut.id)!==String(id);
  });

  if(!CH_STATE.days[date].length){
    delete CH_STATE.days[date];
    CH_STATE.activeDate=chAvailableDates()[0]||'';
  }

  saveCHState();
  renderCHorario();
}

function chCriticalText(rows){
  const values=(rows||[]).filter(function(row){
    return Number.isFinite(row.cumplimiento)&&row.cumplimiento<90;
  }).map(function(row){
    return row.linea+' '+chFmt(row.cumplimiento);
  });

  return values.join(', ')||'Sin líneas <90%';
}

function chLatestCutInfo(){
  const cuts=chActiveCuts();
  if(!cuts.length)return null;
  return{
    cut:cuts[cuts.length-1],
    previous:cuts.length>1?cuts[cuts.length-2]:null,
    index:cuts.length-1,
    total:cuts.length,
    date:CH_STATE.activeDate
  };
}

function chShareText(){
  const info=chLatestCutInfo();
  if(!info)return '';
  const cut=info.cut;
  const prev=info.previous;
  const h=chWeighted(cut.horario,'cumplimiento');
  const c=chWeighted(cut.cabecera,'cumplimiento');
  const ph=prev?chWeighted(prev.horario,'cumplimiento'):null;
  const pc=prev?chWeighted(prev.cabecera,'cumplimiento'):null;
  const th=chTrend(h,ph);
  const tc=chTrend(c,pc);

  return[
    'Cumplimiento · '+chDateLabel(info.date)+' · Corte '+(info.index+1)+' · '+cut.time,
    'C. Horario: '+chFmt(h)+(th.delta===null?'':' '+th.arrow+' '+th.label),
    'C. Cabecera: '+chFmt(c)+(tc.delta===null?'':' '+tc.arrow+' '+tc.label)
  ].join('\n');
}

function chExportNode(){
  const info=chLatestCutInfo();
  if(!info)throw new Error('No hay cortes cargados.');
  const source=document.querySelector('.ch-cut[data-cut-index="'+info.index+'"]');
  if(!source)throw new Error('No se encontro el corte para exportar.');

  const stage=document.createElement('div');
  stage.className='ch-export-stage';
  const clone=source.cloneNode(true);
  clone.classList.add('ch-export-card');

  clone.querySelectorAll('.ch-delete,.ch-note').forEach(function(el){el.remove();});

  const head=clone.querySelector('.ch-cut-head');
  if(head){
    head.innerHTML='<div class="ch-export-meta">Corte '+(info.index+1)+' · '+esc(info.cut.time)+' · '+chDateLabel(info.date)+'</div>';
  }

  stage.appendChild(clone);
  document.body.appendChild(stage);
  return{stage:stage,node:clone,info:info};
}

async function chRenderCanvas(){
  if(!window.html2canvas)throw new Error('No se pudo cargar el generador de imagen.');
  const out=chExportNode();
  try{
    await new Promise(function(resolve){requestAnimationFrame(function(){requestAnimationFrame(resolve);});});
    return await html2canvas(out.node,{
      backgroundColor:'#ffffff',
      scale:2,
      useCORS:true,
      logging:false
    });
  }finally{
    out.stage.remove();
  }
}

function chCanvasBlob(canvas){
  return new Promise(function(resolve,reject){
    canvas.toBlob(function(blob){
      if(blob)resolve(blob);
      else reject(new Error('No se pudo generar la imagen.'));
    },'image/png',0.96);
  });
}

function chDownloadBlob(blob,name){
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;
  a.download=name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function(){URL.revokeObjectURL(url);},1500);
}

async function shareCHWhatsApp(){
  const info=chLatestCutInfo();
  if(!info){
    alert('No hay cortes cargados.');
    return;
  }

  try{
    const canvas=await chRenderCanvas();
    const blob=await chCanvasBlob(canvas);
    const safeTime=String(info.cut.time||'').replace(':','');
    const name='Cumplimiento_'+info.date+'_'+safeTime+'.png';
    const file=new File([blob],name,{type:'image/png'});
    const text=chShareText();

    if(navigator.canShare&&navigator.canShare({files:[file]})&&navigator.share){
      await navigator.share({
        files:[file],
        title:'Cumplimiento horario y cabecera',
        text:text
      });
      return;
    }

    chDownloadBlob(blob,name);
    window.open('https://wa.me/?text='+encodeURIComponent(text+'\n\nSe genero la imagen para adjuntar.'),'_blank','noopener');
  }catch(err){
    if(err&&err.name==='AbortError')return;
    alert(err.message||'No se pudo generar la imagen.');
  }
}

async function copyCHSummary(){
  const text=chShareText();
  if(!text){
    alert('No hay cortes cargados.');
    return;
  }
  try{
    await navigator.clipboard.writeText(text);
    alert('Resumen copiado.');
  }catch(err){
    alert('No se pudo copiar el resumen.');
  }
}

async function exportCHorarioPDF(){
  const info=chLatestCutInfo();
  if(!info){
    alert('No hay cortes cargados.');
    return;
  }
  if(!window.jspdf||!window.jspdf.jsPDF){
    alert('No se pudo cargar el generador de PDF.');
    return;
  }

  try{
    const canvas=await chRenderCanvas();
    const img=canvas.toDataURL('image/png',1);
    const jsPDF=window.jspdf.jsPDF;
    const pdf=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
    const pageW=297;
    const pageH=210;
    const margin=7;
    const maxW=pageW-margin*2;
    const maxH=pageH-margin*2;
    const ratio=Math.min(maxW/canvas.width,maxH/canvas.height);
    const w=canvas.width*ratio;
    const h=canvas.height*ratio;
    const x=(pageW-w)/2;
    const y=(pageH-h)/2;
    pdf.addImage(img,'PNG',x,y,w,h,undefined,'FAST');
    pdf.save('Cumplimiento_'+info.date+'_'+String(info.cut.time||'').replace(':','')+'.pdf');
  }catch(err){
    alert(err.message||'No se pudo generar el PDF.');
  }
}

function renderCHorario(){
  const root=document.querySelector('#chorario');
  if(!root)return;

  const dates=chAvailableDates();
  if(!CH_STATE.activeDate&&dates.length)CH_STATE.activeDate=dates[0];

  const cuts=chActiveCuts();
  const canSave=!!(CH_PENDING.horario&&CH_PENDING.cabecera)&&cuts.length<3;

  const dateSelect=dates.length
    ?'<select class="ch-action" onchange="chSetActiveDate(this.value)">'+
      dates.map(function(date){
        return '<option value="'+date+'" '+(date===CH_STATE.activeDate?'selected':'')+'>'+chDateLabel(date)+'</option>';
      }).join('')+
      '</select>'
    :'';

  const cutsHtml=cuts.length
    ?'<div class="ch-cuts">'+cuts.map(function(cut,index){
      const previous=index>0?cuts[index-1]:null;
      return '<article class="ch-cut" data-cut-index="'+index+'">'+
        '<div class="ch-cut-head">'+
          '<div class="ch-cut-title"><b>Corte '+(index+1)+' · '+esc(cut.time)+'</b><span>'+chDateLabel(CH_STATE.activeDate)+'</span></div>'+
          '<button class="ch-delete" type="button" onclick="chDeleteCut(\''+cut.id+'\')">Eliminar</button>'+
        '</div>'+
        '<div class="ch-summary">'+
          chMetricCard('Cumplimiento Horario',cut.horario,'horario',previous?previous.horario:null)+
          chMetricCard('Cumplimiento Cabecera',cut.cabecera,'cabecera',previous?previous.cabecera:null)+
        '</div>'+
        '<div class="ch-grid">'+
          chTable('C. Horario',cut.horario,'horario',previous?previous.horario:null)+
          chTable('C. Cabecera',cut.cabecera,'cabecera',previous?previous.cabecera:null)+
        '</div>'+
        '<div class="ch-note"><b>Criterio:</b> líneas ordenadas de menor a mayor cumplimiento. Rojo &lt;90% · ámbar 90–94,99% · verde ≥95%.</div>'+
      '</article>';
    }).join('')+'</div>'
    :'<div class="ch-empty">Importá los dos Excel del primer corte para generar el informe del día.</div>';

  root.innerHTML=
    '<div class="section-head">'+
      '<div><h2>C. Horario</h2><div style="font-size:12px;color:var(--muted);margin-top:4px">Cumplimiento Horario y Cumplimiento Cabecera · hasta 3 cortes diarios</div></div>'+
      '<div class="ch-report-actions">'+
        dateSelect+
        '<button class="ch-action" type="button" onclick="copyCHSummary()">Copiar resumen</button>'+
        '<button class="ch-action whatsapp" type="button" onclick="shareCHWhatsApp()">WhatsApp · imagen</button>'+
        '<button class="ch-action primary" type="button" onclick="exportCHorarioPDF()">PDF A4</button>'+
      '</div>'+
    '</div>'+
    '<div class="ch-wrap">'+
      '<div class="ch-toolbar">'+
        '<div class="ch-upload">'+
          '<strong>Excel · C. Horario</strong>'+
          '<label class="ch-file-label"><span class="ch-file-name">'+esc(CH_PENDING.horario?CH_PENDING.horario.name:'Seleccionar archivo')+'</span><span>Importar</span>'+
          '<input type="file" accept=".xlsx,.xls" onchange="handleCHFile(event,\'horario\')"></label>'+
          '<div class="ch-file-state '+(CH_PENDING.horario?'ok':'')+'">'+
            (CH_PENDING.horario?'Listo · '+chDateLabel(CH_PENDING.horario.date):'Línea · C. Horarios · Atrasos · Adelantos')+
          '</div>'+
        '</div>'+
        '<div class="ch-upload">'+
          '<strong>Excel · C. Cabecera</strong>'+
          '<label class="ch-file-label"><span class="ch-file-name">'+esc(CH_PENDING.cabecera?CH_PENDING.cabecera.name:'Seleccionar archivo')+'</span><span>Importar</span>'+
          '<input type="file" accept=".xlsx,.xls" onchange="handleCHFile(event,\'cabecera\')"></label>'+
          '<div class="ch-file-state '+(CH_PENDING.cabecera?'ok':'')+'">'+
            (CH_PENDING.cabecera?'Listo · '+chDateLabel(CH_PENDING.cabecera.date):'Línea · Cumplimiento · Atrasos · Adelantos')+
          '</div>'+
        '</div>'+
        '<div class="ch-control">'+
          '<strong>Hora del corte</strong>'+
          '<input class="ch-time" type="time" value="'+esc(CH_DRAFT_TIME)+'" onchange="CH_DRAFT_TIME=this.value">'+
          '<div class="ch-file-state">Ej.: 09:15 · 13:00 · 18:00</div>'+
        '</div>'+
        '<button class="ch-save" type="button" onclick="chSaveCut()" '+(canSave?'':'disabled')+'>Guardar corte</button>'+
      '</div>'+
      cutsHtml+
    '</div>';
}
