
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

function chMetricCard(label,rows,kind){
  const avg=chWeighted(rows,'cumplimiento');
  const atraso=chWeighted(rows,'atraso');
  const adelanto=chWeighted(rows,'adelanto');
  return '<div class="ch-summary-card '+kind+'">'+
    '<div class="eyebrow">'+esc(label)+'</div>'+
    '<div class="big '+chClass(avg)+'">'+chFmt(avg)+'</div>'+
    '<div class="sub">Promedio ponderado · Atraso '+chFmt(atraso)+' · Adelanto '+chFmt(adelanto)+'</div>'+
    '</div>';
}

function chTable(title,rows,kind){
  const body=(rows||[]).map(function(row){
    return '<tr class="'+chRowClass(row.cumplimiento)+'">'+
      '<td class="ch-line">'+esc(row.linea)+'</td>'+
      '<td><span class="ch-score '+chClass(row.cumplimiento)+'">'+chFmt(row.cumplimiento)+'</span></td>'+
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

function chShareText(){
  const cuts=chActiveCuts();
  const date=CH_STATE.activeDate;
  if(!cuts.length)return '';

  const parts=['TSJ · Cumplimiento horario y cabecera · '+chDateLabel(date)];

  cuts.forEach(function(cut){
    parts.push(
      '',
      'Corte '+cut.time,
      'C. Horario: '+chFmt(chWeighted(cut.horario,'cumplimiento')),
      '<90% Horario: '+chCriticalText(cut.horario),
      'C. Cabecera: '+chFmt(chWeighted(cut.cabecera,'cumplimiento')),
      '<90% Cabecera: '+chCriticalText(cut.cabecera)
    );
  });

  return parts.join('\n');
}

function shareCHWhatsApp(){
  const text=chShareText();
  if(!text){
    alert('No hay cortes cargados.');
    return;
  }
  window.open('https://wa.me/?text='+encodeURIComponent(text),'_blank','noopener');
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

function exportCHorarioPDF(){
  if(!chActiveCuts().length){
    alert('No hay cortes cargados.');
    return;
  }

  const previousTitle=document.title;
  document.title='TSJ_Cumplimiento_'+(CH_STATE.activeDate||'dia');
  document.body.classList.add('print-chorario');

  const cleanup=function(){
    document.body.classList.remove('print-chorario');
    document.title=previousTitle;
  };

  window.addEventListener('afterprint',cleanup,{once:true});
  window.print();
  setTimeout(cleanup,3000);
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
      return '<article class="ch-cut">'+
        '<div class="ch-cut-head">'+
          '<div class="ch-cut-title"><b>Corte '+(index+1)+' · '+esc(cut.time)+'</b><span>'+chDateLabel(CH_STATE.activeDate)+'</span></div>'+
          '<button class="ch-delete" type="button" onclick="chDeleteCut(\''+cut.id+'\')">Eliminar</button>'+
        '</div>'+
        '<div class="ch-summary">'+
          chMetricCard('Cumplimiento Horario',cut.horario,'horario')+
          chMetricCard('Cumplimiento Cabecera',cut.cabecera,'cabecera')+
        '</div>'+
        '<div class="ch-grid">'+
          chTable('C. Horario',cut.horario,'horario')+
          chTable('C. Cabecera',cut.cabecera,'cabecera')+
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
        '<button class="ch-action whatsapp" type="button" onclick="shareCHWhatsApp()">WhatsApp</button>'+
        '<button class="ch-action primary" type="button" onclick="exportCHorarioPDF()">PDF</button>'+
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
