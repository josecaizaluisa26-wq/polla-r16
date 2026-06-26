// ============================================================
// GOOGLE APPS SCRIPT – POLLA REAL COHOLICOS R16 2026
// Nueva polla desde Ronda de 16
// ============================================================

// ══════════════════════════════════════════════════════════════
// EQUIPOS — Reemplaza cuando se conozcan los clasificados
// ══════════════════════════════════════════════════════════════
const TEAMS_R16 = [
  'Clasificado 1A','Clasificado 2B','Clasificado 1B','Clasificado 2A',
  'Clasificado 1C','Clasificado 2D','Clasificado 1D','Clasificado 2C',
  'Clasificado 1E','Clasificado 2F','Clasificado 1F','Clasificado 2E',
  'Clasificado 1G','Clasificado 2H','Clasificado 1H','Clasificado 2G',
  'Clasificado 1I','Clasificado 2J','Clasificado 1J','Clasificado 2I',
  'Clasificado 1K','Clasificado 2L','Clasificado 1L','Clasificado 2K',
  'Mejor 3ro 1','Mejor 3ro 2','Mejor 3ro 3','Mejor 3ro 4',
  'Mejor 3ro 5','Mejor 3ro 6','Mejor 3ro 7','Mejor 3ro 8'
];

const MATCHES_R16 = [
  ['R16_1','Clasificado 1A','Clasificado 2B'],
  ['R16_2','Clasificado 1B','Clasificado 2A'],
  ['R16_3','Clasificado 1C','Clasificado 2D'],
  ['R16_4','Clasificado 1D','Clasificado 2C'],
  ['R16_5','Clasificado 1E','Clasificado 2F'],
  ['R16_6','Clasificado 1F','Clasificado 2E'],
  ['R16_7','Clasificado 1G','Clasificado 2H'],
  ['R16_8','Clasificado 1H','Clasificado 2G'],
  ['R16_9','Mejor 3ro 1','Mejor 3ro 2'],
  ['R16_10','Mejor 3ro 3','Mejor 3ro 4'],
  ['R16_11','Mejor 3ro 5','Mejor 3ro 6'],
  ['R16_12','Mejor 3ro 7','Mejor 3ro 8'],
  ['R16_13','Clasificado 1I','Clasificado 2J'],
  ['R16_14','Clasificado 1J','Clasificado 2I'],
  ['R16_15','Clasificado 1K','Clasificado 2L'],
  ['R16_16','Clasificado 1L','Clasificado 2K'],
  ['QF_1','Ganador R16_1','Ganador R16_2'],
  ['QF_2','Ganador R16_3','Ganador R16_4'],
  ['QF_3','Ganador R16_5','Ganador R16_6'],
  ['QF_4','Ganador R16_7','Ganador R16_8'],
  ['QF_5','Ganador R16_9','Ganador R16_10'],
  ['QF_6','Ganador R16_11','Ganador R16_12'],
  ['QF_7','Ganador R16_13','Ganador R16_14'],
  ['QF_8','Ganador R16_15','Ganador R16_16'],
  ['SF_1','Ganador QF_1','Ganador QF_2'],
  ['SF_2','Ganador QF_3','Ganador QF_4'],
  ['SF_3','Ganador QF_5','Ganador QF_6'],
  ['SF_4','Ganador QF_7','Ganador QF_8'],
  ['TERCERO','Perdedor SF_1','Perdedor SF_2'],
  ['FINAL','Ganador SF_1','Ganador SF_2'],
];

const POINTS = {exact:3, winner:1, champ:8, goleador:5};

// ── HANDLER ──────────────────────────────────────────────────
function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const action = (e&&e.parameter&&e.parameter.action)?e.parameter.action:'';
    let result = {};
    if      (action==='ping')                  result={ok:true,msg:'API R16 activa ✅'};
    else if (action==='getParticipantes')      result={ok:true,data:getParticipantes_()};
    else if (action==='addParticipante')       result=addParticipante_(e.parameter.nombre||'',e.parameter.wa||'');
    else if (action==='getPredictions')        result={ok:true,data:getPredictions_()};
    else if (action==='savePredictionDirect')  result=savePredictionDirect_(e);
    else if (action==='saveResultados')        result=saveResultados_(JSON.parse(decodeURIComponent(e.parameter.data||'{}')));
    else if (action==='saveRanking')           result=saveRanking_(JSON.parse(decodeURIComponent(e.parameter.data||'[]')));
    else if (action==='getRanking')            result={ok:true,data:getRanking_()};
    else if (action==='updatePuntos')          {updatePuntos_(e.parameter.id||'',Number(e.parameter.pts)||0);result={ok:true};}
    else if (action==='calcularRanking')       result={ok:true,data:calcularYGuardarRanking()};
    else                                        result={ok:true,msg:'API R16 activa'};
    output.setContent(JSON.stringify(result));
  } catch(err) {
    output.setContent(JSON.stringify({ok:false,error:err.toString()}));
  }
  return output;
}

// ── PARTICIPANTES ─────────────────────────────────────────────
function getSheet_(name){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(name);
  if(!sh)sh=ss.insertSheet(name);
  return sh;
}

function getParticipantes_(){
  const sh=getSheet_('Participantes');
  const lastRow=sh.getLastRow();
  if(lastRow<=1)return[];
  return sh.getRange(1,1,lastRow,5).getValues().slice(1).map(function(r){
    return{id:String(r[0]),nombre:String(r[1]),wa:String(r[2]),
           predicciones:r[3]===true||r[3]==='TRUE',pts:Number(r[4])||0};
  });
}

function findOrCreateParticipante_(nombre){
  const sh=getSheet_('Participantes');
  if(sh.getLastRow()===0){
    sh.appendRow(['ID','Nombre','WhatsApp','Predicciones','Puntos']);
    sh.getRange(1,1,1,5).setFontWeight('bold');
  }
  const lastRow=sh.getLastRow();
  if(lastRow>1){
    const data=sh.getRange(2,1,lastRow-1,2).getValues();
    for(var i=0;i<data.length;i++){
      if(String(data[i][1]).trim().toLowerCase()===nombre.toLowerCase())return String(data[i][0]);
    }
  }
  const id='p'+new Date().getTime();
  sh.appendRow([id,nombre,'',false,0]);
  return id;
}

function addParticipante_(nombre,wa){
  if(!nombre)return{ok:false,msg:'Nombre requerido'};
  const id=findOrCreateParticipante_(nombre);
  return{ok:true,id:id};
}

// ── PREDICCIONES ──────────────────────────────────────────────
function getPredictions_(){
  const sh=getSheet_('Predicciones');
  const lastRow=sh.getLastRow();
  if(lastRow<=1)return{};
  const result={};
  sh.getRange(1,1,lastRow,2).getValues().slice(1).forEach(function(r){
    try{result[String(r[0])]=JSON.parse(String(r[1]));}catch(e){}
  });
  return result;
}

function savePrediction_(pid,predData){
  if(!pid)return{ok:false};
  const sh=getSheet_('Predicciones');
  if(sh.getLastRow()===0){
    sh.appendRow(['ID_Participante','Prediccion_JSON','Fecha']);
    sh.getRange(1,1,1,3).setFontWeight('bold');
  }
  const lastRow=sh.getLastRow();
  if(lastRow>1){
    const data=sh.getRange(2,1,lastRow-1,1).getValues();
    for(var i=0;i<data.length;i++){
      if(String(data[i][0])===pid){
        sh.getRange(i+2,2,1,2).setValues([[JSON.stringify(predData),new Date()]]);
        markPrediccion_(pid,true);
        return{ok:true,updated:true};
      }
    }
  }
  sh.appendRow([pid,JSON.stringify(predData),new Date()]);
  markPrediccion_(pid,true);
  return{ok:true,created:true};
}

function savePredictionDirect_(e){
  const nombre=decodeURIComponent(e.parameter.nombre||'');
  const wa=decodeURIComponent(e.parameter.wa||'');
  const data=JSON.parse(decodeURIComponent(e.parameter.data||'{}'));
  const pid=findOrCreateParticipante_(nombre);
  if(wa){
    const sh=getSheet_('Participantes');
    const lastRow=sh.getLastRow();
    if(lastRow>1){
      const d=sh.getRange(2,1,lastRow-1,2).getValues();
      for(var i=0;i<d.length;i++){if(String(d[i][0])===pid){sh.getRange(i+2,3).setValue(wa);break;}}
    }
  }
  return savePrediction_(pid,data);
}

function markPrediccion_(pid,val){
  const sh=getSheet_('Participantes');
  const lastRow=sh.getLastRow();
  if(lastRow<=1)return;
  const data=sh.getRange(2,1,lastRow-1,1).getValues();
  for(var i=0;i<data.length;i++){
    if(String(data[i][0])===pid){sh.getRange(i+2,4).setValue(val);break;}
  }
}

// ── RESULTADOS ────────────────────────────────────────────────
function saveResultados_(obj){
  const sh=getSheet_('Resultados');
  sh.clearContents();
  sh.appendRow(['Partido','Goles_Local','Goles_Visitante','Equipo_Local','Equipo_Visitante','Fecha']);
  sh.getRange(1,1,1,6).setFontWeight('bold');
  Object.keys(obj).forEach(function(k){
    const r=obj[k];
    sh.appendRow([k,r.h||'',r.a||'',r.tH||'',r.tA||'',new Date()]);
  });
  return{ok:true,saved:Object.keys(obj).length};
}

// ── RANKING ───────────────────────────────────────────────────
function saveRanking_(data){
  const sh=getSheet_('Ranking');
  sh.clearContents();
  sh.appendRow(['Posicion','ID','Nombre','Puntos','Campeon_Pred','Fecha']);
  sh.getRange(1,1,1,6).setFontWeight('bold');
  data.forEach(function(p,i){sh.appendRow([i+1,p.id||'',p.nombre||'',p.pts||0,p.champ||'',new Date()]);});
  return{ok:true};
}

function getRanking_(){
  const sh=getSheet_('Ranking');
  const lastRow=sh.getLastRow();
  if(lastRow<=1)return[];
  return sh.getRange(2,1,lastRow-1,5).getValues().map(function(r){
    return{pos:r[0],id:String(r[1]),nombre:String(r[2]),pts:Number(r[3])||0,champ:String(r[4])};
  });
}

function updatePuntos_(id,pts){
  const sh=getSheet_('Participantes');
  const lastRow=sh.getLastRow();
  if(lastRow<=1)return;
  const data=sh.getRange(2,1,lastRow-1,1).getValues();
  for(var i=0;i<data.length;i++){
    if(String(data[i][0])===String(id)){sh.getRange(i+2,5).setValue(pts);break;}
  }
}

// ── CALCULAR RANKING DESDE SHEETS ────────────────────────────
function calcularYGuardarRanking(){
  const resultsSheet=getSheet_('Resultados');
  const results={};
  if(resultsSheet.getLastRow()>1){
    resultsSheet.getRange(2,1,resultsSheet.getLastRow()-1,5).getValues().forEach(function(r){
      results[String(r[0])]={h:r[1],a:r[2],tH:String(r[3]),tA:String(r[4])};
    });
  }
  const preds=getPredictions_();
  const parts=getParticipantes_();

  const ranked=parts.map(function(p){
    const pred=preds[p.id]||null;
    if(!pred)return{id:p.id,nombre:p.nombre,pts:0,champ:''};
    var pts=0;

    // Puntos por marcadores
    MATCHES_R16.forEach(function(m){
      const k=m[0];
      const r=results[k];
      const pd=pred.scores?pred.scores[k]:null;
      if(!r||r.h===''||r.a==='')return;
      if(!pd||pd.h===''||pd.a==='')return;
      const rh=Number(r.h),ra=Number(r.a),ph=Number(pd.h),pa=Number(pd.a);
      if(rh===ph&&ra===pa)pts+=POINTS.exact;
      else{
        const rw=rh>ra?'h':rh<ra?'a':'d';
        const pw=ph>pa?'h':ph<pa?'a':'d';
        if(rw===pw)pts+=POINTS.winner;
      }
    });

    // Campeón
    const finalR=results['FINAL'];
    const champ=finalR&&finalR.h!==''&&finalR.a!==''?(Number(finalR.h)>Number(finalR.a)?finalR.tH:Number(finalR.a)>Number(finalR.h)?finalR.tA:null):null;
    if(champ&&pred.champ===champ)pts+=POINTS.champ;

    // Goleador
    const gReal=String(results['goleador']||'').trim().toLowerCase();
    const gPred=String(pred.goleador||'').trim().toLowerCase();
    if(gReal&&gPred&&gReal===gPred)pts+=POINTS.goleador;

    return{id:p.id,nombre:p.nombre,pts:pts,champ:pred.champ||''};
  }).sort(function(a,b){return b.pts-a.pts;});

  // Guarda en Sheets
  const sh=getSheet_('Ranking');
  sh.clearContents();
  sh.appendRow(['Posicion','ID','Nombre','Puntos','Campeon_Pred','Fecha']);
  sh.getRange(1,1,1,6).setFontWeight('bold');
  ranked.forEach(function(p,i){
    sh.appendRow([i+1,p.id,p.nombre,p.pts,p.champ,new Date()]);
    updatePuntos_(p.id,p.pts);
  });

  Logger.log('✅ Ranking R16 calculado. '+ranked.length+' participantes.');
  ranked.forEach(function(p,i){Logger.log((i+1)+'. '+p.nombre+' — '+p.pts+' pts');});
  return ranked;
}

function testScript(){
  Logger.log('Script R16 OK');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Hoja: '+ss.getName());
  Logger.log('Partidos configurados: '+MATCHES_R16.length);
}
