// ════════════════════════════════════════════════════════════════════════════
// APPS SCRIPT — BACKEND SAR
// Pegar en: Google Sheets → Extensiones → Apps Script
// Publicar como: Aplicación web → Ejecutar como: Yo → Acceso: Cualquier usuario
// ════════════════════════════════════════════════════════════════════════════

const SHEET_USUARIOS   = "USUARIOS";
const SHEET_AERONAVES  = "AERONAVES";
const SHEET_BUQUES     = "BUQUES";
const SHEET_CONFIG     = "CONFIG";

// ── Respuesta JSON con CORS ───────────────────────────────────────────────────
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET: router de acciones ───────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';

  if (action === 'ping')        return jsonResponse({ok:true, msg:'SAR backend activo'});
  if (action === 'login')       return handleLogin(e.parameter.user, e.parameter.pass);
  if (action === 'getRegistros')return getRegistros();
  if (action === 'getPowerBI')  return getPowerBIData();
  if (action === 'getConfig')   return getConfig();

  return jsonResponse({ok:false, msg:'Acción no reconocida'});
}

// ── POST: guardar datos ───────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';

    if (action === 'addRegistro') return addRegistro(body.data);
    if (action === 'addUsuario')  return addUsuario(body.data);
    if (action === 'updateConfig')return updateConfig(body.data);

    return jsonResponse({ok:false, msg:'Acción POST no reconocida'});
  } catch(err) {
    return jsonResponse({ok:false, msg:'Error: '+err.message});
  }
}

// ════════════════════════════════════════════════════════════════════════════
// AUTENTICACIÓN
// ════════════════════════════════════════════════════════════════════════════
function handleLogin(usuario, pass) {
  if (!usuario || !pass) return jsonResponse({ok:false, msg:'Credenciales incompletas'});

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ws   = ss.getSheetByName(SHEET_USUARIOS);
  if (!ws) return jsonResponse({ok:false, msg:'Hoja USUARIOS no encontrada'});

  const data = ws.getDataRange().getValues();
  // Columnas: 0=usuario, 1=pass, 2=nombre, 3=rol, 4=base, 5=estado, 6=tabs(JSON)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[5] !== 'ACTIVO') continue;
    if (String(row[0]).toLowerCase() === usuario.toLowerCase() && String(row[1]) === pass) {
      // Determinar tabs según rol
      const rolTabs = {
        'AÉREO':   ['carga-aire','dashboard','disponibilidad','service','actividades','bases'],
        'NAVAL':   ['carga-buques','dashboard','disponibilidad','service','actividades','bases'],
        'CONTROL': ['dashboard','disponibilidad','service','actividades','bases'],
        'ADMIN':   ['carga-aire','carga-buques','dashboard','disponibilidad','service','actividades','bases','admin','config'],
      };
      const rol   = String(row[3]);
      const tabs  = row[6] ? JSON.parse(row[6]) : (rolTabs[rol] || ['dashboard']);
      const color = {AÉREO:'#4899EE',NAVAL:'#2ECC71',CONTROL:'#C49A3A',ADMIN:'#E74C3C'}[rol] || '#888';
      return jsonResponse({
        ok: true,
        user: {
          usuario: row[0], nombre: row[2], rol, base: row[4],
          color, tabs, timestamp: new Date().toISOString()
        }
      });
    }
  }
  return jsonResponse({ok:false, msg:'Credenciales incorrectas'});
}

// ════════════════════════════════════════════════════════════════════════════
// REGISTROS DE CARGA
// ════════════════════════════════════════════════════════════════════════════
function addRegistro(reg) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const tab = reg.tipo === 'NAVAL' ? SHEET_BUQUES : SHEET_AERONAVES;
  let   ws  = ss.getSheetByName(tab);

  if (!ws) {
    ws = ss.insertSheet(tab);
    // Headers
    const hdrs = reg.tipo === 'NAVAL'
      ? ['TIMESTAMP','FECHA','USUARIO','ID','CLASE','BASE','CONDICION','HS_HOY','MI_HOY',
         'HS_ACUM','MI_ACUM','INT_HS','INT_MI','ACTIVIDAD','ES_SAR','ZONA','TRIPULANTES',
         'ESTADO_SVC','DISP_SAR','OBSERVACIONES']
      : ['TIMESTAMP','FECHA','USUARIO','ID','MODELO','BASE','CONDICION','HS_HOY','HS_ACUM',
         'INTERVAL','HS_REM','PCT_USO','ACTIVIDAD','ES_SAR','ORIGEN','DESTINO','HORARIO',
         'TRIPULANTES','HS_MES','ESTADO_SVC','DISP_SAR','OBSERVACIONES'];
    ws.appendRow(hdrs);
    ws.getRange(1,1,1,hdrs.length).setFontWeight('bold').setBackground('#0D1B3E').setFontColor('#C8A84B');
    ws.setFrozenRows(1);
  }

  const ts = new Date().toISOString();

  if (reg.tipo === 'NAVAL') {
    ws.appendRow([
      ts, reg.fecha, reg.usuario||'', reg.id, reg.clase||'', reg.base||'',
      reg.cond, reg.hsHoy||0, reg.miHoy||0, reg.hsAcum||0, reg.miAcum||0,
      reg.intHs||250, reg.intMi||500, reg.actividad||'', reg.esSAR||'N',
      reg.zona||'', reg.tripulantes||0, reg.estado||'', reg.dispSAR||'',
      reg.obs||''
    ]);
  } else {
    ws.appendRow([
      ts, reg.fecha, reg.usuario||'', reg.id, reg.modelo||'', reg.base||'',
      reg.cond, reg.hsHoy||0, reg.hsAcum||0, reg.interval||100,
      reg.hsRem||'', reg.pctUso||'', reg.actividad||'', reg.esSAR||'N',
      reg.origen||'', reg.destino||'', reg.horario||'', reg.tripulantes||0,
      reg.hsMes||0, reg.estado||'', reg.dispSAR||'', reg.obs||''
    ]);
  }

  // Colorear fila según condición
  const lastRow = ws.getLastRow();
  const condColors = {ALFA:'#D5F5E3',BRAVO:'#FEFDE0',CHARLIE:'#FDEBD0',ZULU:'#FADBD8'};
  const bg = condColors[reg.cond] || '#FFFFFF';
  ws.getRange(lastRow, 1, 1, ws.getLastColumn()).setBackground(bg);

  return jsonResponse({ok:true, row:lastRow, timestamp:ts});
}

function getRegistros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registros = [];
  const condiciones = {};

  for (const tab of [SHEET_AERONAVES, SHEET_BUQUES]) {
    const ws = ss.getSheetByName(tab);
    if (!ws || ws.getLastRow() < 2) continue;
    const data = ws.getDataRange().getValues();
    const hdrs = data[0].map(h => String(h).toLowerCase());

    for (let i = 1; i < data.length; i++) {
      const row = {};
      hdrs.forEach((h, j) => row[h] = data[i][j]);
      const r = {
        tipo:      tab === SHEET_BUQUES ? 'NAVAL' : 'AÉREO',
        id:        row['id'] || '',
        fecha:     formatDate(row['fecha']),
        usuario:   row['usuario'] || '',
        cond:      row['condicion'] || '',
        hsHoy:     row['hs_hoy'] || 0,
        miHoy:     row['mi_hoy'] || 0,
        hsAcum:    row['hs_acum'] || 0,
        miAcum:    row['mi_acum'] || 0,
        interval:  row['interval'] || row['int_hs'] || 100,
        actividad: row['actividad'] || '',
        esSAR:     row['es_sar'] || 'N',
        origen:    row['origen'] || '',
        destino:   row['destino'] || '',
        obs:       row['observaciones'] || '',
        estado:    row['estado_svc'] || '',
        dispSAR:   row['disp_sar'] || '',
      };
      registros.push(r);

      // Actualizar condición más reciente por medio
      if (!condiciones[r.id] || r.fecha > (condiciones[r.id].fecha || '')) {
        condiciones[r.id] = {
          cond:     r.cond,
          hsAcum:   r.hsAcum,
          interval: r.interval,
          ult:      r.actividad,
          fecha:    r.fecha,
        };
      }
    }
  }

  return jsonResponse({ok:true, registros, condiciones});
}

// ════════════════════════════════════════════════════════════════════════════
// ENDPOINT POWER BI — devuelve JSON estructurado para consumo directo
// ════════════════════════════════════════════════════════════════════════════
function getPowerBIData() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const out = { aeronaves:[], buques:[], actividades:[], generado: new Date().toISOString() };

  // Aeronaves
  const wsA = ss.getSheetByName(SHEET_AERONAVES);
  if (wsA && wsA.getLastRow() > 1) {
    const data = wsA.getDataRange().getValues();
    const hdrs = data[0].map(h => String(h).toLowerCase());
    for (let i = 1; i < data.length; i++) {
      const row = {};
      hdrs.forEach((h,j) => row[h] = data[i][j]);
      out.aeronaves.push({
        fecha:     formatDate(row['fecha']),
        id:        row['id'],
        modelo:    row['modelo'],
        base:      row['base'],
        condicion: row['condicion'],
        hs_hoy:    row['hs_hoy'] || 0,
        hs_acum:   row['hs_acum'] || 0,
        interval:  row['interval'] || 100,
        hs_rem:    row['hs_rem'] || '',
        pct_uso:   row['pct_uso'] || '',
        actividad: row['actividad'] || '',
        es_sar:    row['es_sar'] || 'N',
        origen:    row['origen'] || '',
        destino:   row['destino'] || '',
        estado_svc:row['estado_svc'] || '',
        disp_sar:  row['disp_sar'] || '',
        mes:       new Date(row['fecha']).getMonth()+1,
        anio:      new Date(row['fecha']).getFullYear(),
      });
      out.actividades.push({tipo:'AÉREO',id:row['id'],fecha:formatDate(row['fecha']),
        actividad:row['actividad']||'',es_sar:row['es_sar']||'N',
        origen:row['origen']||'',destino:row['destino']||'',hs_hoy:row['hs_hoy']||0});
    }
  }

  // Buques
  const wsB = ss.getSheetByName(SHEET_BUQUES);
  if (wsB && wsB.getLastRow() > 1) {
    const data = wsB.getDataRange().getValues();
    const hdrs = data[0].map(h => String(h).toLowerCase());
    for (let i = 1; i < data.length; i++) {
      const row = {};
      hdrs.forEach((h,j) => row[h] = data[i][j]);
      out.buques.push({
        fecha:     formatDate(row['fecha']),
        id:        row['id'],
        clase:     row['clase'],
        base:      row['base'],
        condicion: row['condicion'],
        hs_hoy:    row['hs_hoy'] || 0,
        mi_hoy:    row['mi_hoy'] || 0,
        hs_acum:   row['hs_acum'] || 0,
        mi_acum:   row['mi_acum'] || 0,
        estado_svc:row['estado_svc'] || '',
        disp_sar:  row['disp_sar'] || '',
        actividad: row['actividad'] || '',
        es_sar:    row['es_sar'] || 'N',
        mes:       new Date(row['fecha']).getMonth()+1,
        anio:      new Date(row['fecha']).getFullYear(),
      });
      out.actividades.push({tipo:'NAVAL',id:row['id'],fecha:formatDate(row['fecha']),
        actividad:row['actividad']||'',es_sar:row['es_sar']||'N',
        zona:row['zona']||'',hs_hoy:row['hs_hoy']||0,mi_hoy:row['mi_hoy']||0});
    }
  }

  return jsonResponse({ok:true, data:out});
}

// ════════════════════════════════════════════════════════════════════════════
// USUARIOS
// ════════════════════════════════════════════════════════════════════════════
function addUsuario(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let   ws = ss.getSheetByName(SHEET_USUARIOS);
  if (!ws) {
    ws = ss.insertSheet(SHEET_USUARIOS);
    ws.appendRow(['USUARIO','PASS','NOMBRE','ROL','BASE','ESTADO','TABS_JSON']);
    ws.getRange(1,1,1,7).setFontWeight('bold').setBackground('#0D1B3E').setFontColor('#C8A84B');
  }
  const rolTabs = {
    'AÉREO':   JSON.stringify(['carga-aire','dashboard','disponibilidad','service','actividades','bases']),
    'NAVAL':   JSON.stringify(['carga-buques','dashboard','disponibilidad','service','actividades','bases']),
    'CONTROL': JSON.stringify(['dashboard','disponibilidad','service','actividades','bases']),
    'ADMIN':   JSON.stringify(['carga-aire','carga-buques','dashboard','disponibilidad','service','actividades','bases','admin','config']),
  };
  ws.appendRow([data.usuario, data.pass, data.nombre, data.rol, data.base||'', data.estado||'ACTIVO', rolTabs[data.rol]||'[]']);
  return jsonResponse({ok:true});
}

// ════════════════════════════════════════════════════════════════════════════
// CONFIG DE INTERVALOS
// ════════════════════════════════════════════════════════════════════════════
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(SHEET_CONFIG);
  if (!ws) return jsonResponse({ok:true, config:[]});
  const data = ws.getDataRange().getValues();
  const hdrs = data[0].map(h => String(h).toLowerCase());
  const config = data.slice(1).map(row => {
    const obj = {};
    hdrs.forEach((h,j) => obj[h] = row[j]);
    return obj;
  });
  return jsonResponse({ok:true, config});
}

// ════════════════════════════════════════════════════════════════════════════
// HELPER
// ════════════════════════════════════════════════════════════════════════════
function formatDate(val) {
  if (!val) return '';
  try {
    const d = new Date(val);
    return isNaN(d) ? String(val) : d.toISOString().split('T')[0];
  } catch(e) { return String(val); }
}

// ════════════════════════════════════════════════════════════════════════════
// CONFIGURACIÓN INICIAL — ejecutar UNA SOLA VEZ
// Crea las solapas y agrega los usuarios iniciales
// ════════════════════════════════════════════════════════════════════════════
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Crear solapa USUARIOS
  let wsU = ss.getSheetByName(SHEET_USUARIOS);
  if (!wsU) wsU = ss.insertSheet(SHEET_USUARIOS);
  wsU.clearContents();
  wsU.appendRow(['USUARIO','PASS','NOMBRE','ROL','BASE','ESTADO','TABS_JSON']);
  wsU.getRange(1,1,1,7).setFontWeight('bold').setBackground('#0D1B3E').setFontColor('#C8A84B');
  // Usuarios iniciales — CAMBIAR CONTRASEÑAS antes de usar en producción
  const usuarios = [
    ['aire.eaba',    'AireEABA2026',    'Responsable EABA',      'AÉREO',   'Buenos Aires', 'ACTIVO', '["carga-aire","dashboard","disponibilidad","service","actividades","bases"]'],
    ['aire.seav',    'AireSEAV2026',    'Responsable SEAV',      'AÉREO',   'Ezeiza',       'ACTIVO', '["carga-aire","dashboard","disponibilidad","service","actividades","bases"]'],
    ['aire.eacr',    'AireEACR2026',    'Responsable EACR',      'AÉREO',   'Cte. Rivadavia','ACTIVO','["carga-aire","dashboard","disponibilidad","service","actividades","bases"]'],
    ['aire.eamp',    'AireEAMP2026',    'Responsable EAMP',      'AÉREO',   'Mar del Plata', 'ACTIVO','["carga-aire","dashboard","disponibilidad","service","actividades","bases"]'],
    ['naval.central','Naval2026!',      'Responsable Buques',    'NAVAL',   'Central',      'ACTIVO', '["carga-buques","dashboard","disponibilidad","service","actividades","bases"]'],
    ['control.sar',  'Control2026!',    'Coordinador SAR',       'CONTROL', 'Central',      'ACTIVO', '["dashboard","disponibilidad","service","actividades","bases"]'],
    ['admin.pna',    'Admin@PNA2026',   'Administrador Sistema', 'ADMIN',   'Central',      'ACTIVO', '["carga-aire","carga-buques","dashboard","disponibilidad","service","actividades","bases","admin","config"]'],
  ];
  usuarios.forEach(u => wsU.appendRow(u));
  wsU.setFrozenRows(1);
  wsU.getRange(2,1,usuarios.length,7).setBackground('#F5F5F5');
  wsU.autoResizeColumns(1,7);

  // Crear solapa CONFIG
  let wsC = ss.getSheetByName(SHEET_CONFIG);
  if (!wsC) wsC = ss.insertSheet(SHEET_CONFIG);
  wsC.clearContents();
  wsC.appendRow(['ID_MEDIO','MODELO','TIPO','HS_INTERVAL','MI_INTERVAL','NOTAS']);
  wsC.getRange(1,1,1,6).setFontWeight('bold').setBackground('#0D1B3E').setFontColor('#C8A84B');
  const intervals = [
    ['PA-40','Dauphin','Helicóptero',100,'','Ajustar según manual'],
    ['PA-41','Dauphin','Helicóptero',100,'','Ajustar según manual'],
    ['PA-43','Dauphin','Helicóptero',100,'','Ajustar según manual'],
    ['PA-94','Schweizer','Helicóptero',100,'','Ajustar según manual'],
    ['PA-22','Beechcraft','Avión',100,'','Ajustar según manual'],
    ['PA-102','Piper','Avión',100,'','Ajustar según manual'],
    ['PA-105','Piper','Avión',100,'','Ajustar según manual'],
    ['PA-112','Diamond','Avión',100,'','Ajustar según manual'],
    ['PA-11','Puma EC-225','Helicóptero',200,'','Según manual EC-225'],
    ['GC-24 MANTILLA','Guardacostas T24','Naval','',250,'Ajustar según manual'],
    ['GC-25 AZOPARDO','Guardacostas T24','Naval','',250,'Ajustar según manual'],
    ['GC-26 THOMPSON','Guardacostas T24','Naval','',250,'Ajustar según manual'],
    ['GC-27 PREFECTO FIQUE','Guardacostas','Naval','',250,'Ajustar según manual'],
    ['GC-28 PREFECTO DERBES','Guardacostas','Naval','',250,'Ajustar según manual'],
    ['GC-189 PREFECTO GARCIA','Guardacostas','Naval','',250,'Ajustar según manual'],
  ];
  intervals.forEach(r => wsC.appendRow(r));
  wsC.setFrozenRows(1); wsC.autoResizeColumns(1,6);

  Logger.log('✓ Setup completado. Solapas USUARIOS y CONFIG creadas con datos iniciales.');
  Logger.log('⚠ IMPORTANTE: Cambiar contraseñas antes de usar en producción.');
}
