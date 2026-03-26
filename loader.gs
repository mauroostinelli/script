/**
 * loader.gs — Carga y ejecuta el código de formato desde GitHub
 *
 * USO:
 *   1. Ejecuta  syncFromGitHub()  una vez para copiar el código en Script Properties
 *   2. El menú 'Formato corporativo' queda disponible desde onOpen()
 *
 * NOTA: Apps Script NO puede ejecutar código descargado dinámicamente (eval no funciona
 * de forma fiable). El flujo correcto es:
 *   → Descarga el código de GitHub
 *   → Lo guarda en Script Properties como caché
 *   → Muestra en log si hubo cambios respecto a la versión local
 *
 * Para EJECUTAR el código, copia el contenido de formato.gs directamente
 * en el editor de Apps Script (o usa clasp push desde local).
 */

var GITHUB_RAW_URL = 'https://raw.githubusercontent.com/mauroostinelli/script/main/formato.gs';
var PROP_KEY       = 'formato_gs_cache';
var PROP_SHA_KEY   = 'formato_gs_sha';

/**
 * Descarga formato.gs desde GitHub y guarda en Script Properties.
 * Muestra un resumen en el log: si hay cambios o ya está actualizado.
 */
function syncFromGitHub() {
  try {
    var resp = UrlFetchApp.fetch(GITHUB_RAW_URL, {
      method: 'get',
      muteHttpExceptions: true
    });

    var status = resp.getResponseCode();
    if (status !== 200) {
      Logger.log('ERROR: GitHub respondió HTTP ' + status);
      return;
    }

    var contenido = resp.getContentText();
    var shaActual  = calcularHash_(contenido);

    var props    = PropertiesService.getScriptProperties();
    var shaAntes = props.getProperty(PROP_SHA_KEY) || '';

    if (shaAntes === shaActual) {
      Logger.log('✅ El código ya está actualizado (sin cambios desde GitHub)');
      SpreadsheetApp.getActiveSpreadsheet()
        .toast('El código ya está actualizado.', 'Sincronización GitHub', 4);
      return;
    }

    props.setProperty(PROP_KEY,     contenido);
    props.setProperty(PROP_SHA_KEY, shaActual);
    props.setProperty('formato_gs_updated', new Date().toISOString());

    Logger.log('🔄 Código actualizado desde GitHub');
    Logger.log('   Tamaño: ' + contenido.length + ' caracteres');
    Logger.log('   SHA:    ' + shaActual);
    Logger.log('   Fecha:  ' + new Date().toISOString());

    SpreadsheetApp.getActiveSpreadsheet()
      .toast('✅ Código sincronizado desde GitHub. Tamano: ' + contenido.length + ' chars.', 'Sincronización', 6);

  } catch (e) {
    Logger.log('EXCEPCIÓN en syncFromGitHub: ' + e.message);
  }
}

/**
 * Muestra en log el estado actual del código cacheado.
 */
function estadoCache() {
  var props   = PropertiesService.getScriptProperties();
  var cache   = props.getProperty(PROP_KEY)     || '';
  var sha     = props.getProperty(PROP_SHA_KEY) || '(ninguno)';
  var updated = props.getProperty('formato_gs_updated') || '(nunca)';

  Logger.log('=== Estado caché GitHub ===');
  Logger.log('Tamaño almacenado : ' + cache.length + ' chars');
  Logger.log('SHA               : ' + sha);
  Logger.log('Última sync        : ' + updated);

  if (cache.length > 0) {
    Logger.log('--- Primeras 300 líneas ---');
    Logger.log(cache.substring(0, 300));
  } else {
    Logger.log('(caché vacía — ejecuta syncFromGitHub() primero)');
  }
}

/**
 * Devuelve un hash simple (suma de códigos char) para detectar cambios.
 * No es SHA-256 real pero sirve para comparar versiones.
 */
function calcularHash_(str) {
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    var c = str.charCodeAt(i);
    hash  = ((hash << 5) - hash) + c;
    hash  = hash & hash;
  }
  return hash.toString(16);
}
