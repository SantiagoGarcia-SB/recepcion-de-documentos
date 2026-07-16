/**
 * ═══════════════════════════════════════════════════════════════
 *  RESPALDO DOCUMENTOS INFOBIP — VICTORIA
 *  Trigger: cada 1 minuto
 * ═══════════════════════════════════════════════════════════════
 */

// ─────────────────────────────────────────────────────────────────
//  CONFIGURACIÓN VICTORIA
// ─────────────────────────────────────────────────────────────────
const INFOBIP_API_KEY = PropertiesService.getScriptProperties().getProperty('INFOBIP_API_KEY') || '';

const ID_SHEET_VICTORIA = '1_wSkdh3eD0mG474De6RUrj9yd9L8SKnnSjqO3Pg4Jsg';
const NOMBRE_HOJA_VICTORIA = 'Anexar documentos a la solicitud';

// Columnas de la hoja Victoria (estructura real)
const COL_FECHA_VICTORIA       = 1;   // A - Fecha ingreso
const COL_NRO_SOLICITUD_VICTORIA = 2; // B - Número de solicitud
const COL_URL_INFOBIP          = 5;   // E - documento_adjunto (URL Infobip)
const COL_OBSERVACIONES_VICTORIA = 17; // Q - Observaciones
const COL_LINK_DRIVE_VICTORIA  = 18;  // R - Link Drive (URL carpeta)

// ─────────────────────────────────────────────────────────────────
//  FUNCIÓN PRINCIPAL
// ─────────────────────────────────────────────────────────────────

function respaldarDocumentosInfobip() {
  var libro = SpreadsheetApp.openById(ID_SHEET_VICTORIA);
  var hoja = libro.getSheetByName(NOMBRE_HOJA_VICTORIA);

  if (!hoja) {
    Logger.log('[Victoria] No se encontró la hoja origen: ' + NOMBRE_HOJA_VICTORIA);
    return;
  }

  // Abrir hoja destino (consolidado ORIGEN)
  var hojaDestino = null;
  try {
    var libroDestino = SpreadsheetApp.openById(ID_SHEET_DESTINO);
    hojaDestino = libroDestino.getSheetByName(NOMBRE_HOJA_DESTINO) || libroDestino.getSheets()[0];
  } catch (e) {
    Logger.log('[Victoria] Error al abrir la hoja destino: ' + e.message);
  }

  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  var datos = hoja.getRange(2, 1, ultimaFila - 1, COL_LINK_DRIVE_VICTORIA).getValues();
  var procesados = 0;
  var errores = 0;

  // Carpeta raíz (misma que usa el módulo de correo)
  var carpetaRaiz = DriveApp.getFolderById(CONFIG.FOLDER_ID);

  for (var i = 0; i < datos.length; i++) {
    var nroSolicitud     = datos[i][COL_NRO_SOLICITUD_VICTORIA - 1];
    var fechaIngreso     = datos[i][COL_FECHA_VICTORIA - 1];
    var urlInfobip       = datos[i][COL_URL_INFOBIP - 1];
    var observacion      = datos[i][COL_OBSERVACIONES_VICTORIA - 1];
    var urlDriveGuardada = datos[i][COL_LINK_DRIVE_VICTORIA - 1];

    // Omitir filas ya marcadas como inválidas
    if (observacion && observacion.toString().trim() === 'Número de solicitud invalido') {
      continue;
    }

    // Solo procesar filas con link Infobip y sin link Drive
    if (!urlInfobip || !urlInfobip.toString().includes('infobip.com') || urlDriveGuardada) {
      continue;
    }

    // Validar que la URL sea un link real (empieza con http)
    var urlStr = urlInfobip.toString().trim();
    if (!urlStr.startsWith('http')) {
      continue;
    }

    // Validar número de solicitud (no vacío, sin letras)
    var nroStr = nroSolicitud ? nroSolicitud.toString().trim() : '';

    // Si vienen múltiples números (separados por salto de línea, coma o espacio),
    // tomar solo el primero y anotar los demás en observaciones
    var numerosMultiples = nroStr.split(/[\n\r,;]+/).map(function(n) { return n.trim(); }).filter(function(n) { return n !== ''; });
    var obsMultiple = '';
    if (numerosMultiples.length > 1) {
      nroStr = numerosMultiples[0];
      obsMultiple = 'Números adicionales recibidos: ' + numerosMultiples.slice(1).join(', ');
    }

    var esInvalido = nroStr === '' || /[a-zA-ZáéíóúÁÉÍÓÚñÑ]/.test(nroStr);

    if (esInvalido) {
      hoja.getRange(i + 2, COL_OBSERVACIONES_VICTORIA).setValue('Número de solicitud invalido');
      errores++;
      continue;
    }

    // Si había múltiples números, registrar en observaciones
    if (obsMultiple) {
      hoja.getRange(i + 2, COL_OBSERVACIONES_VICTORIA).setValue(obsMultiple);
    }

    // Formatear fecha
    var fechaFormateada = '';
    if (fechaIngreso instanceof Date) {
      fechaFormateada = Utilities.formatDate(fechaIngreso, Session.getScriptTimeZone(), 'dd-MM-yyyy HH.mm');
    } else {
      fechaFormateada = fechaIngreso.toString().replace(/[\/:\*?"<>|]/g, '-');
    }

    // Buscar o crear carpeta de solicitud
    var nombreCarpeta = 'SOLICITUD -' + nroStr;
    var carpetaSolicitud = _obtenerOCrearCarpetaVictoria(carpetaRaiz, nombreCarpeta);

    // Descargar y guardar archivo
    var nombreArchivo = 'Solicitud_' + nroStr + '_' + fechaFormateada;
    var resultado = _descargarArchivoInfobip(urlInfobip, nombreArchivo, carpetaSolicitud);

    if (resultado === 'OK') {
      var urlCarpeta = carpetaSolicitud.getUrl();

      // Guardar URL de la carpeta en la hoja Victoria
      hoja.getRange(i + 2, COL_LINK_DRIVE_VICTORIA).setValue(urlCarpeta);

      // Escribir en hoja ORIGEN del consolidado (18 columnas)
      if (hojaDestino) {
        hojaDestino.appendRow([
          fechaIngreso,        // 1. fechaRadicacion
          nroStr,              // 2. solicitud
          urlCarpeta,          // 3. linkDrive
          'VICTORIA',          // 4. origen
          'Anexo',             // 5. tipoDeProceso
          'Reestudio',         // 6. claseDeSolicitud
          '',                  // 7. analistaAsignado
          '',                  // 8. nombreAnalista
          '',                  // 9. fechaAsignacion
          '',                  // 10. fechaFinGestion
          '',                  // 11. estadoGestion
          '',                  // 12. motivoAplazamiento
          '',                  // 13. motivoNegacion
          '',                  // 14. observaciones
          '',                  // 15. minutos_cola
          '',                  // 16. minutos_gestion
          '',                  // 17. minutos_general
          '',                  // 18. poliza
        ]);
      }

      procesados++;
      Utilities.sleep(1000);
    } else {
      errores++;
    }
  }

  Logger.log('[Victoria] Proceso finalizado. Procesados: ' + procesados + ' | Errores: ' + errores);
}

// ─────────────────────────────────────────────────────────────────
//  HELPERS VICTORIA
// ─────────────────────────────────────────────────────────────────

function _obtenerOCrearCarpetaVictoria(carpetaPadre, nombreCarpeta) {
  var iter = carpetaPadre.searchFolders('title = "' + nombreCarpeta + '" and trashed = false');
  if (iter.hasNext()) {
    return iter.next();
  }
  return carpetaPadre.createFolder(nombreCarpeta);
}

function _descargarArchivoInfobip(url, nombreArchivo, carpetaDestino) {
  try {
    if (!url || !url.toString().trim().startsWith('http')) {
      Logger.log('[Victoria] URL inválida para ' + nombreArchivo + ': ' + url);
      return 'URL inválida';
    }

    var opciones = {
      method: 'get',
      headers: {
        'Authorization': 'App ' + INFOBIP_API_KEY,
        'Accept': '*/*'
      },
      muteHttpExceptions: true
    };

    var respuesta = UrlFetchApp.fetch(url, opciones);
    var codigo = respuesta.getResponseCode();

    if (codigo === 200) {
      var blob = respuesta.getBlob();
      var contentType = respuesta.getHeaders()['Content-Type'] || '';

      var extension = '';
      if (contentType.includes('application/pdf'))                         extension = '.pdf';
      else if (contentType.includes('image/jpeg'))                         extension = '.jpg';
      else if (contentType.includes('image/png'))                          extension = '.png';
      else if (contentType.includes('image/webp'))                         extension = '.webp';
      else if (contentType.includes('application/msword'))                 extension = '.doc';
      else if (contentType.includes('application/vnd.openxmlformats-officedocument.wordprocessingml.document')) extension = '.docx';
      else if (contentType.includes('application/vnd.ms-excel'))           extension = '.xls';
      else if (contentType.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))       extension = '.xlsx';

      blob.setName(nombreArchivo + extension);
      carpetaDestino.createFile(blob);
      return 'OK';
    } else {
      Logger.log('[Victoria] Error descargando ' + nombreArchivo + ' — código: ' + codigo);
      return 'Error (' + codigo + ')';
    }
  } catch (e) {
    Logger.log('[Victoria] Error de ejecución con ' + nombreArchivo + ': ' + e.message);
    return 'Error en Script';
  }
}
