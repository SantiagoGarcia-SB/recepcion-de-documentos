/**
 * ═══════════════════════════════════════════════════════════════
 *  RECEPCIÓN DE REESTUDIOS — EL LIBERTADOR
 *  Backend Google Apps Script
 * ═══════════════════════════════════════════════════════════════
 */

// ─────────────────────────────────────────────────────────────────
//  CONFIGURACIÓN
// ─────────────────────────────────────────────────────────────────
const CONFIG = {
  SPREADSHEET_ID: '1jGa30nF7DTlu6bRoU8cOBqU8c_AP-bq6LP8D52JpaPQ',
  SHEET_NAME:     'Solicitudes',
  FOLDER_ID:      '18W_7bpKOMQrY4YJrrcqPw9Jr1G93pVa0',
};

const TIPOS_ESTUDIO = {
  'Aceptación LMI':                           true,
  'Actas':                                    true,
  'Actualizar Resultado':                     true,
  'Ampliación Canon':                         true,
  'Anular Estudio':                           true,
  'AVS Cerrada Por 85':                       false,
  'AVS Re Asignado':                          false,
  'AVS Traslado Cerrada Por 86':              false,
  'Biometría Fallida':                        false,
  'Cámara De Comercio':                       true,
  'Cambio De Inmueble':                       true,
  'Cambio De Roles':                          true,
  'Confirmación Destino':                     true,
  'Desistimiento Y Continuidad De Solicitud': true,
  'Deudor UAR':                               true,
  'Disminución Canon':                        true,
  'Extractos Y/O DR':                         true,
  'Opción 1+1':                               true,
  'Opción CDT':                               true,
  'Nueva UAR':                                true,
  'Paz Y Salvo':                              true,
  'Reactivar Deudor':                         true,
  'Reconsideración':                          true,
  'Retiro De Deudor':                         true,
  'Retoma Pendiente':                         true,
  'Traslado':                                 true,
  'Unificación De Solicitudes':               true,
};

const HEADERS = [
  'ID Registro',
  'Número de Solicitud',
  'Número de Póliza',
  'Tipo de Estudio',
  'Archivos Adjuntos',
  'Fecha y Hora de Llegada del Correo',
  'URL Carpeta de Solicitud',
  'Fecha y Hora de Registro',
  'Registrado Por (Nombre)',
  'Email Asignador',
];

// ═════════════════════════════════════════════════════════════════
//  FUNCIONES PÚBLICAS
// ═════════════════════════════════════════════════════════════════

function doGet() {
  console.log('[doGet] Iniciando renderizado de la webapp');
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Recepción de Reestudios — El Libertador')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  console.log('[include] Incluyendo archivo: ' + filename);
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getUserData() {
  console.log('[getUserData] Obteniendo datos del usuario activo');
  var email = _getUser();
  var localPart = email.split('@')[0] || '';
  var parts = localPart.split('.');
  var firstName = _capitalize(parts[0] || '');
  var lastName  = _capitalize(parts[1] || '');
  var fullName  = (firstName + ' ' + lastName).trim();
  var initials  = ((firstName.charAt(0) || '') + (lastName.charAt(0) || '')).toUpperCase() || 'U';
  var result = { email: email, name: fullName, initials: initials };
  console.log('[getUserData] Resultado: ' + JSON.stringify(result));
  return result;
}

function getTiposEstudio() {
  console.log('[getTiposEstudio] Retornando lista de tipos de estudio');
  return Object.keys(TIPOS_ESTUDIO).map(function(key) {
    return { nombre: key, requiereAnexo: TIPOS_ESTUDIO[key] };
  });
}

// ─────────────────────────────────────────────────────────────────
//  UPLOAD DE ARCHIVOS
//  Estrategia: cada archivo se sube individualmente en una sola llamada.
//  El frontend envía el base64 completo por archivo.
//  Para archivos > 30MB, el frontend los divide en partes y usa
//  uploadChunkToTemp + finalizeUpload.
// ─────────────────────────────────────────────────────────────────

/**
 * Sube un archivo completo (para archivos <= 30 MB aprox).
 * @param {Object} params - { base64, fileName, fileType, solicitud, tipoEstudio }
 * @returns {{success: boolean}}
 */
function uploadSingleFile(params) {
  console.log('[uploadSingleFile] Entrada — params: ' + JSON.stringify({
    fileName: params.fileName,
    fileType: params.fileType,
    solicitud: params.solicitud,
    tipoEstudio: params.tipoEstudio,
    base64Length: params.base64 ? params.base64.length : 0
  }));
  try {
    var folder = _getOrCreateSolicitudFolder(params.solicitud);
    var now = new Date();
    var datePrefix = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    var newName = datePrefix + ' - ' + params.tipoEstudio + ' - ' + params.fileName;

    var decoded = Utilities.base64Decode(params.base64);
    var blob = Utilities.newBlob(decoded, params.fileType || 'application/octet-stream', newName);

    // LOG CRÍTICO: Datos antes de crear archivo en Drive
    console.log('[uploadSingleFile] Antes de folder.createFile — ' + JSON.stringify({
      folderName: folder.getName(),
      folderId: folder.getId(),
      blobName: newName,
      blobType: params.fileType || 'application/octet-stream',
      decodedBytes: decoded.length
    }));

    var file = folder.createFile(blob);

    console.log('[uploadSingleFile] Archivo creado exitosamente — fileId: ' + file.getId());
    return { success: true };
  } catch (e) {
    console.error('[uploadSingleFile] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    Logger.log('[uploadSingleFile] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    return { success: false, error: e.message };
  }
}

/**
 * Crea un archivo temporal para upload por partes (archivos grandes).
 * @param {Object} params - { uploadId }
 * @returns {{success: boolean, tempFileId: string}}
 */
function createTempFile(params) {
  console.log('[createTempFile] Entrada — params: ' + JSON.stringify(params));
  try {
    var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);

    // LOG CRÍTICO: Datos antes de crear archivo temporal
    console.log('[createTempFile] Antes de folder.createFile — ' + JSON.stringify({
      folderName: folder.getName(),
      folderId: CONFIG.FOLDER_ID,
      tempFileName: '_temp_' + params.uploadId
    }));

    var tempFile = folder.createFile('_temp_' + params.uploadId, '', 'text/plain');
    console.log('[createTempFile] Archivo temporal creado — tempFileId: ' + tempFile.getId());
    return { success: true, tempFileId: tempFile.getId() };
  } catch (e) {
    console.error('[createTempFile] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    Logger.log('[createTempFile] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    return { success: false, error: e.message };
  }
}

/**
 * Appenda un chunk de base64 al archivo temporal.
 * Usa DriveApp API para append real sin leer todo el archivo.
 * @param {Object} params - { tempFileId, chunk }
 * @returns {{success: boolean}}
 */
function appendChunkToTemp(params) {
  console.log('[appendChunkToTemp] Entrada — tempFileId: ' + params.tempFileId + ', chunkLength: ' + (params.chunk ? params.chunk.length : 0));
  try {
    var file = DriveApp.getFileById(params.tempFileId);
    // Leer contenido actual y concatenar
    var current = file.getBlob().getDataAsString();

    // LOG CRÍTICO: Datos antes de setContent
    console.log('[appendChunkToTemp] Antes de file.setContent — ' + JSON.stringify({
      tempFileId: params.tempFileId,
      currentLength: current.length,
      chunkLength: params.chunk ? params.chunk.length : 0,
      newTotalLength: current.length + (params.chunk ? params.chunk.length : 0)
    }));

    file.setContent(current + params.chunk);
    console.log('[appendChunkToTemp] Chunk agregado exitosamente');
    return { success: true };
  } catch (e) {
    console.error('[appendChunkToTemp] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    Logger.log('[appendChunkToTemp] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    return { success: false, error: e.message };
  }
}

/**
 * Finaliza el upload: lee el temp, decodifica, crea archivo final, borra temp.
 * @param {Object} params - { tempFileId, fileName, fileType, solicitud, tipoEstudio }
 * @returns {{success: boolean}}
 */
function finalizeUpload(params) {
  console.log('[finalizeUpload] Entrada — params: ' + JSON.stringify({
    tempFileId: params.tempFileId,
    fileName: params.fileName,
    fileType: params.fileType,
    solicitud: params.solicitud,
    tipoEstudio: params.tipoEstudio
  }));
  try {
    var tempFile = DriveApp.getFileById(params.tempFileId);
    var fullBase64 = tempFile.getBlob().getDataAsString();
    console.log('[finalizeUpload] Base64 leído del temp — longitud: ' + fullBase64.length);

    var folder = _getOrCreateSolicitudFolder(params.solicitud);
    var now = new Date();
    var datePrefix = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    var newName = datePrefix + ' - ' + params.tipoEstudio + ' - ' + params.fileName;

    var decoded = Utilities.base64Decode(fullBase64);
    var blob = Utilities.newBlob(decoded, params.fileType || 'application/octet-stream', newName);

    // LOG CRÍTICO: Datos antes de crear archivo final en Drive
    console.log('[finalizeUpload] Antes de folder.createFile — ' + JSON.stringify({
      folderName: folder.getName(),
      folderId: folder.getId(),
      blobName: newName,
      blobType: params.fileType || 'application/octet-stream',
      decodedBytes: decoded.length
    }));

    var file = folder.createFile(blob);

    console.log('[finalizeUpload] Archivo final creado — fileId: ' + file.getId());

    tempFile.setTrashed(true);
    console.log('[finalizeUpload] Archivo temporal eliminado');
    return { success: true };
  } catch (e) {
    console.error('[finalizeUpload] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    Logger.log('[finalizeUpload] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    return { success: false, error: e.message };
  }
}

/**
 * Registra la solicitud en el Sheet (los archivos ya están en Drive).
 */
function submitFormData(data) {
  console.log('[submitFormData] Entrada — data: ' + JSON.stringify(data));
  try {
    var tipoEstudio = (data.tipoEstudio || '').toString().trim();
    var tiposArray = tipoEstudio.split(',').map(function(t) { return t.trim(); }).filter(function(t) { return t.length > 0; });

    if (tiposArray.length === 0) {
      console.log('[submitFormData] No se recibieron tipos de estudio');
      return { success: false, error: 'Debes seleccionar al menos un tipo de estudio.' };
    }

    // Validar que todos los tipos existan
    for (var i = 0; i < tiposArray.length; i++) {
      if (!TIPOS_ESTUDIO.hasOwnProperty(tiposArray[i])) {
        console.log('[submitFormData] Tipo de estudio no reconocido: ' + tiposArray[i]);
        return { success: false, error: 'Tipo de estudio no reconocido: ' + tiposArray[i] };
      }
    }

    var requiereAnexo = tiposArray.some(function(tipo) { return TIPOS_ESTUDIO[tipo]; });
    var filesCount = parseInt(data.filesCount) || 0;

    if (requiereAnexo && filesCount === 0) {
      console.log('[submitFormData] Validación fallida — requiere anexo pero filesCount=0');
      return { success: false, error: 'Al menos un tipo de estudio seleccionado requiere documento anexo.' };
    }

    var sheet    = _getOrCreateSheet();
    var nextId   = _getNextId(sheet);
    var uData    = getUserData();
    var now      = new Date();

    var solicitudStr = (data.solicitud || '').toString().trim();
    var folder       = _getOrCreateSolicitudFolder(solicitudStr);
    var folderUrl    = folder.getUrl();
    var arrivalDate  = data.fechaHora ? new Date(data.fechaHora) : null;

    // LOG CRÍTICO: Datos exactos antes de appendRow
    var rowData = [
      nextId,
      solicitudStr,
      (data.poliza || '').toString().trim(),
      tipoEstudio,
      filesCount,
      arrivalDate,
      folderUrl,
      now,
      uData.name,
      uData.email,
    ];
    console.log('[submitFormData] Antes de sheet.appendRow — ' + JSON.stringify({
      nextId: nextId,
      solicitud: solicitudStr,
      poliza: (data.poliza || '').toString().trim(),
      tipoEstudio: tipoEstudio,
      filesCount: filesCount,
      arrivalDate: arrivalDate ? arrivalDate.toISOString() : null,
      folderUrl: folderUrl,
      now: now.toISOString(),
      userName: uData.name,
      userEmail: uData.email
    }));

    sheet.appendRow(rowData);

    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 6).setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.getRange(lastRow, 8).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    var result = {
      success:    true,
      id:         nextId,
      userName:   uData.name,
      userEmail:  uData.email,
      ts:         Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
      filesCount: filesCount,
      folderUrl:  folderUrl,
    };
    console.log('[submitFormData] Registro exitoso — resultado: ' + JSON.stringify(result));
    return result;
  } catch (e) {
    console.error('[submitFormData] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    Logger.log('[submitFormData] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    return { success: false, error: e.message };
  }
}

// ═════════════════════════════════════════════════════════════════
//  HELPERS
// ═════════════════════════════════════════════════════════════════

function _capitalize(str) {
  if (!str) return '';
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

function _getOrCreateSolicitudFolder(solicitudStr) {
  console.log('[_getOrCreateSolicitudFolder] Buscando/creando carpeta para solicitud: ' + solicitudStr);
  var parentFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  var folderName   = 'Solicitud ' + solicitudStr;
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    var existing = folders.next();
    console.log('[_getOrCreateSolicitudFolder] Carpeta existente encontrada — id: ' + existing.getId());
    return existing;
  }

  // LOG CRÍTICO: Antes de crear carpeta nueva
  console.log('[_getOrCreateSolicitudFolder] Antes de parentFolder.createFolder — ' + JSON.stringify({
    parentFolderId: CONFIG.FOLDER_ID,
    newFolderName: folderName
  }));

  var newFolder = parentFolder.createFolder(folderName);
  console.log('[_getOrCreateSolicitudFolder] Carpeta creada — id: ' + newFolder.getId());
  return newFolder;
}

function _getOrCreateSheet() {
  console.log('[_getOrCreateSheet] Abriendo spreadsheet: ' + CONFIG.SPREADSHEET_ID + ', hoja: ' + CONFIG.SHEET_NAME);
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    console.log('[_getOrCreateSheet] Hoja no encontrada, creando...');
    var firstSheet = ss.getSheets()[0];
    if (firstSheet && firstSheet.getLastRow() === 0) {
      firstSheet.setName(CONFIG.SHEET_NAME);
      sheet = firstSheet;
    } else {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    }
    _setupSheetHeaders(sheet);
  } else if (sheet.getLastRow() === 0) {
    console.log('[_getOrCreateSheet] Hoja vacía, configurando headers...');
    _setupSheetHeaders(sheet);
  }
  return sheet;
}

function _setupSheetHeaders(sheet) {
  console.log('[_setupSheetHeaders] Configurando encabezados de la hoja');
  var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);

  // LOG CRÍTICO: Datos antes de setValues en headers
  console.log('[_setupSheetHeaders] Antes de headerRange.setValues — ' + JSON.stringify({ headers: HEADERS }));

  headerRange.setValues([HEADERS]);
  headerRange.setFontWeight('bold');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setBackground('#253150');
  headerRange.setHorizontalAlignment('center');
  var widths = [90, 160, 140, 260, 130, 200, 300, 180, 200, 220];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.setFrozenRows(1);
}

function _getNextId(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  var lastId = sheet.getRange(lastRow, 1).getValue();
  var nextId = (parseInt(lastId) || 0) + 1;
  console.log('[_getNextId] lastRow: ' + lastRow + ', lastId: ' + lastId + ', nextId: ' + nextId);
  return nextId;
}

function _getUser() {
  try {
    var email = Session.getActiveUser().getEmail();
    var result = email || Session.getEffectiveUser().getEmail() || 'usuario@desconocido.com';
    console.log('[_getUser] Email obtenido: ' + result);
    return result;
  } catch (e) {
    console.error('[_getUser] ERROR obteniendo email — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    return 'usuario@desconocido.com';
  }
}
