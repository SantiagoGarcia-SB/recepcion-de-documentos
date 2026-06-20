/**
 * ═══════════════════════════════════════════════════════════════
 *  RECEPCIÓN DE REESTUDIOS — EL LIBERTADOR (PROCESO CORREO)
 *  Backend Google Apps Script
 * ═══════════════════════════════════════════════════════════════
 */

// ─────────────────────────────────────────────────────────────────
//  CONFIGURACIÓN
// ─────────────────────────────────────────────────────────────────
const CONFIG = {
  SPREADSHEET_ID: '1jGa30nF7DTlu6bRoU8cOBqU8c_AP-bq6LP8D52JpaPQ',
  SHEET_NAME:     'Solicitudes',
  FOLDER_ID:      '0APmEW6M8yBcVUk9PVA',
};

// NUEVA CONFIGURACIÓN: Hoja destino de consolidado
const ID_SHEET_DESTINO = '1slgykTgjoAtCd6KmlG7Lqiuw-nM1hSguQbi0XqeLu7U';
const NOMBRE_HOJA_DESTINO = 'ORIGEN'; // Sincronizado con la pestaña de tu consolidador

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
  'Email Registrador',
  'Nombre Evaluador',
  'Estado Evaluación',
  'Motivo Devolución',
  'Notificacion Gmail Chat',
  'Tipo de solicitud',
  'Estado Registro',
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
    
    // ✅ UNIFORMIDAD ABSOLUTA: Formato "Solicitud_{nro}_{fechaFormateada}"
    var fechaFormateada = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd-MM-yyyy HH.mm');
    var baseName = 'Solicitud_' + params.solicitud + '_' + fechaFormateada;
    
    // Obtener la extensión original para no corromper el archivo
    var extension = '';
    if (params.fileName && params.fileName.lastIndexOf('.') !== -1) {
      extension = params.fileName.substring(params.fileName.lastIndexOf('.'));
    }
    var newName = baseName + extension;

    var decoded = Utilities.base64Decode(params.base64);
    var blob = Utilities.newBlob(decoded, params.fileType || 'application/octet-stream', newName);

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
 * @param {Object} params - { tempFileId, chunk }
 * @returns {{success: boolean}}
 */
function appendChunkToTemp(params) {
  console.log('[appendChunkToTemp] Entrada — tempFileId: ' + params.tempFileId + ', chunkLength: ' + (params.chunk ? params.chunk.length : 0));
  try {
    var file = DriveApp.getFileById(params.tempFileId);
    var current = file.getBlob().getDataAsString();

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
    
    // ✅ UNIFORMIDAD ABSOLUTA: Formato "Solicitud_{nro}_{fechaFormateada}"
    var fechaFormateada = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd-MM-yyyy HH.mm');
    var baseName = 'Solicitud_' + params.solicitud + '_' + fechaFormateada;
    
    // Extraer extensión original para no corromper el archivo
    var extension = '';
    if (params.fileName && params.fileName.lastIndexOf('.') !== -1) {
      extension = params.fileName.substring(params.fileName.lastIndexOf('.'));
    }
    var newName = baseName + extension;

    var decoded = Utilities.base64Decode(fullBase64);
    var blob = Utilities.newBlob(decoded, params.fileType || 'application/octet-stream', newName);

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

function submitFormData(data) {
  console.log('[submitFormData] Entrada — data: ' + JSON.stringify(data));
  try {
    var tipoEstudio = (data.tipoEstudio || '').toString().trim();
    var tiposArray = tipoEstudio.split(',').map(function(t) { return t.trim(); }).filter(function(t) { return t.length > 0; });

    if (tiposArray.length === 0) {
      console.log('[submitFormData] No se recibieron tipos de estudio');
      return { success: false, error: 'Debes seleccionar al menos un tipo de estudio.' };
    }

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
      '',  // Nombre Evaluador
      '',  // Estado Evaluación
      '',  // Motivo Devolución
      '',  // Notificacion Gmail Chat
      (data.tipoSolicitud || '').toString().trim(),
      'Pendiente',
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
      userEmail: uData.email,
      tipoSolicitud: (data.tipoSolicitud || '').toString().trim()
    }));

    // 1. Guarda en el Sheet Principal
    sheet.appendRow(rowData);
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 6).setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.getRange(lastRow, 8).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    // 2. Resultado (Pendiente — NO escribe en ORIGEN ni envía correo)
    var result = {
      success:    true,
      id:         nextId,
      rowNumber:  lastRow,
      userName:   uData.name,
      userEmail:  uData.email,
      ts:         Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
      filesCount: filesCount,
      folderUrl:  folderUrl,
    };
    console.log('[submitFormData] Registro pendiente — resultado: ' + JSON.stringify(result));
    return result;

  } catch (e) {
    console.error('[submitFormData] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    Logger.log('[submitFormData] ERROR — message: ' + e.message + ' | stack: ' + (e.stack || 'N/A'));
    return { success: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────
//  CONFIRMAR REGISTRO (Fase 2: escribe en ORIGEN + envía correo)
// ─────────────────────────────────────────────────────────────────

function confirmRegistro(registroId) {
  console.log('[confirmRegistro] Confirmando registro ID: ' + registroId);
  try {
    var sheet = _getOrCreateSheet();
    var rowNum = _findRowById(sheet, registroId);
    if (rowNum === -1) {
      return { success: false, error: 'Registro no encontrado: ' + registroId };
    }

    var rowData = sheet.getRange(rowNum, 1, 1, HEADERS.length).getValues()[0];
    var estadoActual = rowData[15]; // columna P = Estado Registro
    if (estadoActual === 'Confirmado') {
      return { success: false, error: 'Este registro ya fue confirmado.' };
    }

    // Actualizar estado a Confirmado
    sheet.getRange(rowNum, 16).setValue('Confirmado');

    // Leer datos de la fila
    var solicitudStr = String(rowData[1]);
    var polizaStr    = String(rowData[2]);
    var tipoEstudio  = String(rowData[3]);
    var filesCount   = parseInt(rowData[4]) || 0;
    var arrivalDate  = rowData[5] instanceof Date ? rowData[5] : (rowData[5] ? new Date(rowData[5]) : null);
    var folderUrl    = String(rowData[6]);
    var registroDate = rowData[7];
    var userName     = String(rowData[8]);
    var userEmail    = String(rowData[9]);
    var tipoSolicitud = String(rowData[14]);

    var tiposArray = tipoEstudio.split(',').map(function(t) { return t.trim(); }).filter(function(t) { return t.length > 0; });

    // Lógica de negocio para consolidador
    var claseDeSolicitud = "Reestudio";
    var tipoDeProceso = "Anexo";
    if (tiposArray.indexOf("Nueva UAR") !== -1) {
      claseDeSolicitud = "Nueva";
    } else if (tiposArray.indexOf("Deudor UAR") !== -1) {
      claseDeSolicitud = "Adicional";
    } else if (tiposArray.indexOf("Biometría Fallida") !== -1) {
      claseDeSolicitud = "Biometría Fallida";
    }

    // Escribir en ORIGEN
    try {
      var destSs = SpreadsheetApp.openById(ID_SHEET_DESTINO);
      var destSheet = destSs.getSheetByName(NOMBRE_HOJA_DESTINO) || destSs.getSheets()[0];
      destSheet.appendRow([arrivalDate || new Date(), solicitudStr, folderUrl, "CORREO", claseDeSolicitud, tipoDeProceso]);
      console.log('[confirmRegistro] Registro agregado a hoja de consolidado.');
    } catch (destErr) {
      console.error('[confirmRegistro] Error al guardar en hoja destino: ' + destErr.message);
    }

    // Enviar correo de confirmación
    var result = {
      success:    true,
      id:         registroId,
      userName:   userName,
      userEmail:  userEmail,
      ts:         registroDate instanceof Date
        ? Utilities.formatDate(registroDate, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
        : String(registroDate),
      filesCount: filesCount,
      folderUrl:  folderUrl,
    };

    try {
      _sendConfirmationEmail(userEmail, result, {
        solicitud:      solicitudStr,
        poliza:         polizaStr,
        tipoEstudio:    tipoEstudio,
        filesCount:     filesCount,
        arrivalDateStr: arrivalDate
          ? Utilities.formatDate(arrivalDate, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
          : '—',
      }, tiposArray);
    } catch (mailErr) {
      console.error('[confirmRegistro] Error enviando correo: ' + mailErr.message);
    }

    console.log('[confirmRegistro] Registro confirmado exitosamente.');
    return result;

  } catch (e) {
    console.error('[confirmRegistro] ERROR: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────
//  ACTUALIZAR REGISTRO PENDIENTE (Corrección de datos)
// ─────────────────────────────────────────────────────────────────

function updateRegistro(data) {
  console.log('[updateRegistro] Actualizando registro ID: ' + data.registroId);
  try {
    var sheet = _getOrCreateSheet();
    var rowNum = _findRowById(sheet, data.registroId);
    if (rowNum === -1) {
      return { success: false, error: 'Registro no encontrado: ' + data.registroId };
    }

    var estadoActual = sheet.getRange(rowNum, 16).getValue();
    if (estadoActual === 'Confirmado') {
      return { success: false, error: 'No se puede editar un registro ya confirmado.' };
    }

    var solicitudStr  = (data.solicitud || '').toString().trim();
    var polizaStr     = (data.poliza || '').toString().trim();
    var tipoEstudio   = (data.tipoEstudio || '').toString().trim();
    var filesCount    = parseInt(data.filesCount) || 0;
    var arrivalDate   = data.fechaHora ? new Date(data.fechaHora) : null;
    var tipoSolicitud = (data.tipoSolicitud || '').toString().trim();

    // Actualizar campos en la fila
    sheet.getRange(rowNum, 2).setValue(solicitudStr);        // B: Solicitud
    sheet.getRange(rowNum, 3).setValue(polizaStr);           // C: Póliza
    sheet.getRange(rowNum, 4).setValue(tipoEstudio);         // D: Tipo Estudio
    sheet.getRange(rowNum, 5).setValue(filesCount);          // E: Archivos
    sheet.getRange(rowNum, 6).setValue(arrivalDate);         // F: Fecha Llegada
    sheet.getRange(rowNum, 6).setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.getRange(rowNum, 15).setValue(tipoSolicitud);      // O: Tipo Solicitud

    // Si cambió el número de solicitud, actualizar carpeta y URL
    var folder = _getOrCreateSolicitudFolder(solicitudStr);
    var folderUrl = folder.getUrl();
    sheet.getRange(rowNum, 7).setValue(folderUrl);           // G: URL Carpeta

    var uData = getUserData();
    var result = {
      success:    true,
      id:         data.registroId,
      rowNumber:  rowNum,
      userName:   uData.name,
      userEmail:  uData.email,
      ts:         Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
      filesCount: filesCount,
      folderUrl:  folderUrl,
    };
    console.log('[updateRegistro] Registro actualizado: ' + JSON.stringify(result));
    return result;

  } catch (e) {
    console.error('[updateRegistro] ERROR: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────
//  LIMPIAR ARCHIVOS DE UNA CARPETA DE SOLICITUD
// ─────────────────────────────────────────────────────────────────

function clearSolicitudFiles(solicitud) {
  console.log('[clearSolicitudFiles] Limpiando archivos de solicitud: ' + solicitud);
  try {
    var parentFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    var folderName = 'SOLICITUD -' + solicitud;
    var iter = parentFolder.searchFolders('title = "' + folderName + '" and trashed = false');

    if (!iter.hasNext()) {
      return { success: true, deletedCount: 0 };
    }

    var folder = iter.next();
    var files = folder.getFiles();
    var count = 0;
    while (files.hasNext()) {
      files.next().setTrashed(true);
      count++;
    }
    console.log('[clearSolicitudFiles] Archivos eliminados: ' + count);
    return { success: true, deletedCount: count };
  } catch (e) {
    console.error('[clearSolicitudFiles] ERROR: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ═════════════════════════════════════════════════════════════════
//  HELPERS
// ═════════════════════════════════════════════════════════════════

function _findRowById(sheet, id) {
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      return i + 2; // +2 porque empezamos en fila 2 (después del header)
    }
  }
  return -1;
}

function _capitalize(str) {
  if (!str) return '';
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

function _getOrCreateSolicitudFolder(solicitudStr) {
  console.log('[_getOrCreateSolicitudFolder] Buscando/creando carpeta para solicitud: ' + solicitudStr);
  var parentFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  
  // 1. FORMATO DE NOMBRE ESTANDARIZADO
  var folderName   = 'SOLICITUD -' + solicitudStr;
  
  // 2. BÚSQUEDA MEJORADA (Evita duplicados y lee papeleras)
  var iteradorCarpetas = parentFolder.searchFolders('title = "' + folderName + '" and trashed = false');
  if (iteradorCarpetas.hasNext()) {
    var existing = iteradorCarpetas.next();
    console.log('[_getOrCreateSolicitudFolder] Carpeta existente encontrada — id: ' + existing.getId());
    return existing;
  }

  // 3. CREACIÓN SI NO EXISTE
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
  console.log('[_setupSheetHeaders] Antes de headerRange.setValues — ' + JSON.stringify({ headers: HEADERS }));
  headerRange.setValues([HEADERS]);
  headerRange.setFontWeight('bold');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setBackground('#253150');
  headerRange.setHorizontalAlignment('center');
  var widths = [90, 160, 140, 260, 130, 200, 300, 180, 200, 220, 160, 160, 200, 180, 160, 140];
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

/**
 * Envía correo de confirmación al registrador con el resumen de la solicitud.
 * Usa colores corporativos: Rojo #BD0F14, Navy #253150, Gris #706F6F.
 */
function _sendConfirmationEmail(toEmail, result, payload, tiposArray) {
  var firstName = (result.userName || 'Usuario').split(' ')[0];

  var rows = [
    ['ID de Registro',          '#' + result.id],
    ['N.º de Solicitud',        payload.solicitud],
    ['N.º de Póliza',           payload.poliza],
    ['Tipo de Estudio',         payload.tipoEstudio],
    ['Archivos adjuntos',       payload.filesCount + ' archivo(s)'],
    ['Fecha llegada correo',    payload.arrivalDateStr],
    ['Registrado por',          result.userName],
    ['Email',                   result.userEmail],
    ['Fecha de registro',       result.ts],
  ];

  var rowsHtml = rows.map(function(r) {
    return '<tr>' +
      '<td style="padding:12px 16px;color:#706F6F;font-size:13px;font-weight:500;border-bottom:1px solid #f1f5f9;white-space:nowrap;vertical-align:top;">' + r[0] + '</td>' +
      '<td style="padding:12px 16px;color:#253150;font-size:13px;font-weight:600;border-bottom:1px solid #f1f5f9;text-align:right;vertical-align:top;">' + r[1] + '</td>' +
    '</tr>';
  }).join('');

  var folderRow = '<tr>' +
    '<td style="padding:12px 16px;color:#706F6F;font-size:13px;font-weight:500;white-space:nowrap;">Carpeta de Solicitud</td>' +
    '<td style="padding:12px 16px;font-size:13px;font-weight:600;text-align:right;">' +
      '<a href="' + result.folderUrl + '" style="color:#BD0F14;text-decoration:underline;font-weight:700;">Abrir carpeta en Drive →</a>' +
    '</td>' +
  '</tr>';

  var html = '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"></head>' +
    '<body style="margin:0;padding:0;background:#f8fafc;font-family:Inter,Arial,Helvetica,sans-serif;">' +

    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f8fafc;padding:40px 16px;">' +
    '<tr><td align="center">' +
    '<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(37,49,80,0.08);">' +

    // ── Header corporativo rojo ──
    '<tr><td style="background:#BD0F14;padding:32px 36px;">' +
      '<table width="100%" cellpadding="0" cellspacing="0"><tr>' +
        '<td>' +
          '<p style="margin:0;color:#ffffff;font-size:20px;font-weight:800;letter-spacing:-0.3px;">El Libertador</p>' +
          '<p style="margin:6px 0 0;color:rgba(255,255,255,0.75);font-size:12px;font-weight:500;letter-spacing:0.3px;">Recepción de solicitudes UAR y reestudios</p>' +
        '</td>' +
        '<td align="right" valign="middle">' +
          '<span style="display:inline-block;background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.25);border-radius:8px;padding:6px 12px;color:rgba(255,255,255,0.95);font-size:10px;font-weight:700;letter-spacing:0.8px;text-transform:uppercase;">Módulo de Radicación</span>' +
        '</td>' +
      '</tr></table>' +
    '</td></tr>' +

    // ── Indicador de éxito ──
    '<tr><td style="padding:36px 36px 0;text-align:center;">' +
      '<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">' +
        '<div style="width:72px;height:72px;background:#f0fdf4;border:2px solid #86efac;border-radius:50%;margin:0 auto 20px;text-align:center;line-height:72px;">' +
          '<span style="font-size:36px;line-height:72px;">✅</span>' +
        '</div>' +
      '</td></tr></table>' +
      '<h2 style="margin:0 0 8px;color:#253150;font-size:22px;font-weight:800;">¡Excelente trabajo, ' + firstName + '!</h2>' +
      '<p style="margin:0;color:#706F6F;font-size:14px;line-height:1.5;">La solicitud fue registrada y guardada correctamente en el sistema.</p>' +
    '</td></tr>' +

    // ── Tabla resumen ──
    '<tr><td style="padding:28px 36px 12px;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f8fafc;border-radius:12px;border:1px solid #e2e8f0;overflow:hidden;">' +
        '<tr><td colspan="2" style="background:#253150;padding:12px 16px;">' +
          '<p style="margin:0;color:#ffffff;font-size:12px;font-weight:700;letter-spacing:0.5px;text-transform:uppercase;">Resumen del Registro</p>' +
        '</td></tr>' +
        rowsHtml + folderRow +
      '</table>' +
    '</td></tr>' +

    // ── Mensaje informativo ──
    '<tr><td style="padding:12px 36px 28px;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="background:#FFF5F5;border-radius:8px;border:1px solid #FECACA;padding:14px 16px;">' +
        '<tr><td style="padding:14px 16px;">' +
          '<p style="margin:0;color:#BD0F14;font-size:12px;font-weight:600;">Próximos pasos</p>' +
          '<p style="margin:6px 0 0;color:#706F6F;font-size:12px;line-height:1.5;">Tu solicitud será asignada al equipo correspondiente, gracias por tu gestión</p>' +
        '</td></tr>' +
      '</table>' +
    '</td></tr>' +

    // ── Footer navy ──
    '<tr><td style="background:#253150;padding:24px 36px;">' +
      '<table width="100%" cellpadding="0" cellspacing="0"><tr>' +
        '<td>' +
          '<p style="margin:0;color:rgba(255,255,255,0.9);font-size:12px;font-weight:700;">El Libertador</p>' +
          '<p style="margin:4px 0 0;color:rgba(255,255,255,0.5);font-size:11px;">Servicio al Cliente — Uso interno</p>' +
        '</td>' +
        '<td align="right" valign="middle">' +
          '<p style="margin:0;color:rgba(255,255,255,0.4);font-size:10px;">Este correo fue generado automáticamente.<br>No es necesario responder.</p>' +
        '</td>' +
      '</tr></table>' +
    '</td></tr>' +

    '</table></td></tr></table>' +
    '</body></html>';

  var subject = '✅ Solicitud ' + payload.solicitud + ' registrada para asignación';

  MailApp.sendEmail({ to: toEmail, subject: subject, htmlBody: html, name: 'Radicación · El Libertador' });
  console.log('[_sendConfirmationEmail] Correo enviado a: ' + toEmail);
}
