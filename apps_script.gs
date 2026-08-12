const FOLDER_ID = '1P5RZe6yAOmJMf4rdUv2ly0Hz3nbEjsTF';
const MAX_FILE_SIZE = 10 * 1024 * 1024;

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return jsonResponse({ ok: true, message: 'HWDaily upload service is running' });
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const expectedToken = PropertiesService.getScriptProperties().getProperty('UPLOAD_TOKEN');
    if (!expectedToken || payload.token !== expectedToken) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }

    const action = String(payload.action || 'upload').toLowerCase();
    if (action === 'upload') {
      return uploadFile(payload);
    }
    if (action === 'delete') {
      return trashFile(payload);
    }
    if (action === 'download') {
      return downloadFile(payload);
    }
    return jsonResponse({ ok: false, error: 'Unsupported action' });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error.message || error) });
  }
}

function uploadFile(payload) {
  const fileName = String(payload.fileName || '').trim();
  const encodedData = String(payload.data || '');
  const mimeType = String(payload.mimeType || 'application/octet-stream');
  if (!fileName || !encodedData) {
    return jsonResponse({ ok: false, error: '缺少檔名或檔案內容' });
  }

  const bytes = Utilities.base64Decode(encodedData);
  if (bytes.length > MAX_FILE_SIZE) {
    return jsonResponse({ ok: false, error: '檔案不可超過 10 MB' });
  }

  const folder = DriveApp.getFolderById(FOLDER_ID);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  return jsonResponse({
    ok: true,
    id: file.getId(),
    name: file.getName(),
    url: file.getUrl(),
  });
}

function trashFile(payload) {
  const fileId = String(payload.fileId || '').trim();
  if (!fileId) {
    return jsonResponse({ ok: false, error: '缺少檔案 ID' });
  }

  const file = DriveApp.getFileById(fileId);
  if (!fileBelongsToUploadFolder(file)) {
    return jsonResponse({ ok: false, error: '拒絕刪除：檔案不在指定附件資料夾' });
  }

  const fileName = file.getName();
  file.setTrashed(true);
  return jsonResponse({ ok: true, id: fileId, name: fileName, trashed: true });
}

function downloadFile(payload) {
  const fileId = String(payload.fileId || '').trim();
  if (!fileId) {
    return jsonResponse({ ok: false, error: '缺少檔案 ID' });
  }

  const file = DriveApp.getFileById(fileId);
  if (!fileBelongsToUploadFolder(file)) {
    return jsonResponse({ ok: false, error: '拒絕下載：檔案不在指定附件資料夾' });
  }

  const blob = file.getBlob();
  const bytes = blob.getBytes();
  if (bytes.length > MAX_FILE_SIZE) {
    return jsonResponse({ ok: false, error: '檔案不可超過 10 MB' });
  }
  return jsonResponse({
    ok: true,
    id: fileId,
    name: file.getName(),
    mimeType: blob.getContentType() || 'application/octet-stream',
    data: Utilities.base64Encode(bytes),
  });
}

function fileBelongsToUploadFolder(file) {
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === FOLDER_ID) {
      return true;
    }
  }
  return false;
}
