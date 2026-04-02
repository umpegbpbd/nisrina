const MAX_FILE_SIZE = (CONFIG.MAX_FILE_SIZE_MB || 5) * 1024 * 1024;
const ALLOWED_EXTENSIONS = ["pdf", "jpg", "jpeg", "png"];
const ALLOWED_MIME_TYPES = ["application/pdf", "image/jpeg", "image/png"];

const state = {
  records: [],
  filteredRecords: [],
  selectedFile: null,
  previewObjectUrl: null,
  isSubmitting: false,
  accessToken: null,
  tokenExpiresAt: 0,
  tokenClient: null,
  connectedEmailHint: "",
  metadataFileId: null,
  folderInfo: null
};

const dom = {
  form: document.getElementById("uploadForm"),
  namaPegawai: document.getElementById("namaPegawai"),
  jenisPenghargaan: document.getElementById("jenisPenghargaan"),
  lingkupPenghargaan: document.getElementById("lingkupPenghargaan"),
  nomorPiagam: document.getElementById("nomorPiagam"),
  tanggalPiagam: document.getElementById("tanggalPiagam"),
  tahunPerolehan: document.getElementById("tahunPerolehan"),
  fileSertifikat: document.getElementById("fileSertifikat"),
  submitBtn: document.getElementById("submitBtn"),
  resetBtn: document.getElementById("resetBtn"),
  clearFileBtn: document.getElementById("clearFileBtn"),
  selectedFileMeta: document.getElementById("selectedFileMeta"),
  imagePreview: document.getElementById("imagePreview"),
  pdfPreviewMessage: document.getElementById("pdfPreviewMessage"),
  dataTableBody: document.getElementById("dataTableBody"),
  emptyState: document.getElementById("emptyState"),
  loadingOverlay: document.getElementById("loadingOverlay"),
  loadingText: document.getElementById("loadingText"),
  alertContainer: document.getElementById("alertContainer"),
  searchInput: document.getElementById("searchInput"),
  yearFilter: document.getElementById("yearFilter"),
  lingkupFilter: document.getElementById("lingkupFilter"),
  resetFilterBtn: document.getElementById("resetFilterBtn"),
  reloadDataBtn: document.getElementById("reloadDataBtn"),
  connectDriveBtn: document.getElementById("connectDriveBtn"),
  disconnectDriveBtn: document.getElementById("disconnectDriveBtn"),
  driveStatus: document.getElementById("driveStatus")
};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  try {
    validateConfig();
    populateEmployees();
    attachEventListeners();
    await waitForGoogleIdentity();
    initGoogleAuth();
    updateDriveStatus(false);
    showAlert("Aplikasi siap. Hubungkan Google Drive dulu sebelum upload.", "info", 7000);
  } catch (error) {
    console.error(error);
    showAlert(error.message || "Gagal memuat aplikasi.", "error", 10000);
  }
}

function validateConfig() {
  if (typeof CONFIG !== "object" || CONFIG === null) {
    throw new Error("config.js tidak valid.");
  }

  const requiredKeys = ["GOOGLE_CLIENT_ID", "GOOGLE_SCOPE", "TARGET_FOLDER_ID", "METADATA_FILE_NAME"];
  for (const key of requiredKeys) {
    if (!CONFIG[key] || typeof CONFIG[key] !== "string") {
      throw new Error(`Konfigurasi ${key} belum diisi dengan benar.`);
    }
  }

  if (CONFIG.GOOGLE_CLIENT_ID.includes("PASTE_YOUR_GOOGLE_CLIENT_ID_HERE")) {
    throw new Error("GOOGLE_CLIENT_ID masih placeholder. Isi dulu di config.js.");
  }
}

function attachEventListeners() {
  dom.connectDriveBtn.addEventListener("click", async () => {
    await connectDrive(true);
  });

  dom.disconnectDriveBtn.addEventListener("click", disconnectDrive);

  dom.form.addEventListener("submit", handleSubmit);
  dom.resetBtn.addEventListener("click", resetFormCompletely);
  dom.clearFileBtn.addEventListener("click", clearFileSelection);
  dom.fileSertifikat.addEventListener("change", handleFilePreview);

  dom.tanggalPiagam.addEventListener("change", () => {
    if (dom.tanggalPiagam.value) {
      dom.tahunPerolehan.value = new Date(dom.tanggalPiagam.value).getFullYear();
      clearFieldError("tahunPerolehan");
    }
  });

  dom.searchInput.addEventListener("input", applyFilters);
  dom.yearFilter.addEventListener("change", applyFilters);
  dom.lingkupFilter.addEventListener("change", applyFilters);
  dom.resetFilterBtn.addEventListener("click", resetFilters);

  dom.reloadDataBtn.addEventListener("click", async () => {
    try {
      await ensureDriveAccess(true);
      await refreshDataFromDrive();
      showAlert("Data berhasil dimuat ulang dari Google Drive.", "success");
    } catch (error) {
      console.error(error);
      showAlert(error.message || "Gagal memuat ulang data.", "error");
    }
  });

  [
    dom.namaPegawai,
    dom.jenisPenghargaan,
    dom.lingkupPenghargaan,
    dom.nomorPiagam,
    dom.tanggalPiagam,
    dom.tahunPerolehan,
    dom.fileSertifikat
  ].forEach((el) => {
    const eventName = el.tagName === "SELECT" || el.type === "file" || el.type === "date" ? "change" : "input";
    el.addEventListener(eventName, () => clearFieldError(el.id));
  });

  dom.dataTableBody.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) return;

    const action = button.dataset.action;
    const id = button.dataset.id;
    if (!id) return;

    if (action === "view") {
      const record = state.records.find((item) => item.id === id);
      if (!record) {
        showAlert("Data tidak ditemukan.", "error");
        return;
      }
      const url = record.file_url || record.webViewLink;
      if (!url) {
        showAlert("Link file tidak tersedia.", "error");
        return;
      }
      window.open(url, "_blank", "noopener,noreferrer");
      return;
    }

    if (action === "delete") {
      await handleDelete(id);
    }
  });
}

function populateEmployees() {
  if (!Array.isArray(EMPLOYEES)) {
    throw new Error("employees.js tidak valid. EMPLOYEES harus array.");
  }

  const fragment = document.createDocumentFragment();
  EMPLOYEES.forEach((employee) => {
    const option = document.createElement("option");
    option.value = employee;
    option.textContent = employee;
    fragment.appendChild(option);
  });

  dom.namaPegawai.appendChild(fragment);
}

async function waitForGoogleIdentity() {
  const timeout = 15000;
  const interval = 150;
  const start = Date.now();

  while (Date.now() - start < timeout) {
    if (window.google && google.accounts && google.accounts.oauth2) {
      return;
    }
    await delay(interval);
  }

  throw new Error("Library Google Identity Services gagal dimuat.");
}

function initGoogleAuth() {
  state.tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.GOOGLE_CLIENT_ID,
    scope: CONFIG.GOOGLE_SCOPE,
    callback: () => {}
  });
}

async function connectDrive(interactive = true) {
  await ensureDriveAccess(interactive);
  await verifyFolderAccess();
  await refreshDataFromDrive();
  showAlert("Google Drive berhasil terhubung.", "success");
}

function disconnectDrive() {
  if (state.accessToken && window.google && google.accounts && google.accounts.oauth2) {
    try {
      google.accounts.oauth2.revoke(state.accessToken, () => {});
    } catch (error) {
      console.warn("Revoke token gagal:", error);
    }
  }

  state.accessToken = null;
  state.tokenExpiresAt = 0;
  state.metadataFileId = null;
  state.folderInfo = null;
  state.records = [];
  state.filteredRecords = [];
  renderTable();
  updateYearFilterOptions();
  updateDriveStatus(false);
  dom.disconnectDriveBtn.classList.add("hidden");
  showAlert("Koneksi Google Drive diputuskan dari sesi ini.", "info");
}

async function ensureDriveAccess(interactive = false) {
  if (state.accessToken && Date.now() < state.tokenExpiresAt - 60_000) {
    return state.accessToken;
  }

  if (!state.tokenClient) {
    throw new Error("Google OAuth belum siap.");
  }

  if (!interactive) {
    throw new Error("Google Drive belum terhubung. Klik 'Hubungkan Google Drive' dulu.");
  }

  const accessToken = await requestAccessToken(!state.accessToken);
  state.accessToken = accessToken;
  updateDriveStatus(true);
  dom.disconnectDriveBtn.classList.remove("hidden");
  return accessToken;
}

function requestAccessToken(firstTime = true) {
  return new Promise((resolve, reject) => {
    state.tokenClient.callback = (response) => {
      if (!response) {
        reject(new Error("Tidak ada respons dari Google OAuth."));
        return;
      }

      if (response.error) {
        reject(new Error(response.error_description || response.error));
        return;
      }

      state.accessToken = response.access_token;
      state.tokenExpiresAt = Date.now() + ((response.expires_in || 3600) * 1000);
      resolve(response.access_token);
    };

    try {
      state.tokenClient.requestAccessToken({
        prompt: firstTime ? "consent" : ""
      });
    } catch (error) {
      reject(error);
    }
  });
}

function updateDriveStatus(isConnected) {
  if (isConnected) {
    dom.driveStatus.textContent = state.folderInfo
      ? `Terhubung ke Google Drive: ${state.folderInfo.name}`
      : "Terhubung ke Google Drive";
    dom.driveStatus.className = "status-badge status-online";
    return;
  }

  dom.driveStatus.textContent = "Belum terhubung ke Google Drive";
  dom.driveStatus.className = "status-badge status-offline";
}

async function verifyFolderAccess() {
  setLoading(true, "Memeriksa akses folder Google Drive...");

  const folder = await driveApiRequest(
    `/files/${encodeURIComponent(CONFIG.TARGET_FOLDER_ID)}?fields=id,name,mimeType,webViewLink`,
    { method: "GET" }
  );

  if (!folder || !folder.id) {
    throw new Error("Folder target tidak ditemukan atau tidak bisa diakses.");
  }

  state.folderInfo = folder;
  updateDriveStatus(true);

  if (folder.mimeType !== "application/vnd.google-apps.folder") {
    throw new Error("TARGET_FOLDER_ID bukan folder Google Drive yang valid.");
  }

  return folder;
}

async function refreshDataFromDrive() {
  setLoading(true, "Memuat metadata dari Google Drive...");
  const data = await loadPnsBerprestasiData();
  state.records = sortRecordsNewestFirst(data);
  updateYearFilterOptions();
  applyFilters();
  setLoading(false);
}

async function loadPnsBerprestasiData() {
  const metadataFile = await findMetadataFile();
  if (!metadataFile) {
    state.metadataFileId = null;
    return [];
  }

  state.metadataFileId = metadataFile.id;

  const rawText = await driveApiRequest(`/files/${encodeURIComponent(metadataFile.id)}?alt=media`, {
    method: "GET",
    rawText: true
  });

  if (!rawText) return [];

  try {
    const parsed = JSON.parse(rawText);
    if (!Array.isArray(parsed)) {
      throw new Error();
    }
    return parsed;
  } catch {
    throw new Error("Isi file metadata JSON di Google Drive rusak atau bukan array.");
  }
}

async function findMetadataFile() {
  const query = [
    `'${escapeDriveQueryString(CONFIG.TARGET_FOLDER_ID)}' in parents`,
    `name = '${escapeDriveQueryString(CONFIG.METADATA_FILE_NAME)}'`,
    `trashed = false`
  ].join(" and ");

  const result = await driveApiRequest(
    `/files?q=${encodeURIComponent(query)}&fields=files(id,name,modifiedTime)&pageSize=10&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    { method: "GET" }
  );

  if (!result || !Array.isArray(result.files) || result.files.length === 0) {
    return null;
  }

  return result.files[0];
}

async function savePnsBerprestasiData(records) {
  const sortedData = sortRecordsNewestFirst(records);
  const jsonString = JSON.stringify(sortedData, null, 2);
  const jsonBlob = new Blob([jsonString], { type: "application/json" });

  const existingFile = await findMetadataFile();
  if (existingFile) {
    state.metadataFileId = existingFile.id;
    await updateMultipartFile(existingFile.id, {
      name: CONFIG.METADATA_FILE_NAME,
      mimeType: "application/json",
      dataBlob: jsonBlob
    });
    return;
  }

  const created = await createMultipartFile({
    name: CONFIG.METADATA_FILE_NAME,
    mimeType: "application/json",
    parents: [CONFIG.TARGET_FOLDER_ID],
    dataBlob: jsonBlob
  });

  state.metadataFileId = created.id;
}

async function handleSubmit(event) {
  event.preventDefault();

  if (state.isSubmitting) return;

  try {
    await ensureDriveAccess(true);
    await verifyFolderAccess();
  } catch (error) {
    console.error(error);
    showAlert(error.message || "Google Drive belum siap.", "error");
    return;
  }

  const validation = validateForm();
  if (!validation.isValid) {
    showAlert("Form belum valid. Periksa lagi input yang masih salah.", "error");
    return;
  }

  state.isSubmitting = true;
  dom.submitBtn.disabled = true;

  let uploadedDriveFile = null;

  try {
    const currentData = await loadPnsBerprestasiData();
    state.records = sortRecordsNewestFirst(currentData);

    const data = validation.data;
    const generatedFileName = formatFileName(data.nama, data.jenis_penghargaan, state.selectedFile.name);

    setLoading(true, "Mengupload file ke Google Drive...");
    uploadedDriveFile = await uploadCertificateFile(state.selectedFile, generatedFileName);

    const now = new Date();
    const record = {
      id: `pns-${now.getTime()}`,
      nama: data.nama,
      jenis_penghargaan: data.jenis_penghargaan,
      lingkup_penghargaan: data.lingkup_penghargaan,
      nomor_piagam: data.nomor_piagam,
      tanggal_piagam: data.tanggal_piagam,
      tahun_perolehan: data.tahun_perolehan,
      nama_file: uploadedDriveFile.name,
      file_id: uploadedDriveFile.id,
      file_url: uploadedDriveFile.webViewLink || "",
      download_url: uploadedDriveFile.webContentLink || "",
      mime_type: uploadedDriveFile.mimeType || state.selectedFile.type,
      size: uploadedDriveFile.size || state.selectedFile.size,
      uploaded_at: now.toISOString(),
      created_at_local: formatDateTimeIndonesia(now.toISOString())
    };

    const nextData = sortRecordsNewestFirst([record, ...state.records]);

    setLoading(true, "Menyimpan metadata JSON ke Google Drive...");
    await savePnsBerprestasiData(nextData);

    state.records = nextData;
    updateYearFilterOptions();
    applyFilters();
    resetFormCompletely();

    showAlert("Upload berhasil. File dan metadata tersimpan di Google Drive.", "success");
  } catch (error) {
    console.error(error);

    if (uploadedDriveFile && uploadedDriveFile.id) {
      try {
        setLoading(true, "Rollback file upload...");
        await deleteDriveFile(uploadedDriveFile.id);
      } catch (rollbackError) {
        console.error("Rollback gagal:", rollbackError);
        showAlert(
          "File sempat terupload, tetapi simpan metadata gagal dan rollback juga gagal. Cek folder Drive secara manual.",
          "error",
          12000
        );
        return;
      }
    }

    showAlert(error.message || "Gagal upload ke Google Drive.", "error", 10000);
  } finally {
    state.isSubmitting = false;
    dom.submitBtn.disabled = false;
    setLoading(false);
  }
}

function validateForm() {
  clearAllFieldErrors();

  const data = {
    nama: dom.namaPegawai.value.trim(),
    jenis_penghargaan: dom.jenisPenghargaan.value.trim(),
    lingkup_penghargaan: dom.lingkupPenghargaan.value.trim(),
    nomor_piagam: dom.nomorPiagam.value.trim(),
    tanggal_piagam: dom.tanggalPiagam.value.trim(),
    tahun_perolehan: dom.tahunPerolehan.value.trim()
  };

  let isValid = true;

  if (!data.nama) {
    setFieldError("namaPegawai", "Nama pegawai wajib dipilih.");
    isValid = false;
  }

  if (!data.jenis_penghargaan) {
    setFieldError("jenisPenghargaan", "Jenis penghargaan wajib diisi.");
    isValid = false;
  }

  if (!data.lingkup_penghargaan) {
    setFieldError("lingkupPenghargaan", "Lingkup penghargaan wajib dipilih.");
    isValid = false;
  }

  if (!data.nomor_piagam) {
    setFieldError("nomorPiagam", "Nomor piagam wajib diisi.");
    isValid = false;
  }

  if (!data.tanggal_piagam) {
    setFieldError("tanggalPiagam", "Tanggal wajib diisi.");
    isValid = false;
  }

  if (!data.tahun_perolehan) {
    setFieldError("tahunPerolehan", "Tahun perolehan wajib diisi.");
    isValid = false;
  } else if (!/^\d{4}$/.test(data.tahun_perolehan)) {
    setFieldError("tahunPerolehan", "Tahun perolehan harus 4 digit.");
    isValid = false;
  }

  if (!state.selectedFile) {
    setFieldError("fileSertifikat", "File sertifikat wajib dipilih.");
    isValid = false;
  } else {
    const validation = validateFile(state.selectedFile);
    if (!validation.valid) {
      setFieldError("fileSertifikat", validation.message);
      isValid = false;
    }
  }

  if (isValid) {
    const duplicate = state.records.some((item) => {
      return (
        normalizeText(item.nama) === normalizeText(data.nama) &&
        normalizeText(item.jenis_penghargaan) === normalizeText(data.jenis_penghargaan) &&
        normalizeText(item.nomor_piagam) === normalizeText(data.nomor_piagam) &&
        item.tanggal_piagam === data.tanggal_piagam
      );
    });

    if (duplicate) {
      setFieldError("nomorPiagam", "Data serupa sudah ada. Jangan upload dua kali.");
      isValid = false;
    }
  }

  return { isValid, data };
}

function validateFile(file) {
  if (!file) {
    return { valid: false, message: "File tidak ditemukan." };
  }

  const extension = getFileExtension(file.name);
  const isAllowedExtension = ALLOWED_EXTENSIONS.includes(extension);
  const isAllowedMime = ALLOWED_MIME_TYPES.includes(file.type) || file.type === "";

  if (!isAllowedExtension || !isAllowedMime) {
    return { valid: false, message: "Format file hanya PDF, JPG, JPEG, atau PNG." };
  }

  if (file.size > MAX_FILE_SIZE) {
    return { valid: false, message: `Ukuran file maksimal ${CONFIG.MAX_FILE_SIZE_MB} MB.` };
  }

  return { valid: true, message: "" };
}

function handleFilePreview() {
  clearFieldError("fileSertifikat");

  const file = dom.fileSertifikat.files && dom.fileSertifikat.files[0] ? dom.fileSertifikat.files[0] : null;

  if (!file) {
    clearFileSelection();
    return;
  }

  const validation = validateFile(file);
  if (!validation.valid) {
    clearFileSelection();
    setFieldError("fileSertifikat", validation.message);
    showAlert(validation.message, "error");
    return;
  }

  state.selectedFile = file;
  dom.selectedFileMeta.innerHTML = `
    <strong>${escapeHtml(file.name)}</strong><br>
    Ukuran file: ${formatBytes(file.size)}<br>
    Tipe: ${escapeHtml(file.type || "Tidak terdeteksi")}
  `;

  dom.clearFileBtn.classList.remove("hidden");
  dom.imagePreview.classList.add("hidden");
  dom.pdfPreviewMessage.classList.add("hidden");

  revokePreviewObjectUrl();

  if (file.type === "application/pdf" || getFileExtension(file.name) === "pdf") {
    dom.pdfPreviewMessage.textContent = "File PDF siap diupload.";
    dom.pdfPreviewMessage.classList.remove("hidden");
    return;
  }

  if (file.type.startsWith("image/")) {
    const objectUrl = URL.createObjectURL(file);
    state.previewObjectUrl = objectUrl;
    dom.imagePreview.src = objectUrl;
    dom.imagePreview.classList.remove("hidden");
  }
}

function clearFileSelection() {
  state.selectedFile = null;
  dom.fileSertifikat.value = "";
  dom.selectedFileMeta.textContent = "Belum ada file dipilih.";
  dom.imagePreview.src = "";
  dom.imagePreview.classList.add("hidden");
  dom.pdfPreviewMessage.classList.add("hidden");
  dom.clearFileBtn.classList.add("hidden");
  revokePreviewObjectUrl();
}

function revokePreviewObjectUrl() {
  if (state.previewObjectUrl) {
    URL.revokeObjectURL(state.previewObjectUrl);
    state.previewObjectUrl = null;
  }
}

async function uploadCertificateFile(file, generatedFileName) {
  const response = await createMultipartFile({
    name: generatedFileName,
    mimeType: file.type || inferMimeTypeFromFileName(file.name),
    parents: [CONFIG.TARGET_FOLDER_ID],
    dataBlob: file
  });

  return response;
}

async function createMultipartFile({ name, mimeType, parents = [], dataBlob }) {
  const metadata = { name, parents, mimeType };
  const url = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType,webViewLink,webContentLink,size,createdTime";

  const response = await multipartDriveRequest(url, "POST", metadata, dataBlob, mimeType);
  return response;
}

async function updateMultipartFile(fileId, { name, mimeType, dataBlob }) {
  const metadata = { name, mimeType };
  const url = `https://www.googleapis.com/upload/drive/v3/files/${encodeURIComponent(fileId)}?uploadType=multipart&fields=id,name,mimeType,webViewLink,webContentLink,size,modifiedTime`;

  const response = await multipartDriveRequest(url, "PATCH", metadata, dataBlob, mimeType);
  return response;
}

async function deleteDriveFile(fileId) {
  await driveApiRequest(`/files/${encodeURIComponent(fileId)}?supportsAllDrives=true`, {
    method: "DELETE"
  });
}

async function handleDelete(id) {
  const record = state.records.find((item) => item.id === id);
  if (!record) {
    showAlert("Data yang akan dihapus tidak ditemukan.", "error");
    return;
  }

  const confirmed = window.confirm(
    `Hapus data untuk ${record.nama}?\n\nFile sertifikat dan record metadata akan dihapus dari Google Drive.`
  );

  if (!confirmed) return;

  try {
    await ensureDriveAccess(true);
    setLoading(true, "Memuat metadata terbaru...");

    const latestData = await loadPnsBerprestasiData();
    const targetRecord = latestData.find((item) => item.id === id);

    if (!targetRecord) {
      throw new Error("Record tidak ditemukan di metadata terbaru.");
    }

    if (targetRecord.file_id) {
      setLoading(true, "Menghapus file sertifikat dari Google Drive...");
      await deleteDriveFile(targetRecord.file_id);
    }

    const nextData = latestData.filter((item) => item.id !== id);

    setLoading(true, "Mengupdate metadata JSON...");
    await savePnsBerprestasiData(nextData);

    state.records = sortRecordsNewestFirst(nextData);
    updateYearFilterOptions();
    applyFilters();

    showAlert("Data dan file berhasil dihapus.", "success");
  } catch (error) {
    console.error(error);
    showAlert(error.message || "Gagal menghapus data.", "error", 10000);
  } finally {
    setLoading(false);
  }
}

function applyFilters() {
  const searchTerm = normalizeText(dom.searchInput.value);
  const selectedYear = dom.yearFilter.value.trim();
  const selectedLingkup = dom.lingkupFilter.value.trim();

  state.filteredRecords = state.records.filter((record) => {
    const matchSearch = !searchTerm || normalizeText(record.nama).includes(searchTerm);
    const matchYear = !selectedYear || String(record.tahun_perolehan) === selectedYear;
    const matchLingkup = !selectedLingkup || String(record.lingkup_penghargaan) === selectedLingkup;
    return matchSearch && matchYear && matchLingkup;
  });

  renderTable(state.filteredRecords);
}

function renderTable(records = state.filteredRecords) {
  dom.dataTableBody.innerHTML = "";

  if (!records.length) {
    dom.emptyState.style.display = "block";
    return;
  }

  dom.emptyState.style.display = "block";
  dom.emptyState.textContent = "";

  const rows = records.map((record, index) => {
    return `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(record.nama || "-")}</td>
        <td>${escapeHtml(record.jenis_penghargaan || "-")}</td>
        <td><span class="badge">${escapeHtml(record.lingkup_penghargaan || "-")}</span></td>
        <td class="mono">${escapeHtml(record.nomor_piagam || "-")}</td>
        <td>${escapeHtml(formatDateIndonesia(record.tanggal_piagam))}</td>
        <td>${escapeHtml(String(record.tahun_perolehan || "-"))}</td>
        <td class="file-name">${escapeHtml(record.nama_file || "-")}</td>
        <td>${escapeHtml(formatDateTimeIndonesia(record.uploaded_at || record.created_at_local || ""))}</td>
        <td>
          <div class="table-actions">
            <button type="button" class="btn btn-ghost btn-sm" data-action="view" data-id="${escapeHtml(record.id)}">Lihat</button>
            <button type="button" class="btn btn-danger btn-sm" data-action="delete" data-id="${escapeHtml(record.id)}">Hapus</button>
          </div>
        </td>
      </tr>
    `;
  });

  dom.dataTableBody.innerHTML = rows.join("");
}

function updateYearFilterOptions() {
  const currentValue = dom.yearFilter.value;
  const years = [...new Set(state.records.map((item) => String(item.tahun_perolehan)).filter(Boolean))]
    .sort((a, b) => Number(b) - Number(a));

  dom.yearFilter.innerHTML = `<option value="">Semua Tahun</option>`;
  years.forEach((year) => {
    const option = document.createElement("option");
    option.value = year;
    option.textContent = year;
    dom.yearFilter.appendChild(option);
  });

  if (years.includes(currentValue)) {
    dom.yearFilter.value = currentValue;
  }
}

function resetFilters() {
  dom.searchInput.value = "";
  dom.yearFilter.value = "";
  dom.lingkupFilter.value = "";
  applyFilters();
}

function resetFormCompletely() {
  dom.form.reset();
  clearAllFieldErrors();
  clearFileSelection();
}

function showAlert(message, type = "info", timeout = 5000) {
  const alert = document.createElement("div");
  alert.className = `alert alert-${type}`;
  alert.innerHTML = escapeHtml(message);

  dom.alertContainer.prepend(alert);

  window.setTimeout(() => {
    alert.remove();
  }, timeout);
}

function setLoading(isLoading, message = "Memproses...") {
  dom.loadingText.textContent = message;
  dom.loadingOverlay.classList.toggle("hidden", !isLoading);
}

function setFieldError(fieldId, message) {
  const errorEl = document.getElementById(`error-${fieldId}`);
  if (errorEl) {
    errorEl.textContent = message;
  }
}

function clearFieldError(fieldId) {
  const errorEl = document.getElementById(`error-${fieldId}`);
  if (errorEl) {
    errorEl.textContent = "";
  }
}

function clearAllFieldErrors() {
  [
    "namaPegawai",
    "jenisPenghargaan",
    "lingkupPenghargaan",
    "nomorPiagam",
    "tanggalPiagam",
    "tahunPerolehan",
    "fileSertifikat"
  ].forEach(clearFieldError);
}

async function driveApiRequest(pathOrUrl, options = {}) {
  const {
    method = "GET",
    body = null,
    headers = {},
    rawText = false
  } = options;

  if (!state.accessToken) {
    throw new Error("Token akses Google Drive belum tersedia.");
  }

  const url = pathOrUrl.startsWith("http")
    ? pathOrUrl
    : `https://www.googleapis.com/drive/v3${pathOrUrl}`;

  const requestOptions = {
    method,
    headers: {
      Authorization: `Bearer ${state.accessToken}`,
      ...headers
    }
  };

  if (body !== null) {
    requestOptions.body = body;
  }

  const response = await fetch(url, requestOptions);

  if (method === "DELETE" && response.ok) {
    return true;
  }

  if (!response.ok) {
    let errorPayload = null;
    try {
      errorPayload = await response.json();
    } catch {
      try {
        errorPayload = await response.text();
      } catch {
        errorPayload = null;
      }
    }
    throw new Error(buildDriveErrorMessage(response.status, errorPayload));
  }

  if (rawText) {
    return await response.text();
  }

  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    return await response.json();
  }

  return true;
}

async function multipartDriveRequest(url, method, metadata, dataBlob, mimeType) {
  const boundary = "-------314159265358979323846";
  const delimiter = `\r\n--${boundary}\r\n`;
  const closeDelimiter = `\r\n--${boundary}--`;

  const metadataPart =
    delimiter +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    JSON.stringify(metadata);

  const fileHeader = delimiter + `Content-Type: ${mimeType || "application/octet-stream"}\r\n\r\n`;

  const body = new Blob(
    [
      metadataPart,
      fileHeader,
      dataBlob,
      closeDelimiter
    ],
    { type: `multipart/related; boundary=${boundary}` }
  );

  return await driveApiRequest(url, {
    method,
    body,
    headers: {
      "Content-Type": `multipart/related; boundary=${boundary}`
    }
  });
}

function buildDriveErrorMessage(status, payload) {
  const errorObj = payload && payload.error ? payload.error : null;
  const message =
    (errorObj && errorObj.message) ||
    (typeof payload === "string" ? payload : "Terjadi kesalahan Google Drive API.");

  if (status === 400) {
    return `Request Google Drive tidak valid: ${message}`;
  }

  if (status === 401) {
    return "Akses Google Drive tidak valid atau sesi login sudah kadaluarsa.";
  }

  if (status === 403) {
    return `Akses Google Drive ditolak. Biasanya akun tidak punya izin ke folder atau scope OAuth tidak cukup. Detail: ${message}`;
  }

  if (status === 404) {
    return "Folder atau file Google Drive tidak ditemukan.";
  }

  if (status === 429) {
    return "Terlalu banyak request ke Google Drive. Coba lagi sebentar lagi.";
  }

  return `Google Drive API error ${status}: ${message}`;
}

function formatFileName(employeeName, awardType, originalName) {
  const extension = getFileExtension(originalName);
  const cleanedEmployee = slugify(employeeName);
  const cleanedAward = slugify(awardType);
  const timestamp = getCompactTimestamp();
  return `${cleanedEmployee}_${cleanedAward}_${timestamp}.${extension}`;
}

function slugify(text) {
  return text
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .replace(/_+/g, "_")
    .toLowerCase();
}

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function getFileExtension(fileName) {
  const parts = String(fileName || "").split(".");
  return parts.length > 1 ? parts.pop().toLowerCase() : "";
}

function inferMimeTypeFromFileName(fileName) {
  const ext = getFileExtension(fileName);
  if (ext === "pdf") return "application/pdf";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "png") return "image/png";
  return "application/octet-stream";
}

function formatBytes(bytes) {
  if (!bytes && bytes !== 0) return "-";
  const units = ["B", "KB", "MB", "GB"];
  let size = Number(bytes);
  let unitIndex = 0;

  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex += 1;
  }

  return `${size.toFixed(size >= 10 || unitIndex === 0 ? 0 : 2)} ${units[unitIndex]}`;
}

function formatDateIndonesia(dateString) {
  if (!dateString) return "-";
  const date = new Date(dateString);
  if (Number.isNaN(date.getTime())) return String(dateString);

  return date.toLocaleDateString("id-ID", {
    day: "2-digit",
    month: "long",
    year: "numeric"
  });
}

function formatDateTimeIndonesia(dateString) {
  if (!dateString) return "-";
  const date = new Date(dateString);
  if (Number.isNaN(date.getTime())) return String(dateString);

  return date.toLocaleString("id-ID", {
    day: "2-digit",
    month: "long",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function getCompactTimestamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return [
    now.getFullYear(),
    pad(now.getMonth() + 1),
    pad(now.getDate()),
    "_",
    pad(now.getHours()),
    pad(now.getMinutes()),
    pad(now.getSeconds())
  ].join("");
}

function sortRecordsNewestFirst(records) {
  return [...records].sort((a, b) => {
    const dateA = new Date(a.uploaded_at || 0).getTime();
    const dateB = new Date(b.uploaded_at || 0).getTime();
    return dateB - dateA;
  });
}

function escapeDriveQueryString(value) {
  return String(value || "")
    .replace(/\\/g, "\\\\")
    .replace(/'/g, "\\'");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
