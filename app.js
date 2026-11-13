(function () {
  "use strict";

  const C = window.CrewConverter;
  const state = {
    sourceFile: null,
    sheetName: "",
    records: [],
    movements: [],
    search: "",
    issuesOnly: false,
    nextMovementId: 1,
  };

  const elements = {
    hero: document.querySelector("#hero"),
    workspace: document.querySelector("#workspace"),
    fileInput: document.querySelector("#file-input"),
    dropzone: document.querySelector("#dropzone"),
    changeFile: document.querySelector("#change-file"),
    sourceName: document.querySelector("#source-name"),
    sourceMeta: document.querySelector("#source-meta"),
    crewCount: document.querySelector("#crew-count"),
    validCount: document.querySelector("#valid-count"),
    warningCount: document.querySelector("#warning-count"),
    crewBody: document.querySelector("#crew-body"),
    crewSearch: document.querySelector("#crew-search"),
    showIssues: document.querySelector("#show-issues"),
    tableResultCount: document.querySelector("#table-result-count"),
    emptyTable: document.querySelector("#empty-table"),
    movementList: document.querySelector("#movement-list"),
    addMovement: document.querySelector("#add-movement"),
    exportAll: document.querySelector("#export-all"),
    exportTitle: document.querySelector("#export-title"),
    exportDescription: document.querySelector("#export-description"),
    progressCrew: document.querySelector("#progress-crew"),
    progressMovements: document.querySelector("#progress-movements"),
    progressExport: document.querySelector("#progress-export"),
    crewDialog: document.querySelector("#crew-dialog"),
    crewForm: document.querySelector("#crew-form"),
    modalTitle: document.querySelector("#modal-title"),
    helpButton: document.querySelector("#help-button"),
    helpDialog: document.querySelector("#help-dialog"),
    toast: document.querySelector("#toast"),
  };

  let toastTimer;

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function formatBytes(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function showToast(message, type) {
    clearTimeout(toastTimer);
    elements.toast.textContent = message;
    elements.toast.classList.toggle("is-error", type === "error");
    elements.toast.classList.add("is-visible");
    toastTimer = setTimeout(() => elements.toast.classList.remove("is-visible"), 3600);
  }

  function crewValidation() {
    return state.records.map((record) => C.validateRecord(record));
  }

  function invalidCrewCount() {
    return crewValidation().filter((result) => !result.valid).length;
  }

  function movementById(id) {
    return state.movements.find((movement) => String(movement.id) === String(id));
  }

  function uniquePorts() {
    const seen = new Set();
    return Object.values(C.PORTS).filter((port) => {
      if (seen.has(port.code)) return false;
      seen.add(port.code);
      return true;
    });
  }

  function ensurePortList() {
    if (document.querySelector("#port-codes")) return;
    const datalist = document.createElement("datalist");
    datalist.id = "port-codes";
    datalist.innerHTML = uniquePorts()
      .map((port) => `<option value="${escapeHtml(port.code)}">${escapeHtml(port.label)}</option>`)
      .join("");
    document.body.appendChild(datalist);
  }

  async function importFile(file) {
    if (!file) return;
    if (!/\.(xls|xlsx|csv)$/i.test(file.name)) {
      showToast("Selecciona un archivo .xls, .xlsx o .csv.", "error");
      return;
    }
    if (file.size > 10 * 1024 * 1024) {
      showToast("El archivo supera el límite de 10 MB.", "error");
      return;
    }
    if (!window.XLSX) {
      showToast("No se ha podido cargar el lector de Excel.", "error");
      return;
    }

    elements.dropzone.classList.add("is-loading");
    try {
      const buffer = await file.arrayBuffer();
      const parsed = C.parseWorkbook(buffer, window.XLSX);
      state.sourceFile = file;
      state.sheetName = parsed.sheetName;
      state.records = parsed.records;
      state.search = "";
      state.issuesOnly = false;
      state.nextMovementId = 1;
      const inferred = C.inferMovement(file.name, parsed.records);
      state.movements = [{ id: state.nextMovementId++, ...inferred }];

      elements.crewSearch.value = "";
      elements.showIssues.textContent = "Mostrar incidencias";
      elements.hero.classList.add("is-hidden");
      elements.workspace.classList.remove("is-hidden");
      renderAll();
      window.scrollTo({ top: 0, behavior: "smooth" });
      showToast(`${parsed.records.length} tripulantes importados correctamente.`);
    } catch (error) {
      console.error(error);
      showToast(error && error.message ? error.message : "No se ha podido leer el archivo.", "error");
    } finally {
      elements.dropzone.classList.remove("is-loading");
      elements.fileInput.value = "";
    }
  }

  function renderSummary() {
    const validations = crewValidation();
    const invalid = validations.filter((result) => !result.valid).length;
    elements.sourceName.textContent = state.sourceFile ? state.sourceFile.name : "—";
    elements.sourceMeta.textContent = state.sourceFile
      ? `${formatBytes(state.sourceFile.size)} · Hoja: ${state.sheetName}`
      : "—";
    elements.crewCount.textContent = String(state.records.length);
    elements.validCount.textContent = String(state.records.length - invalid);
    elements.warningCount.textContent = String(invalid);
  }

  function issueText(validation) {
    const parts = [];
    if (validation.missing.length) parts.push(`Faltan: ${validation.missing.join(", ")}`);
    if (validation.warnings.length) parts.push(validation.warnings.join(" · "));
    return parts.join(" · ");
  }

  function renderCrew() {
    const query = state.search.trim().toLocaleLowerCase("es");
    const validations = crewValidation();
    const visible = state.records
      .map((record, index) => ({ record, index, validation: validations[index] }))
      .filter(({ record, validation }) => {
        if (state.issuesOnly && validation.valid) return false;
        if (!query) return true;
        return [record.Nombre, record.Apellidos, record.NumeroDocumento, record.Nacionalidad, record.CategoriaDes]
          .join(" ")
          .toLocaleLowerCase("es")
          .includes(query);
      });

    elements.crewBody.innerHTML = visible
      .map(({ record, index, validation }) => {
        const issue = issueText(validation);
        return `
          <tr class="${validation.valid ? "" : "has-issue"}">
            <td class="person-cell"><strong>${escapeHtml(`${record.Nombre} ${record.Apellidos}`.trim())}</strong><span>Registro ${index + 1}</span></td>
            <td class="document-cell"><strong>${escapeHtml(record.NumeroDocumento || "—")}</strong><span>${escapeHtml(record.TipoDocumento || "Sin tipo")} · Cad. ${escapeHtml(record.FechaCaducidadDocumento || "—")}</span></td>
            <td>${escapeHtml(record.Nacionalidad || "—")}</td>
            <td>${escapeHtml(record.FechaNacimiento || "—")}</td>
            <td class="category-cell"><strong>${escapeHtml(record.Categoria || "—")}</strong><span>${escapeHtml(record.CategoriaDes || "Sin descripción")}</span></td>
            <td>${escapeHtml(record.Sexo || "—")}</td>
            <td><span class="status-badge ${validation.valid ? "" : "issue"}" title="${escapeHtml(issue)}">${validation.valid ? "Correcto" : "Revisar"}</span></td>
            <td><button class="edit-button" type="button" data-edit-record="${index}">Editar</button></td>
          </tr>`;
      })
      .join("");

    elements.emptyTable.classList.toggle("is-hidden", visible.length !== 0);
    elements.tableResultCount.textContent = `${visible.length} ${visible.length === 1 ? "registro" : "registros"}`;
  }

  function directionName(direction) {
    return direction === "E" ? "Entrada" : "Salida";
  }

  function routeName(movement) {
    const origin = C.shortPort(movement.embarkPort) || "ORIGEN";
    const destination = C.shortPort(movement.disembarkPort) || "DESTINO";
    return `${directionName(movement.direction)} · ${origin} → ${destination}`;
  }

  function movementTemplate(movement, index) {
    const errors = C.movementErrors(movement);
    const invalid = (field) => {
      const checks = {
        reference: !String(movement.reference || "").trim(),
        scheduleDate: !movement.scheduleDate,
        scheduleTime: !movement.scheduleTime,
        embarkPort: !C.normalizePort(movement.embarkPort),
        disembarkPort: !C.normalizePort(movement.disembarkPort),
        embarkDate: !movement.embarkDate,
        embarkTime: !movement.embarkTime,
        disembarkDate: !movement.disembarkDate,
        disembarkTime: !movement.disembarkTime,
      };
      return checks[field] ? 'aria-invalid="true"' : "";
    };
    const errorCopy = errors.length ? `${errors.length} ${errors.length === 1 ? "dato pendiente" : "datos pendientes"}: ${errors.join(", ")}` : "Listo para exportar";

    return `
      <article class="movement-card ${errors.length ? "has-errors" : ""}" data-movement-id="${movement.id}">
        <div class="movement-head">
          <div class="movement-ident">
            <span class="movement-index">${String(index + 1).padStart(2, "0")}</span>
            <div><strong>Movimiento ${index + 1}</strong><span data-route-title>${escapeHtml(routeName(movement))}</span></div>
          </div>
          <div class="movement-tools">
            <button class="tool-button" type="button" data-action="duplicate" title="Duplicar movimiento"><svg aria-hidden="true" viewBox="0 0 24 24"><path d="M8 8V3h13v13h-5v5H3V8h5Zm2 0h6v6h3V5h-9v3Zm-5 2v9h9v-9H5Z" /></svg></button>
            <button class="tool-button" type="button" data-action="download" title="Descargar este CSV"><svg aria-hidden="true" viewBox="0 0 24 24"><path d="M13 4h-2v8.2L8.4 9.6 7 11l5 5 5-5-1.4-1.4-2.6 2.6V4ZM5 18v2h14v-2H5Z" /></svg></button>
            <button class="tool-button danger" type="button" data-action="remove" title="Eliminar movimiento"><svg aria-hidden="true" viewBox="0 0 24 24"><path d="M8 3h8l1 2h4v2H3V5h4l1-2Zm-2 6h12l-1 12H7L6 9Zm3 2v7h2v-7H9Zm4 0v7h2v-7h-2Z" /></svg></button>
          </div>
        </div>
        <div class="movement-body">
          <div class="form-grid">
            <label class="field"><span>Sentido <b>*</b></span><select data-field="direction"><option value="S" ${movement.direction === "S" ? "selected" : ""}>S · Salida</option><option value="E" ${movement.direction === "E" ? "selected" : ""}>E · Entrada</option></select></label>
            <label class="field"><span>Número de escala <b>*</b></span><input data-field="reference" value="${escapeHtml(movement.reference)}" placeholder="Ej. 29941" inputmode="numeric" ${invalid("reference")} /><small class="field-help">Se incluirá en el nombre del archivo.</small></label>
            <label class="field"><span>Fecha oficial del viaje <b>*</b></span><input type="date" data-field="scheduleDate" value="${escapeHtml(movement.scheduleDate)}" ${invalid("scheduleDate")} /><small class="field-help">Fecha usada en el nombre.</small></label>
            <label class="field"><span>Hora oficial del viaje <b>*</b></span><input type="time" data-field="scheduleTime" value="${escapeHtml(movement.scheduleTime)}" ${invalid("scheduleTime")} /><small class="field-help">Hora usada en el nombre.</small></label>

            <div class="route-fields">
              <label class="field"><span>Puerto de embarque <b>*</b></span><input data-field="embarkPort" value="${escapeHtml(movement.embarkPort)}" list="port-codes" maxlength="5" placeholder="ESALG" ${invalid("embarkPort")} /></label>
              <button class="swap-button" type="button" data-action="swap" title="Intercambiar puertos"><svg aria-hidden="true" viewBox="0 0 24 24"><path d="m7 7 3-3H4v6l3-3Zm10 10-3 3h6v-6l-3 3ZM5.4 15.7l1.4-1.4A7.1 7.1 0 0 0 13 18v2a9.1 9.1 0 0 1-7.6-4.3ZM11 4a9.1 9.1 0 0 1 7.6 4.3l-1.4 1.4A7.1 7.1 0 0 0 11 6V4Z" /></svg></button>
              <label class="field"><span>Puerto de desembarque <b>*</b></span><input data-field="disembarkPort" value="${escapeHtml(movement.disembarkPort)}" list="port-codes" maxlength="5" placeholder="ESCEU" ${invalid("disembarkPort")} /></label>
            </div>
            <label class="field span-2"><span>Naviera <small>(opcional)</small></span><input data-field="shippingCompany" value="${escapeHtml(movement.shippingCompany)}" placeholder="Se conserva la del Excel si queda vacío" /></label>
            <label class="field"><span>Fecha embarque <b>*</b></span><input type="date" data-field="embarkDate" value="${escapeHtml(movement.embarkDate)}" ${invalid("embarkDate")} /></label>
            <label class="field"><span>Hora embarque <b>*</b></span><input type="time" data-field="embarkTime" value="${escapeHtml(movement.embarkTime)}" ${invalid("embarkTime")} /></label>
            <label class="field"><span>Fecha desembarque <b>*</b></span><input type="date" data-field="disembarkDate" value="${escapeHtml(movement.disembarkDate)}" ${invalid("disembarkDate")} /></label>
            <label class="field"><span>Hora desembarque <b>*</b></span><input type="time" data-field="disembarkTime" value="${escapeHtml(movement.disembarkTime)}" ${invalid("disembarkTime")} /></label>
          </div>
          <div class="movement-foot">
            <div class="filename-preview">
              <svg aria-hidden="true" viewBox="0 0 24 24"><path d="M6 2h8l4 4v16H6V2Zm8 1.8V7h3.2L14 3.8ZM8 11v1.7h8V11H8Zm0 4v1.7h8V15H8Z" /></svg>
              <div><small>NOMBRE DE SALIDA</small><code data-filename>${escapeHtml(C.buildFilename(movement))}</code></div>
            </div>
            <span class="movement-error" data-movement-error>${escapeHtml(errorCopy)}</span>
          </div>
        </div>
      </article>`;
  }

  function renderMovements() {
    ensurePortList();
    elements.movementList.innerHTML = state.movements.map(movementTemplate).join("");
  }

  function refreshMovementCard(movement) {
    const card = elements.movementList.querySelector(`[data-movement-id="${movement.id}"]`);
    if (!card) return;
    const errors = C.movementErrors(movement);
    const errorCopy = errors.length ? `${errors.length} ${errors.length === 1 ? "dato pendiente" : "datos pendientes"}: ${errors.join(", ")}` : "Listo para exportar";
    card.classList.toggle("has-errors", errors.length > 0);
    card.querySelector("[data-route-title]").textContent = routeName(movement);
    card.querySelector("[data-filename]").textContent = C.buildFilename(movement);
    card.querySelector("[data-movement-error]").textContent = errorCopy;

    const fieldChecks = {
      reference: !String(movement.reference || "").trim(),
      scheduleDate: !movement.scheduleDate,
      scheduleTime: !movement.scheduleTime,
      embarkPort: !C.normalizePort(movement.embarkPort),
      disembarkPort: !C.normalizePort(movement.disembarkPort),
      embarkDate: !movement.embarkDate,
      embarkTime: !movement.embarkTime,
      disembarkDate: !movement.disembarkDate,
      disembarkTime: !movement.disembarkTime,
    };
    Object.entries(fieldChecks).forEach(([field, invalid]) => {
      const input = card.querySelector(`[data-field="${field}"]`);
      if (input) input.setAttribute("aria-invalid", invalid ? "true" : "false");
    });
  }

  function setProgressState(element, stateName) {
    element.classList.remove("is-active", "is-complete");
    if (stateName) element.classList.add(`is-${stateName}`);
  }

  function updateExportState() {
    const crewIssues = invalidCrewCount();
    const validMovements = state.movements.filter((movement) => C.movementErrors(movement).length === 0).length;
    const allMovementsValid = validMovements === state.movements.length && state.movements.length > 0;
    const ready = crewIssues === 0 && allMovementsValid;

    elements.exportAll.disabled = !ready;
    const exportLabel = elements.exportAll.querySelector("span");
    exportLabel.textContent = state.movements.length > 1 ? `Descargar ${state.movements.length} archivos` : "Descargar CSV";

    if (crewIssues > 0) {
      elements.exportTitle.textContent = `Revisa ${crewIssues} ${crewIssues === 1 ? "tripulante" : "tripulantes"}`;
      elements.exportDescription.textContent = "Corrige los registros marcados antes de generar los archivos.";
    } else if (!allMovementsValid) {
      const pending = state.movements.length - validMovements;
      elements.exportTitle.textContent = `Completa ${pending} ${pending === 1 ? "movimiento" : "movimientos"}`;
      elements.exportDescription.textContent = "Faltan datos de escala, ruta o fecha en las tarjetas señaladas.";
    } else {
      elements.exportTitle.textContent = state.movements.length > 1 ? `${state.movements.length} CSV listos para descargar` : "CSV listo para descargar";
      elements.exportDescription.textContent = `${state.records.length} tripulantes · 29 columnas · Separador “;” · Formato CRLF`;
    }

    setProgressState(elements.progressCrew, crewIssues === 0 ? "complete" : "active");
    setProgressState(elements.progressMovements, crewIssues > 0 ? null : allMovementsValid ? "complete" : "active");
    setProgressState(elements.progressExport, ready ? "active" : null);
  }

  function renderAll() {
    renderSummary();
    renderCrew();
    renderMovements();
    updateExportState();
  }

  function openCrewEditor(index) {
    const record = state.records[index];
    if (!record) return;
    elements.crewForm.reset();
    elements.crewForm.elements.recordIndex.value = String(index);
    [
      "Nombre", "Apellidos", "TipoDocumento", "NumeroDocumento", "NumeroVisado",
      "FechaCaducidadDocumento", "Nacionalidad", "FechaNacimiento", "LugarNacimiento",
      "Categoria", "CategoriaDes", "Sexo", "Observaciones",
    ].forEach((field) => {
      elements.crewForm.elements[field].value = record[field] || "";
    });
    elements.modalTitle.textContent = `${record.Nombre} ${record.Apellidos}`.trim() || "Datos del tripulante";
    elements.crewDialog.showModal();
  }

  function saveCrewRecord() {
    if (!elements.crewForm.reportValidity()) return;
    const index = Number(elements.crewForm.elements.recordIndex.value);
    const previous = state.records[index];
    if (!previous) return;
    const updated = { ...previous };
    [
      "Nombre", "Apellidos", "TipoDocumento", "NumeroDocumento", "NumeroVisado",
      "FechaCaducidadDocumento", "Nacionalidad", "FechaNacimiento", "LugarNacimiento",
      "Categoria", "CategoriaDes", "Sexo", "Observaciones",
    ].forEach((field) => {
      updated[field] = elements.crewForm.elements[field].value;
    });
    state.records[index] = C.normalizeRecord(updated);
    elements.crewDialog.close();
    renderSummary();
    renderCrew();
    updateExportState();
    showToast("Registro actualizado.");
  }

  function addMovement(source) {
    const base = source || state.movements[state.movements.length - 1] || {
      direction: "S",
      reference: "",
      scheduleDate: "",
      scheduleTime: "",
      embarkPort: "ESALG",
      disembarkPort: "",
      embarkDate: "",
      embarkTime: "",
      disembarkDate: "",
      disembarkTime: "",
      shippingCompany: "",
    };
    state.movements.push({ ...base, id: state.nextMovementId++, reference: "" });
    renderMovements();
    updateExportState();
    const last = elements.movementList.lastElementChild;
    if (last) last.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function csvFile(movement) {
    return {
      name: C.buildFilename(movement),
      content: C.buildCsv(state.records, movement),
    };
  }

  function canDownloadMovement(movement) {
    if (invalidCrewCount() > 0) {
      showToast("Corrige primero los registros de tripulación marcados.", "error");
      return false;
    }
    const errors = C.movementErrors(movement);
    if (errors.length) {
      showToast(`Completa este movimiento: ${errors.join(", ")}.`, "error");
      return false;
    }
    return true;
  }

  function downloadMovement(movement) {
    if (!canDownloadMovement(movement)) return;
    const file = csvFile(movement);
    downloadBlob(new Blob([file.content], { type: "text/csv;charset=utf-8" }), file.name);
    showToast("CSV generado con el formato de Autoridad Portuaria.");
  }

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i += 1) {
      crc ^= bytes[i];
      for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function concatBytes(parts) {
    const length = parts.reduce((total, part) => total + part.length, 0);
    const output = new Uint8Array(length);
    let offset = 0;
    parts.forEach((part) => {
      output.set(part, offset);
      offset += part.length;
    });
    return output;
  }

  function zipDateTime(date) {
    const year = Math.max(1980, date.getFullYear());
    return {
      time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
      date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
    };
  }

  function createZip(files) {
    const encoder = new TextEncoder();
    const localParts = [];
    const centralParts = [];
    const now = zipDateTime(new Date());
    let offset = 0;

    files.forEach((file) => {
      const name = encoder.encode(file.name);
      const data = encoder.encode(file.content);
      const crc = crc32(data);
      const local = new Uint8Array(30);
      const lv = new DataView(local.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(6, 0x0800, true);
      lv.setUint16(8, 0, true);
      lv.setUint16(10, now.time, true);
      lv.setUint16(12, now.date, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, data.length, true);
      lv.setUint32(22, data.length, true);
      lv.setUint16(26, name.length, true);
      lv.setUint16(28, 0, true);
      localParts.push(local, name, data);

      const central = new Uint8Array(46);
      const cv = new DataView(central.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, now.time, true);
      cv.setUint16(14, now.date, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true);
      cv.setUint32(24, data.length, true);
      cv.setUint16(28, name.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      centralParts.push(central, name);
      offset += local.length + name.length + data.length;
    });

    const localBytes = concatBytes(localParts);
    const centralBytes = concatBytes(centralParts);
    const end = new Uint8Array(22);
    const ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, centralBytes.length, true);
    ev.setUint32(16, localBytes.length, true);
    ev.setUint16(20, 0, true);
    return concatBytes([localBytes, centralBytes, end]);
  }

  function uniqueOutputFiles() {
    const used = new Map();
    return state.movements.map((movement) => {
      const file = csvFile(movement);
      const count = (used.get(file.name) || 0) + 1;
      used.set(file.name, count);
      if (count > 1) file.name = file.name.replace(/\.csv$/i, `(${count}).csv`);
      return file;
    });
  }

  function downloadAll() {
    if (invalidCrewCount() > 0 || state.movements.some((movement) => C.movementErrors(movement).length)) {
      showToast("Aún hay datos pendientes de revisión.", "error");
      return;
    }
    const files = uniqueOutputFiles();
    if (files.length === 1) {
      downloadBlob(new Blob([files[0].content], { type: "text/csv;charset=utf-8" }), files[0].name);
    } else {
      const zip = createZip(files);
      const firstDate = state.movements[0].embarkDate || "movimientos";
      downloadBlob(new Blob([zip], { type: "application/zip" }), `Crewlists-${firstDate}.zip`);
    }
    showToast(files.length === 1 ? "CSV generado correctamente." : `${files.length} CSV incluidos en el paquete ZIP.`);
  }

  elements.dropzone.addEventListener("click", () => elements.fileInput.click());
  elements.changeFile.addEventListener("click", () => elements.fileInput.click());
  elements.fileInput.addEventListener("change", (event) => importFile(event.target.files[0]));

  ["dragenter", "dragover"].forEach((eventName) => {
    elements.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropzone.classList.add("is-dragging");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    elements.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropzone.classList.remove("is-dragging");
    });
  });
  elements.dropzone.addEventListener("drop", (event) => importFile(event.dataTransfer.files[0]));

  elements.crewSearch.addEventListener("input", (event) => {
    state.search = event.target.value;
    renderCrew();
  });
  elements.showIssues.addEventListener("click", () => {
    state.issuesOnly = !state.issuesOnly;
    elements.showIssues.textContent = state.issuesOnly ? "Mostrar todos" : "Mostrar incidencias";
    renderCrew();
  });
  elements.crewBody.addEventListener("click", (event) => {
    const button = event.target.closest("[data-edit-record]");
    if (button) openCrewEditor(Number(button.dataset.editRecord));
  });

  elements.crewForm.addEventListener("submit", (event) => {
    event.preventDefault();
    if (event.submitter && event.submitter.value === "cancel") {
      elements.crewDialog.close();
      return;
    }
    saveCrewRecord();
  });

  elements.addMovement.addEventListener("click", () => addMovement());
  elements.movementList.addEventListener("input", (event) => {
    const input = event.target.closest("[data-field]");
    const card = event.target.closest("[data-movement-id]");
    if (!input || !card) return;
    const movement = movementById(card.dataset.movementId);
    if (!movement) return;
    let value = input.value;
    if (["embarkPort", "disembarkPort"].includes(input.dataset.field)) value = value.toUpperCase().replace(/\s/g, "");
    movement[input.dataset.field] = value;
    refreshMovementCard(movement);
    updateExportState();
  });
  elements.movementList.addEventListener("change", (event) => {
    const input = event.target.closest("[data-field]");
    if (input && ["embarkPort", "disembarkPort"].includes(input.dataset.field)) {
      input.value = C.normalizePort(input.value);
      input.dispatchEvent(new Event("input", { bubbles: true }));
    }
  });
  elements.movementList.addEventListener("click", (event) => {
    const button = event.target.closest("[data-action]");
    const card = event.target.closest("[data-movement-id]");
    if (!button || !card) return;
    const movement = movementById(card.dataset.movementId);
    if (!movement) return;

    if (button.dataset.action === "swap") {
      [movement.embarkPort, movement.disembarkPort] = [movement.disembarkPort, movement.embarkPort];
      movement.direction = movement.disembarkPort === "ESALG" ? "E" : movement.embarkPort === "ESALG" ? "S" : movement.direction;
      renderMovements();
      updateExportState();
    } else if (button.dataset.action === "duplicate") {
      addMovement(movement);
    } else if (button.dataset.action === "remove") {
      if (state.movements.length === 1) {
        showToast("Debe existir al menos un movimiento.", "error");
        return;
      }
      state.movements = state.movements.filter((item) => item.id !== movement.id);
      renderMovements();
      updateExportState();
    } else if (button.dataset.action === "download") {
      downloadMovement(movement);
    }
  });

  elements.exportAll.addEventListener("click", downloadAll);
  elements.helpButton.addEventListener("click", () => elements.helpDialog.showModal());

  window.addEventListener("dragover", (event) => event.preventDefault());
  window.addEventListener("drop", (event) => {
    if (!event.target.closest || !event.target.closest("#dropzone")) event.preventDefault();
  });
})();
