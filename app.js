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

  const CREW_FIELDS = [
    "Nombre", "Apellidos", "TipoDocumento", "NumeroDocumento", "NumeroVisado",
    "FechaCaducidadDocumento", "Nacionalidad", "FechaNacimiento", "LugarNacimiento",
    "Categoria", "CategoriaDes", "Sexo", "EntSal", "Observaciones",
  ];

  const elements = {
    hero: document.querySelector("#hero"),
    workspace: document.querySelector("#workspace"),
    fileInput: document.querySelector("#file-input"),
    dropzone: document.querySelector("#dropzone"),
    dropLabel: document.querySelector("#dropzone-label"),
    addCrew: document.querySelector("#add-crew"),
    saveCrew: document.querySelector("#save-crew"),
    crewBody: document.querySelector("#crew-body"),
    crewSearch: document.querySelector("#crew-search"),
    showIssues: document.querySelector("#show-issues"),
    tableResultCount: document.querySelector("#table-result-count"),
    emptyTable: document.querySelector("#empty-table"),
    movementList: document.querySelector("#movement-list"),
    exportAll: document.querySelector("#export-all"),
    crewDialog: document.querySelector("#crew-dialog"),
    crewForm: document.querySelector("#crew-form"),
    modalTitle: document.querySelector("#modal-title"),
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

  function portOptions(selected) {
    const current = C.normalizePort(selected);
    return `<option value="" ${current ? "" : "selected"}>Seleccionar</option>${uniquePorts()
      .map((port) => `<option value="${escapeHtml(port.code)}" ${port.code === current ? "selected" : ""}>${escapeHtml(port.code)} · ${escapeHtml(port.label)}</option>`)
      .join("")}`;
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

    elements.dropzone.classList.add("is-busy");
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
      renderAll();
      window.scrollTo({ top: 0, behavior: "smooth" });
      const expiredDocuments = parsed.records.filter((record) => C.documentExpiry(record) === "caducada").length;
      const expiredNotice = expiredDocuments
        ? ` ${expiredDocuments} ${expiredDocuments === 1 ? "documento con la caducidad cumplida" : "documentos con la caducidad cumplida"}.`
        : "";
      showToast(`${parsed.records.length} tripulantes importados correctamente.${expiredNotice}`);
    } catch (error) {
      console.error(error);
      showToast(error && error.message ? error.message : "No se ha podido leer el archivo.", "error");
    } finally {
      elements.dropzone.classList.remove("is-busy");
      elements.fileInput.value = "";
    }
  }

  function renderSummary() {
    elements.dropLabel.textContent = state.sourceFile ? state.sourceFile.name : "Introducir archivo Excel o CSV";
  }

  function blankMovement() {
    return {
      id: state.nextMovementId++,
      reference: "",
      scheduleDate: "",
      scheduleTime: "",
      embarkPort: "",
      disembarkPort: "",
      embarkDate: "",
      embarkTime: "",
      disembarkDate: "",
      disembarkTime: "",
    };
  }

  function issueText(validation) {
    const parts = [];
    if (validation.missing.length) parts.push(`Faltan: ${validation.missing.join(", ")}`);
    if (validation.warnings.length) parts.push(validation.warnings.join(" · "));
    if (validation.notices && validation.notices.length) parts.push(validation.notices.join(" · "));
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
            <td><select class="entsal-select" data-entsal="${index}" aria-label="EntSal"><option value="" ${shownEntSal(record) ? "" : "selected"}>Seleccionar</option><option value="S" ${shownEntSal(record) === "S" ? "selected" : ""}>S · Salida</option><option value="E" ${shownEntSal(record) === "E" ? "selected" : ""}>E · Entrada</option></select></td>
            <td><span class="status-badge ${validation.valid ? "" : "issue"}" title="${escapeHtml(issue)}">${validation.valid ? "Correcto" : "Revisar"}</span></td>
            <td><div class="row-actions"><button class="edit-button" type="button" data-edit-record="${index}">Editar</button><button class="edit-button edit-button--danger" type="button" data-delete-record="${index}">Eliminar</button></div></td>
          </tr>`;
      })
      .join("");

    elements.emptyTable.textContent = state.records.length === 0
      ? "Todavía no hay tripulantes. Añade el primero para poder generar el CSV."
      : "No hay tripulantes que coincidan con la búsqueda.";
    elements.emptyTable.classList.toggle("is-hidden", visible.length !== 0);
    elements.tableResultCount.textContent = `${visible.length} ${visible.length === 1 ? "registro" : "registros"}`;
  }

  function movementTemplate(movement) {
    const errors = C.movementErrors(movement);
    const invalid = (field) => {
      const checks = {
        embarkPort: !C.normalizePort(movement.embarkPort),
        disembarkPort: !C.normalizePort(movement.disembarkPort),
        embarkDate: !movement.embarkDate,
        embarkTime: !movement.embarkTime,
        disembarkDate: !movement.disembarkDate,
        disembarkTime: !movement.disembarkTime,
      };
      return checks[field] ? 'aria-invalid="true"' : "";
    };

    return `
      <article data-movement-id="${movement.id}">
        <fieldset class="schedule-fieldset">
          <legend>Datos</legend>
          <div class="field-grid schedule-grid">
            <div class="field"><label>Puerto de embarque</label><select data-field="embarkPort" ${invalid("embarkPort")}>${portOptions(movement.embarkPort)}</select></div>
            <div class="field"><label>Puerto de desembarque</label><select data-field="disembarkPort" ${invalid("disembarkPort")}>${portOptions(movement.disembarkPort)}</select></div>
            <div class="schedule-dates">
              <div class="field">
                <label>Fecha de embarque</label>
                <div class="datetime-pair">
                  <input type="date" data-field="embarkDate" value="${escapeHtml(movement.embarkDate)}" aria-label="Fecha de embarque" ${invalid("embarkDate")} />
                  <input type="time" data-field="embarkTime" value="${escapeHtml(movement.embarkTime)}" aria-label="Hora de embarque" ${invalid("embarkTime")} />
                </div>
              </div>
              <div class="field">
                <label>Fecha de desembarque</label>
                <div class="datetime-pair">
                  <input type="date" data-field="disembarkDate" value="${escapeHtml(movement.disembarkDate)}" aria-label="Fecha de desembarque" ${invalid("disembarkDate")} />
                  <input type="time" data-field="disembarkTime" value="${escapeHtml(movement.disembarkTime)}" aria-label="Hora de desembarque" ${invalid("disembarkTime")} />
                </div>
              </div>
            </div>
          </div>
        </fieldset>
      </article>`;
  }

  function renderMovements() {
    elements.movementList.innerHTML = state.movements.map(movementTemplate).join("");
  }

  function refreshMovementCard(movement) {
    const card = elements.movementList.querySelector(`[data-movement-id="${movement.id}"]`);
    if (!card) return;
    card.classList.toggle("has-errors", C.movementErrors(movement).length > 0);

    const fieldChecks = {
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

  function updateExportState() {
    const crewIssues = invalidCrewCount();
    const validMovements = state.movements.filter((movement) => C.movementErrors(movement).length === 0).length;
    const allMovementsValid = validMovements === state.movements.length && state.movements.length > 0;
    const ready = state.records.length > 0 && crewIssues === 0 && allMovementsValid;

    elements.exportAll.disabled = !ready;
    const exportLabel = elements.exportAll.querySelector("span");
    exportLabel.textContent = state.movements.length > 1 ? `Descargar ${state.movements.length} archivos` : "Descargar CSV";
  }

  function renderAll() {
    renderSummary();
    renderCrew();
    renderMovements();
    updateExportState();
  }

  function shownEntSal(record) {
    return record && (record.EntSal === "E" || record.EntSal === "S") ? record.EntSal : "";
  }

  function fillCategoryOptions(selected) {
    const select = elements.crewForm.elements.Categoria;
    const codes = Object.keys(C.CATEGORY_DESCRIPTIONS);
    const current = String(selected || "").toUpperCase();
    const options = codes.map((code) => ({ code, label: C.CATEGORY_DESCRIPTIONS[code] }));
    if (current && !codes.includes(current)) options.push({ code: current, label: "Categoría del archivo" });
    select.innerHTML = `<option value="">Seleccionar</option>${options
      .map((option) => `<option value="${escapeHtml(option.code)}">${escapeHtml(option.code)} · ${escapeHtml(option.label)}</option>`)
      .join("")}`;
    select.value = current;
  }

  function openCrewEditor(index) {
    const isNew = index == null || Number.isNaN(Number(index));
    const record = isNew ? null : state.records[index];
    if (!isNew && !record) return;
    elements.crewForm.reset();
    elements.crewForm.elements.recordIndex.value = isNew ? "" : String(index);
    fillCategoryOptions(record ? record.Categoria : "");
    if (record) {
      CREW_FIELDS.forEach((field) => {
        if (field === "Categoria") return;
        elements.crewForm.elements[field].value = record[field] || "";
      });
    }
    if (isNew) elements.crewForm.elements.EntSal.value = "S";
    elements.modalTitle.textContent = isNew
      ? "Nuevo tripulante"
      : (`${record.Nombre} ${record.Apellidos}`.trim() || "Datos del tripulante");
    elements.saveCrew.textContent = isNew ? "Añadir tripulante" : "Guardar cambios";
    elements.crewDialog.showModal();
    elements.crewForm.elements.Nombre.focus();
  }

  function saveCrewRecord() {
    if (!elements.crewForm.reportValidity()) return;
    const indexValue = elements.crewForm.elements.recordIndex.value;
    const isNew = indexValue === "";
    const index = Number(indexValue);
    if (!isNew && !state.records[index]) return;
    const updated = isNew ? {} : { ...state.records[index] };
    CREW_FIELDS.forEach((field) => {
      updated[field] = elements.crewForm.elements[field].value;
    });
    const normalized = C.normalizeRecord(updated);
    if (isNew) state.records.push(normalized);
    else state.records[index] = normalized;
    elements.crewDialog.close();
    renderSummary();
    renderCrew();
    updateExportState();
    showToast(isNew ? "Tripulante añadido." : "Registro actualizado.");
  }

  function deleteCrewRecord(index) {
    const record = state.records[index];
    if (!record) return;
    const name = `${record.Nombre} ${record.Apellidos}`.trim() || "este tripulante";
    if (!window.confirm(`¿Eliminar a ${name}?`)) return;
    state.records.splice(index, 1);
    renderCrew();
    updateExportState();
    showToast("Tripulante eliminado.");
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
    if (!state.records.length || invalidCrewCount() > 0 || state.movements.some((movement) => C.movementErrors(movement).length)) {
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

  elements.fileInput.addEventListener("change", (event) => importFile(event.target.files[0]));
  elements.addCrew.addEventListener("click", () => openCrewEditor());
  elements.crewForm.elements.Categoria.addEventListener("change", () => {
    const description = C.CATEGORY_DESCRIPTIONS[elements.crewForm.elements.Categoria.value];
    if (description) elements.crewForm.elements.CategoriaDes.value = description;
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    elements.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropzone.classList.add("is-drag");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    elements.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropzone.classList.remove("is-drag");
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
  elements.crewBody.addEventListener("change", (event) => {
    const select = event.target.closest("[data-entsal]");
    if (!select) return;
    const index = Number(select.dataset.entsal);
    if (!state.records[index]) return;
    state.records[index].EntSal = select.value === "E" || select.value === "S" ? select.value : "";
    renderCrew();
    updateExportState();
  });
  elements.crewBody.addEventListener("click", (event) => {
    const editButton = event.target.closest("[data-edit-record]");
    if (editButton) openCrewEditor(Number(editButton.dataset.editRecord));
    const deleteButton = event.target.closest("[data-delete-record]");
    if (deleteButton) deleteCrewRecord(Number(deleteButton.dataset.deleteRecord));
  });

  elements.crewForm.addEventListener("submit", (event) => {
    event.preventDefault();
    if (event.submitter && event.submitter.value === "cancel") {
      elements.crewDialog.close();
      return;
    }
    saveCrewRecord();
  });

  elements.movementList.addEventListener("input", (event) => {
    const input = event.target.closest("[data-field]");
    const card = event.target.closest("[data-movement-id]");
    if (!input || !card) return;
    const movement = movementById(card.dataset.movementId);
    if (!movement) return;
    let value = input.value;
    movement[input.dataset.field] = value;
    refreshMovementCard(movement);
    updateExportState();
  });
  elements.exportAll.addEventListener("click", downloadAll);

  window.addEventListener("dragover", (event) => event.preventDefault());
  window.addEventListener("drop", (event) => {
    if (!event.target.closest || !event.target.closest("#dropzone")) event.preventDefault();
  });

  state.movements = [blankMovement()];
  renderAll();
})();
