(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  root.CrewConverter = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const HEADERS = [
    "Tipo",
    "EntSal",
    "Nombre",
    "Apellidos",
    "TipoDocumento",
    "NumeroDocumento",
    "NumeroVisado",
    "FechaCaducidadDocumento",
    "Nacionalidad",
    "FechaNacimiento",
    "LugarNacimiento",
    "Categoria",
    "CategoriaDes",
    "Sexo",
    "MatriculaVehiculo",
    "MarcaVehiculo",
    "ModeloVehiculo",
    "PtoEmbarque",
    "PtoDesembarque",
    "FechaEmbarque",
    "FechaDesembarque",
    "TarjetaEmbarque",
    "CouponBillete",
    "Naviera",
    "TipoAcreditacion",
    "Observaciones",
    "CuidadosAsistenciasEspeciales",
    "ContactoEmergencia",
    "SchengenArea",
  ];

  const REQUIRED_CREW_FIELDS = [
    "Nombre",
    "Apellidos",
    "TipoDocumento",
    "NumeroDocumento",
    "FechaCaducidadDocumento",
    "Nacionalidad",
    "FechaNacimiento",
    "LugarNacimiento",
    "Categoria",
    "Sexo",
  ];

  const CATEGORY_DESCRIPTIONS = {
    ZCP: "Capitan",
    ZO1: "1er.Oficial",
    ZO2: "2o Oficial",
    ZO3: "3er.Oficial",
    ZJM: "Jefe de maquinas",
    ZM1: "1er.Oficial de maquinas",
    ZM2: "2o Oficial de maquinas",
    ZM3: "3er.Oficial de maquinas",
    ZMC: "Marinero de cubierta",
    ZMM: "Marinero de maquinas",
    ZOT: "Otros",
  };

  const PORTS = {
    ALG: { code: "ESALG", label: "Algeciras" },
    ESALG: { code: "ESALG", label: "Algeciras" },
    CEU: { code: "ESCEU", label: "Ceuta" },
    ESCEU: { code: "ESCEU", label: "Ceuta" },
    PTM: { code: "MAPTM", label: "Tánger Med" },
    MAPTM: { code: "MAPTM", label: "Tánger Med" },
    TNG: { code: "MATNG", label: "Tánger" },
    MATNG: { code: "MATNG", label: "Tánger" },
    TAR: { code: "ESTAR", label: "Tarifa" },
    ESTAR: { code: "ESTAR", label: "Tarifa" },
  };

  const HEADER_ALIASES = {
    observacion: "Observaciones",
    observaciones: "Observaciones",
    fechacaducidaddocumento: "FechaCaducidadDocumento",
    fechacaducidad: "FechaCaducidadDocumento",
    numerodocumento: "NumeroDocumento",
    documento: "NumeroDocumento",
    tipodocumento: "TipoDocumento",
    numerovisado: "NumeroVisado",
    fechanacimiento: "FechaNacimiento",
    lugarnacimiento: "LugarNacimiento",
    categoriades: "CategoriaDes",
    descripcioncategoria: "CategoriaDes",
    ptoembarque: "PtoEmbarque",
    puertoembarque: "PtoEmbarque",
    ptodesembarque: "PtoDesembarque",
    puertodesembarque: "PtoDesembarque",
    fechaembarque: "FechaEmbarque",
    fechadesembarque: "FechaDesembarque",
    entrasal: "EntSal",
    entsal: "EntSal",
    matriculavehiculo: "MatriculaVehiculo",
    marcavehiculo: "MarcaVehiculo",
    modelovehiculo: "ModeloVehiculo",
    tarjetaembarque: "TarjetaEmbarque",
    couponbillete: "CouponBillete",
    tipoacreditacion: "TipoAcreditacion",
    cuidadosasistenciasespeciales: "CuidadosAsistenciasEspeciales",
    contactoemergencia: "ContactoEmergencia",
    schengenarea: "SchengenArea",
  };

  HEADERS.forEach((header) => {
    HEADER_ALIASES[normalizeKey(header)] = header;
  });

  function normalizeKey(value) {
    return String(value == null ? "" : value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9]/g, "")
      .toLowerCase();
  }

  function text(value) {
    return String(value == null ? "" : value)
      .replace(/[\r\n\t]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function pad(value) {
    return String(value).padStart(2, "0");
  }

  function excelSerialToDate(serial) {
    const utcDays = Math.floor(Number(serial) - 25569);
    const utcValue = utcDays * 86400;
    const date = new Date(utcValue * 1000);
    return Number.isNaN(date.getTime()) ? "" : `${pad(date.getUTCDate())}/${pad(date.getUTCMonth() + 1)}/${date.getUTCFullYear()}`;
  }

  function normalizeDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return `${pad(value.getDate())}/${pad(value.getMonth() + 1)}/${value.getFullYear()}`;
    }

    const source = text(value);
    if (!source) return "";
    if (/^\d{5}(?:\.\d+)?$/.test(source)) return excelSerialToDate(source);

    let match = source.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})(?:\s+.*)?$/);
    if (match) return `${pad(match[1])}/${pad(match[2])}/${match[3]}`;

    match = source.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
    if (match) return `${pad(match[3])}/${pad(match[2])}/${match[1]}`;

    return source;
  }

  function toIsoDate(value) {
    const date = normalizeDate(value);
    const match = date.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    return match ? `${match[3]}-${match[2]}-${match[1]}` : "";
  }

  function normalizePort(value) {
    const source = text(value).toUpperCase().replace(/\s/g, "");
    return PORTS[source] ? PORTS[source].code : source;
  }

  function shortPort(value) {
    const code = normalizePort(value);
    return code.length === 5 ? code.slice(2) : code;
  }

  function normalizeRecord(input) {
    const record = {};
    HEADERS.forEach((header) => {
      record[header] = text(input[header]);
    });

    record.Tipo = record.Tipo || "T";
    record.FechaCaducidadDocumento = normalizeDate(record.FechaCaducidadDocumento);
    record.FechaNacimiento = normalizeDate(record.FechaNacimiento);
    record.Nacionalidad = record.Nacionalidad.toUpperCase();
    record.Categoria = record.Categoria.toUpperCase();
    record.Sexo = record.Sexo.toUpperCase();

    if (CATEGORY_DESCRIPTIONS[record.Categoria]) {
      record.CategoriaDes = CATEGORY_DESCRIPTIONS[record.Categoria];
    } else if (/^others?$/i.test(record.CategoriaDes)) {
      record.CategoriaDes = "Otros";
    } else if (/^captain$/i.test(record.CategoriaDes)) {
      record.CategoriaDes = "Capitan";
    }

    return record;
  }

  function findHeaderRow(rows) {
    let best = { index: -1, score: 0 };
    rows.slice(0, 25).forEach((row, index) => {
      const found = new Set();
      (row || []).forEach((cell) => {
        const canonical = HEADER_ALIASES[normalizeKey(cell)];
        if (canonical) found.add(canonical);
      });
      if (found.size > best.score) best = { index, score: found.size };
    });
    if (best.score < 4) throw new Error("No se reconoce la fila de cabeceras del archivo.");
    return best.index;
  }

  function parseWorkbook(arrayBuffer, XLSX) {
    if (!XLSX || !XLSX.read) throw new Error("No se ha podido cargar el lector de Excel.");
    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: false,
      cellText: true,
      // Keep date/time text from CSV files intact instead of letting SheetJS
      // reinterpret it with the browser's regional date settings.
      raw: true,
    });

    if (!workbook.SheetNames.length) throw new Error("El archivo no contiene hojas.");

    let selected = null;
    let selectedRows = null;
    let selectedHeader = -1;
    let selectedScore = -1;

    workbook.SheetNames.forEach((sheetName) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1,
        defval: "",
        raw: false,
        blankrows: false,
      });
      try {
        const headerIndex = findHeaderRow(rows);
        const score = (rows[headerIndex] || []).filter((cell) => HEADER_ALIASES[normalizeKey(cell)]).length;
        if (score > selectedScore) {
          selected = sheetName;
          selectedRows = rows;
          selectedHeader = headerIndex;
          selectedScore = score;
        }
      } catch (_) {
        // Another sheet may contain the actual crew list.
      }
    });

    if (!selectedRows) throw new Error("No se encuentra una tabla de tripulación compatible.");

    const columnMap = {};
    selectedRows[selectedHeader].forEach((cell, columnIndex) => {
      const canonical = HEADER_ALIASES[normalizeKey(cell)];
      if (canonical && columnMap[canonical] == null) columnMap[canonical] = columnIndex;
    });

    const records = selectedRows
      .slice(selectedHeader + 1)
      .map((row) => {
        const rawRecord = {};
        HEADERS.forEach((header) => {
          const index = columnMap[header];
          rawRecord[header] = index == null ? "" : row[index];
        });
        return normalizeRecord(rawRecord);
      })
      .filter((record) => record.Nombre || record.Apellidos || record.NumeroDocumento);

    if (!records.length) throw new Error("La tabla no contiene tripulantes.");

    return {
      records,
      sheetName: selected,
      sourceHeaders: Object.keys(columnMap),
      missingHeaders: HEADERS.filter((header) => columnMap[header] == null),
    };
  }

  function validateRecord(record) {
    const missing = REQUIRED_CREW_FIELDS.filter((field) => !text(record[field]));
    const warnings = [];
    if (record.Sexo && !["M", "F"].includes(record.Sexo)) warnings.push("Sexo debe ser M o F");
    if (record.Nacionalidad && !/^[A-Z]{2}$/.test(record.Nacionalidad)) warnings.push("Nacionalidad debe tener 2 letras");
    ["FechaCaducidadDocumento", "FechaNacimiento"].forEach((field) => {
      if (record[field] && !/^\d{2}\/\d{2}\/\d{4}$/.test(record[field])) warnings.push(`${field} no tiene formato DD/MM/AAAA`);
    });
    return { valid: missing.length === 0 && warnings.length === 0, missing, warnings };
  }

  function movementErrors(movement) {
    const errors = [];
    if (!text(movement.reference)) errors.push("Número de escala");
    if (!movement.scheduleDate) errors.push("Fecha oficial del viaje");
    if (!movement.scheduleTime) errors.push("Hora oficial del viaje");
    if (!normalizePort(movement.embarkPort)) errors.push("Puerto de embarque");
    if (!normalizePort(movement.disembarkPort)) errors.push("Puerto de desembarque");
    if (!movement.embarkDate) errors.push("Fecha de embarque");
    if (!movement.embarkTime) errors.push("Hora de embarque");
    if (!movement.disembarkDate) errors.push("Fecha de desembarque");
    if (!movement.disembarkTime) errors.push("Hora de desembarque");
    if (movement.embarkPort && movement.disembarkPort && normalizePort(movement.embarkPort) === normalizePort(movement.disembarkPort)) {
      errors.push("Los puertos deben ser distintos");
    }
    return errors;
  }

  function formatMovementDate(date, time) {
    const match = String(date || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return match && time ? `${match[3]}/${match[2]}/${match[1]} ${time}` : "";
  }

  function applyMovement(record, movement) {
    const output = normalizeRecord(record);
    output.EntSal = movement.direction === "E" ? "E" : "S";
    output.PtoEmbarque = normalizePort(movement.embarkPort);
    output.PtoDesembarque = normalizePort(movement.disembarkPort);
    output.FechaEmbarque = formatMovementDate(movement.embarkDate, movement.embarkTime);
    output.FechaDesembarque = formatMovementDate(movement.disembarkDate, movement.disembarkTime);
    if (text(movement.shippingCompany)) output.Naviera = text(movement.shippingCompany);
    return output;
  }

  function safeCsvValue(value) {
    return text(value).replace(/;/g, ",");
  }

  function buildCsv(records, movement) {
    const rows = [HEADERS.join(";")];
    records.forEach((record) => {
      const output = applyMovement(record, movement);
      rows.push(HEADERS.map((header) => safeCsvValue(output[header])).join(";"));
    });
    return `${rows.join("\r\n")}\r\n`;
  }

  function sanitizeFilenamePart(value) {
    return text(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9_-]/g, "")
      .toUpperCase();
  }

  function buildFilename(movement) {
    const date = String(movement.scheduleDate || movement.embarkDate || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
    const year = date ? date[1] : "AAAA";
    const compactDate = date ? `${date[3]}${date[2]}` : "DDMM";
    const compactTime = String(movement.scheduleTime || movement.embarkTime || "HH:MM").replace(":", "");
    const reference = sanitizeFilenamePart(movement.reference) || "ESCALA";
    const origin = sanitizeFilenamePart(shortPort(movement.embarkPort)) || "ORIGEN";
    const destination = sanitizeFilenamePart(shortPort(movement.disembarkPort)) || "DESTINO";
    return `${year}-${reference}-${compactDate}-${compactTime}-${origin}-${destination}-Crew.csv`;
  }

  function inferMovement(filename, records) {
    const base = String(filename || "").replace(/\.[^.]+$/, "");
    const acceptedName = base.match(/^(\d{4})-([^-]+)-(\d{2})(\d{2})-(\d{2})(\d{2})-([A-Za-z]{3,5})-([A-Za-z]{3,5})-Crew(?:\(\d+\))?$/i);
    const simpleName = base.match(/(?:^|\s)(\d{2})(\d{2})\s+([A-Za-z]{3,5})\s*[-–>]\s*([A-Za-z]{3,5})/i);
    const first = records && records[0] ? records[0] : {};
    const date = toIsoDate(first.FechaEmbarque || first.FechaDesembarque);
    const inferredOrigin = acceptedName
      ? normalizePort(acceptedName[7])
      : simpleName
        ? normalizePort(simpleName[3])
        : normalizePort(first.PtoEmbarque);
    const inferredDestination = acceptedName
      ? normalizePort(acceptedName[8])
      : simpleName
        ? normalizePort(simpleName[4])
        : normalizePort(first.PtoDesembarque);
    const scheduleDate = acceptedName ? `${acceptedName[1]}-${acceptedName[4]}-${acceptedName[3]}` : date;
    const scheduleTime = acceptedName
      ? `${acceptedName[5]}:${acceptedName[6]}`
      : simpleName
        ? `${simpleName[1]}:${simpleName[2]}`
        : "";
    const timeFromValue = (value) => {
      const match = text(value).match(/(?:^|\s)(\d{1,2}):(\d{2})(?:\s|$)/);
      return match ? `${pad(match[1])}:${match[2]}` : "";
    };
    return {
      direction: ["E", "S"].includes(first.EntSal) ? first.EntSal : inferredDestination === "ESALG" ? "E" : "S",
      reference: acceptedName ? acceptedName[2] : "",
      scheduleDate,
      scheduleTime,
      embarkPort: inferredOrigin,
      disembarkPort: inferredDestination,
      embarkDate: date,
      embarkTime: timeFromValue(first.FechaEmbarque),
      disembarkDate: toIsoDate(first.FechaDesembarque) || date,
      disembarkTime: timeFromValue(first.FechaDesembarque),
      shippingCompany: first.Naviera || "",
    };
  }

  return {
    HEADERS,
    REQUIRED_CREW_FIELDS,
    CATEGORY_DESCRIPTIONS,
    PORTS,
    parseWorkbook,
    normalizeRecord,
    normalizeDate,
    normalizePort,
    shortPort,
    validateRecord,
    movementErrors,
    applyMovement,
    buildCsv,
    buildFilename,
    inferMovement,
    toIsoDate,
  };
});
