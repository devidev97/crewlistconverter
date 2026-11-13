"use strict";

const assert = require("assert");
const XLSX = require("../public/vendor/xlsx.full.min.js");
const C = require("../converter.js");

const sourceHeaders = C.HEADERS.filter((header) => header !== "SchengenArea").map((header) => header === "Observaciones" ? "Observacion" : header);
sourceHeaders.splice(2, 0, "");

function sourceRow(values) {
  const record = {
    Tipo: "T",
    EntSal: "",
    Nombre: "PERSONA",
    Apellidos: "DE PRUEBA",
    TipoDocumento: "ZZP",
    NumeroDocumento: "TEST0001",
    NumeroVisado: "",
    FechaCaducidadDocumento: "08/09/2030",
    Nacionalidad: "ES",
    FechaNacimiento: "12/12/1990",
    LugarNacimiento: "ES",
    Categoria: "ZOT",
    CategoriaDes: "Others",
    Sexo: "F",
    MatriculaVehiculo: "",
    MarcaVehiculo: "",
    ModeloVehiculo: "",
    PtoEmbarque: "",
    PtoDesembarque: "",
    FechaEmbarque: "12/11/2026",
    FechaDesembarque: "12/11/2026",
    TarjetaEmbarque: "",
    CouponBillete: "",
    Naviera: "",
    TipoAcreditacion: "",
    Observacion: "",
    CuidadosAsistenciasEspeciales: "",
    ContactoEmergencia: "",
    ...values,
  };
  return sourceHeaders.map((header) => header ? record[header] || "" : "");
}

const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.aoa_to_sheet([
  sourceHeaders,
  sourceRow({ Naviera: "DFDS" }),
  sourceRow({
    Nombre: "CAPITAN",
    Apellidos: "DE PRUEBA",
    NumeroDocumento: "TEST0002",
    Categoria: "ZCP",
    CategoriaDes: "Captain",
    Sexo: "M",
    Naviera: "BALEARIA",
  }),
]);
XLSX.utils.book_append_sheet(workbook, worksheet, "2000 Alg- Ceu");
const source = XLSX.write(workbook, { type: "array", bookType: "xls" });
const parsed = C.parseWorkbook(source, XLSX);

assert.strictEqual(parsed.sheetName, "2000 Alg- Ceu");
assert.strictEqual(parsed.records.length, 2);
assert.strictEqual(parsed.records[0].CategoriaDes, "Otros");
assert.strictEqual(parsed.records[1].CategoriaDes, "Capitan");
assert.strictEqual(C.validateRecord(parsed.records[0]).valid, true);

const inferred = C.inferMovement("2000 Alg- Ceu.xls", parsed.records);
assert.strictEqual(inferred.embarkPort, "ESALG");
assert.strictEqual(inferred.disembarkPort, "ESCEU");
assert.strictEqual(inferred.scheduleTime, "20:00");
assert.strictEqual(inferred.embarkTime, "");
assert.strictEqual(inferred.embarkDate, "2026-11-12");

const movement = {
  direction: "S",
  reference: "30001",
  scheduleDate: "2026-11-12",
  scheduleTime: "20:00",
  embarkPort: "ESALG",
  disembarkPort: "ESCEU",
  embarkDate: "2026-11-12",
  embarkTime: "20:00",
  disembarkDate: "2026-11-12",
  disembarkTime: "21:30",
  shippingCompany: "NO-USAR",
};

const csv = C.buildCsv(parsed.records, movement);
const lines = csv.split("\r\n").filter(Boolean);
assert.strictEqual(lines.length, 3);
assert.ok(lines.every((line) => line.split(";").length === 29));
assert.strictEqual(lines[0], C.HEADERS.join(";"));
assert.ok(lines[1].startsWith("T;S;PERSONA;DE PRUEBA;"));
assert.ok(lines[1].includes(";ESALG;ESCEU;12/11/2026 20:00;12/11/2026 21:30;"));
assert.strictEqual(lines[1].split(";")[C.HEADERS.indexOf("Naviera")], "DFDS");
assert.strictEqual(lines[2].split(";")[C.HEADERS.indexOf("Naviera")], "BALEARIA");
assert.strictEqual(C.buildFilename(movement), "Lista de tripulantes.csv");
assert.ok(csv.endsWith("\r\n"));
const roundTrip = C.parseWorkbook(new TextEncoder().encode(csv), XLSX);
assert.strictEqual(roundTrip.records.length, 2);
assert.strictEqual(roundTrip.records[0].PtoEmbarque, "ESALG");
assert.strictEqual(roundTrip.records[0].FechaEmbarque, "12/11/2026 20:00");

for (const sample of [
  {
    name: "2026-29008-1211-0500-PTM-ALG-Crew.csv",
    movement: { ...movement, direction: "E", reference: "29008", scheduleTime: "05:00", embarkPort: "MAPTM", disembarkPort: "ESALG", embarkDate: "2026-11-09", embarkTime: "05:00", disembarkDate: "2026-11-09", disembarkTime: "07:30" },
  },
  {
    name: "2026-29941-1011-0800-ALG-PTM-Crew(1).csv",
    movement: { ...movement, reference: "29941", scheduleDate: "2026-11-10", scheduleTime: "08:00", embarkPort: "ESALG", disembarkPort: "MAPTM", embarkDate: "2026-11-10", embarkTime: "09:00", disembarkDate: "2026-11-10", disembarkTime: "09:30" },
  },
]) {
  const accepted = C.buildCsv(parsed.records, sample.movement);
  const acceptedParsed = C.parseWorkbook(new TextEncoder().encode(accepted), XLSX);
  assert.strictEqual(acceptedParsed.records.length, 2);
  const acceptedMovement = C.inferMovement(sample.name, acceptedParsed.records);
  assert.strictEqual(C.buildFilename(acceptedMovement), "Lista de tripulantes.csv");
  assert.strictEqual(acceptedMovement.embarkPort, sample.movement.embarkPort);
  assert.strictEqual(acceptedMovement.disembarkPort, sample.movement.disembarkPort);
}

const today = new Date(2026, 10, 13);
const passport = { ...parsed.records[0], TipoDocumento: "PASAPORTE", FechaCaducidadDocumento: "13/11/2026" };
const identityCard = { ...parsed.records[0], TipoDocumento: "DNI", FechaCaducidadDocumento: "08/09/2030" };
assert.strictEqual(C.documentExpiry(passport, today), "vigente");
assert.strictEqual(C.validateRecord(passport, today).valid, true);
assert.strictEqual(C.documentExpiry({ ...passport, FechaCaducidadDocumento: "12/11/2026" }, today), "caducada");
const expiredPassport = C.validateRecord({ ...passport, FechaCaducidadDocumento: "12/11/2026" }, today);
assert.strictEqual(expiredPassport.valid, false);
assert.ok(expiredPassport.warnings.includes("Documentación caducada"));
assert.strictEqual(C.documentExpiry(identityCard, today), "vigente");
assert.strictEqual(C.validateRecord(identityCard, today).valid, true);

console.log("OK: conversión XLS/CSV, 29 columnas, fechas, ruta y nombre de archivo.");
