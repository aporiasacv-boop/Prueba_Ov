// Probar conexion — script independiente (~250 lineas).
// IMPORTANTE: cree un script NUEVO o borre TODO el codigo viejo antes de pegar.

interface PcHealthApi {
  success: boolean;
  message: string;
}

const PC_TABLA_RESULTADOS = "tblResultados";
const PC_HOJA_RESULTADO = "Resultado";
const PC_CELDA_API = "B1";
const PC_VERSION = "2026-06-19-probar-v3";

function pcEsVacio(valor: string | number | boolean): boolean {
  if (typeof valor === "string") {
    return valor.trim() === "";
  }
  return false;
}

function pcATexto(valor: string | number | boolean): string {
  if (pcEsVacio(valor)) {
    return "";
  }
  return String(valor).trim();
}

function pcLeerEncabezados(tabla: ExcelScript.Table): string[] {
  const fila = tabla.getHeaderRowRange().getValues()[0];
  const headers: string[] = [];
  for (let i = 0; i < fila.length; i++) {
    headers.push(pcATexto(fila[i] as string | number | boolean));
  }
  return headers;
}

function pcLeerFilas(tabla: ExcelScript.Table): (string | number | boolean)[][] {
  const dataValues = tabla.getRangeBetweenHeaderAndTotal().getValues();
  if (dataValues.length === 0) {
    return [];
  }
  return dataValues as (string | number | boolean)[][];
}

function pcIndiceColumna(headers: string[], nombre: string): number {
  return headers.indexOf(nombre);
}

function pcIndiceOv(headers: string[]): number {
  let idx = headers.indexOf("PedidoDynamics");
  if (idx < 0) {
    idx = headers.indexOf("Pedido Dynamics");
  }
  if (idx < 0) {
    idx = headers.indexOf("Pedido Dynamic");
  }
  return idx;
}

function pcLeerOv(workbook: ExcelScript.Workbook): string {
  const tabla = workbook.getTable(PC_TABLA_RESULTADOS);
  if (!tabla) {
    return "";
  }
  const headers = pcLeerEncabezados(tabla);
  const filas = pcLeerFilas(tabla);
  if (filas.length === 0) {
    return "";
  }
  const idx = pcIndiceOv(headers);
  if (idx < 0) {
    return "";
  }
  return pcATexto(filas[0][idx]);
}

function pcLeerUrl(workbook: ExcelScript.Workbook): string {
  const hoja = workbook.getWorksheet(PC_HOJA_RESULTADO);
  if (!hoja) {
    throw new Error("No existe la hoja Resultado.");
  }

  let url = pcATexto(hoja.getRange(PC_CELDA_API).getValue() as string | number | boolean);
  if (url === "") {
    url = pcATexto(hoja.getRange("A2").getValue() as string | number | boolean);
  }
  if (url === "") {
    throw new Error("Escriba la URL ngrok en Resultado!B1 (https://...)");
  }
  if (url.indexOf("http") !== 0) {
    throw new Error("La URL debe empezar con https://");
  }
  if (url.endsWith("/")) {
    return url.substring(0, url.length - 1);
  }
  return url;
}

function pcHeadersFetch(conCuerpo: boolean): { [key: string]: string } {
  const h: { [key: string]: string } = {
    "ngrok-skip-browser-warning": "true",
  };
  if (conCuerpo) {
    h["Content-Type"] = "application/json";
  }
  return h;
}

async function pcFetch(
  url: string,
  metodo: string,
  cuerpo: string | null
): Promise<object> {
  const conCuerpo = cuerpo !== null && cuerpo !== "";
  const headers = pcHeadersFetch(conCuerpo);
  const response = conCuerpo
    ? await fetch(url, { method: metodo, headers: headers, body: cuerpo })
    : await fetch(url, { method: metodo, headers: headers });
  const texto = await response.text();
  if (!response.ok) {
    throw new Error("HTTP " + response.status + ": " + texto);
  }
  if (texto === "") {
    return {};
  }
  return JSON.parse(texto) as object;
}

function pcMensajeError(error: string | number | boolean | object): string {
  let msg = "";
  if (typeof error === "object" && error !== null) {
    const errObj = error as { message?: string };
    if (errObj.message) {
      msg = String(errObj.message);
    }
  }
  if (msg === "") {
    msg = String(error);
  }
  if (msg.indexOf("failed to fetch") >= 0) {
    return "No se pudo conectar. Revise ngrok en Resultado!B1 y el backend.";
  }
  return msg;
}

function pcEscribirResultado(
  workbook: ExcelScript.Workbook,
  estado: string,
  mensaje: string,
  ov: string
): void {
  const tabla = workbook.getTable(PC_TABLA_RESULTADOS);
  if (!tabla) {
    return;
  }
  if (tabla.getRowCount() === 0) {
    tabla.addRow();
  }

  const headers = pcLeerEncabezados(tabla);
  const idxEstado = pcIndiceColumna(headers, "Estado");
  const idxOv = pcIndiceOv(headers);
  let idxFecha = pcIndiceColumna(headers, "Fecha de ejecucion");
  if (idxFecha < 0) {
    idxFecha = pcIndiceColumna(headers, "Fecha de ejecución");
  }
  const idxUsuario = pcIndiceColumna(headers, "Usuario");
  const idxError = pcIndiceColumna(headers, "Error");

  const cuerpo = tabla.getRangeBetweenHeaderAndTotal();
  if (cuerpo.getValues().length === 0) {
    return;
  }

  const fila = 0;
  const ahora = new Date();
  if (idxEstado >= 0) {
    cuerpo.getCell(fila, idxEstado).setValue(estado);
  }
  if (idxOv >= 0) {
    cuerpo.getCell(fila, idxOv).setValue(ov);
  }
  if (idxFecha >= 0) {
    cuerpo.getCell(fila, idxFecha).setValue(ahora);
  }
  if (idxUsuario >= 0) {
    cuerpo.getCell(fila, idxUsuario).setValue("");
  }
  if (idxError >= 0) {
    cuerpo.getCell(fila, idxError).setValue(mensaje);
  }
}

async function pcProbarConexion(workbook: ExcelScript.Workbook): Promise<void> {
  const ov = pcLeerOv(workbook);
  try {
    const base = pcLeerUrl(workbook);
    const resp = (await pcFetch(
      base + "/api/health/dynamics",
      "GET",
      null
    )) as PcHealthApi;

    if (resp.success) {
      pcEscribirResultado(
        workbook,
        "CONEXION OK",
        (resp.message || "Conexion OK") + " [" + PC_VERSION + "]",
        ov
      );
      return;
    }

    pcEscribirResultado(
      workbook,
      "ERROR",
      resp.message || "Fallo de conexion",
      ov
    );
  } catch (error) {
    pcEscribirResultado(workbook, "ERROR", pcMensajeError(error), ov);
  }
}

async function main(workbook: ExcelScript.Workbook): Promise<void> {
  await pcProbarConexion(workbook);
}
