type ValorCelda = string | number | boolean;

interface TablaLeida {
  headers: string[];
  rows: ValorCelda[][];
}

interface HealthConexionApi {
  success: boolean;
  message: string;
}

const TABLA_RESULTADOS = "tblResultados";
const HOJA_RESULTADO = "Resultado";
const CELDA_API_BASE = "B1";
const COL_RESULTADO_ESTADO = "Estado";
const COL_RESULTADO_FECHA = "Fecha de ejecucion";
const COL_RESULTADO_FECHA_ALT = "Fecha de ejecución";
const COL_RESULTADO_USUARIO = "Usuario";
const COL_RESULTADO_ERROR = "Error";
const SCRIPT_PROBAR_CONEXION_VERSION = "2026-06-19-probar-v1";

function esVacio(valor: ValorCelda | undefined): boolean {
  if (valor === null || valor === undefined) {
    return true;
  }
  if (typeof valor === "string") {
    return valor.trim() === "";
  }
  return false;
}

function aTexto(valor: ValorCelda | undefined): string {
  if (esVacio(valor)) {
    return "";
  }
  return String(valor).trim();
}

function leerTabla(tabla: ExcelScript.Table): TablaLeida {
  const headerRange = tabla.getHeaderRowRange();
  const headerValues = headerRange.getValues();
  const filaEncabezado = headerValues[0] as ValorCelda[];
  const headers: string[] = [];
  for (let i = 0; i < filaEncabezado.length; i++) {
    headers.push(String(filaEncabezado[i]).trim());
  }

  const dataRange = tabla.getRangeBetweenHeaderAndTotal();
  const dataValues = dataRange.getValues();
  const rows: ValorCelda[][] =
    dataValues.length > 0 ? (dataValues as ValorCelda[][]) : [];

  return { headers: headers, rows: rows };
}

function indiceColumna(headers: string[], nombreColumna: string): number {
  return headers.indexOf(nombreColumna);
}

function indiceColumnaPedidoDynamics(headers: string[]): number {
  let idx = headers.indexOf("PedidoDynamics");
  if (idx < 0) {
    idx = headers.indexOf("Pedido Dynamics");
  }
  if (idx < 0) {
    idx = headers.indexOf("Pedido Dynamic");
  }
  return idx;
}

function leerPedidoDynamicsDesdeResultados(workbook: ExcelScript.Workbook): string {
  const tabla = workbook.getTable(TABLA_RESULTADOS);
  if (!tabla) {
    return "";
  }

  const datos = leerTabla(tabla);
  if (datos.rows.length === 0) {
    return "";
  }

  const idx = indiceColumnaPedidoDynamics(datos.headers);
  if (idx < 0) {
    return "";
  }

  return aTexto(datos.rows[0][idx]);
}

function leerApiBaseUrl(workbook: ExcelScript.Workbook): string {
  const hoja = workbook.getWorksheet(HOJA_RESULTADO);
  if (!hoja) {
    throw new Error("No existe la hoja '" + HOJA_RESULTADO + "'.");
  }

  let url = aTexto(hoja.getRange(CELDA_API_BASE).getValue() as ValorCelda);
  if (url === "") {
    url = aTexto(hoja.getRange("A2").getValue() as ValorCelda);
  }

  if (url === "") {
    throw new Error(
      "Escribe la URL ngrok en Resultado!B1 o A2 (ej. https://xxxx.ngrok-free.dev)"
    );
  }

  if (url.indexOf("http") !== 0) {
    throw new Error("La URL debe empezar con https://");
  }

  if (url.endsWith("/")) {
    return url.substring(0, url.length - 1);
  }
  return url;
}

function headersApi(conCuerpo: boolean): { [key: string]: string } {
  const headers: { [key: string]: string } = {
    "ngrok-skip-browser-warning": "true",
  };
  if (conCuerpo) {
    headers["Content-Type"] = "application/json";
  }
  return headers;
}

async function llamarApi(
  url: string,
  metodo: string,
  cuerpo: string | null
): Promise<object> {
  const conCuerpo = cuerpo !== null && cuerpo !== "";
  const headers = headersApi(conCuerpo);
  const response = conCuerpo
    ? await fetch(url, {
        method: metodo,
        headers: headers,
        body: cuerpo,
      })
    : await fetch(url, {
        method: metodo,
        headers: headers,
      });
  const textoRespuesta = await response.text();

  if (!response.ok) {
    throw new Error("HTTP " + response.status + ": " + textoRespuesta);
  }

  if (textoRespuesta === "") {
    return {};
  }

  return JSON.parse(textoRespuesta) as object;
}

function obtenerMensajeError(error: string | number | boolean | object): string {
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
    return "No se pudo conectar con la API. Revise ngrok en Resultado!B1 y que el backend este activo.";
  }
  return msg;
}

function escribirResultado(
  workbook: ExcelScript.Workbook,
  estado: string,
  mensajeError: string,
  pedidoDynamics: string
): void {
  const tabla = workbook.getTable(TABLA_RESULTADOS);
  if (!tabla) {
    return;
  }

  if (tabla.getRowCount() === 0) {
    tabla.addRow();
  }

  const datos = leerTabla(tabla);
  const headers = datos.headers;

  const idxEstado = indiceColumna(headers, COL_RESULTADO_ESTADO);
  const idxPedido = indiceColumnaPedidoDynamics(headers);
  let idxFecha = indiceColumna(headers, COL_RESULTADO_FECHA);
  if (idxFecha < 0) {
    idxFecha = indiceColumna(headers, COL_RESULTADO_FECHA_ALT);
  }
  const idxUsuario = indiceColumna(headers, COL_RESULTADO_USUARIO);
  const idxError = indiceColumna(headers, COL_RESULTADO_ERROR);

  const cuerpo = tabla.getRangeBetweenHeaderAndTotal();
  const filasCuerpo = cuerpo.getValues();
  if (filasCuerpo.length === 0) {
    return;
  }

  const ahora = new Date();
  const fila = 0;

  if (idxEstado >= 0) {
    cuerpo.getCell(fila, idxEstado).setValue(estado);
  }
  if (idxPedido >= 0) {
    cuerpo.getCell(fila, idxPedido).setValue(pedidoDynamics);
  }
  if (idxFecha >= 0) {
    cuerpo.getCell(fila, idxFecha).setValue(ahora);
  }
  if (idxUsuario >= 0) {
    cuerpo.getCell(fila, idxUsuario).setValue("");
  }
  if (idxError >= 0) {
    cuerpo.getCell(fila, idxError).setValue(mensajeError);
  }
}

async function probarConexion(workbook: ExcelScript.Workbook): Promise<void> {
  const ov = leerPedidoDynamicsDesdeResultados(workbook);
  try {
    const base = leerApiBaseUrl(workbook);
    const respuesta = (await llamarApi(
      base + "/api/health/dynamics",
      "GET",
      null
    )) as HealthConexionApi;

    if (respuesta.success) {
      const msg =
        (respuesta.message || "Conexion OK") +
        " [" +
        SCRIPT_PROBAR_CONEXION_VERSION +
        "]";
      escribirResultado(workbook, "CONEXION OK", msg, ov);
      return;
    }

    escribirResultado(
      workbook,
      "ERROR",
      respuesta.message || "Fallo de conexion",
      ov
    );
  } catch (error) {
    escribirResultado(workbook, "ERROR", obtenerMensajeError(error), ov);
  }
}
