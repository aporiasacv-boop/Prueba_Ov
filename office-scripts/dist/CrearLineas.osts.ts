type ValorCelda = string | number | boolean;

interface CrearPedidoBody {
  cliente: string;
  referenciaCliente: string;
  descripcionPedido: string;
  fechaEnvioSolicitada: string;
  fechaRecepcionSolicitada: string;
}

interface LineaBody {
  codigoArticulo: string;
  cantidad: number;
  precioUnitario: number;
  porcentajeDescuento: number;
  fechaEnvio: string;
  comentario: string;
}

interface CrearLineasBody {
  salesOrderNumber: string;
  lineas: LineaBody[];
}

interface CrearPedidoApi {
  success: boolean;
  salesOrderNumber: string;
}

interface CrearLineasApi {
  success: boolean;
  salesOrderNumber: string;
  lineasCreadas: number;
  mensaje: string;
}

interface HealthConexionApi {
  success: boolean;
  message: string;
}

interface TablaLeida {
  headers: string[];
  rows: ValorCelda[][];
}

const TABLA_PEDIDO = "tblPedido";
const TABLA_LINEAS = "tblLineas";
const TABLA_RESULTADOS = "tblResultados";
const TABLA_HISTORIAL = "tblHistorial";

const COL_PEDIDO_CLIENTE = "Cliente";
const COL_PEDIDO_REFERENCIA = "Referencia de cliente";
const COL_PEDIDO_DESCRIPCION = "Descripción de pedido";
const COL_PEDIDO_FECHA_ENVIO = "Fecha de envio solicitada";
const COL_PEDIDO_FECHA_RECEPCION = "Fecha de recepción solicitada";

const COL_LINEAS_REF_PEDIDO = "Referencia de cliente";
const COL_LINEAS_OV = "OrdenVenta";
const COL_LINEAS_CODIGO = "Codigo_Articulo";
const COL_LINEAS_CANTIDAD = "Cantidad";
const COL_LINEAS_PRECIO = "PrecioUnitario";
const COL_LINEAS_DESCUENTO = "Porcentaje de descuento";
const COL_LINEAS_FECHA_ENVIO = "Fecha de envio";
const COL_LINEAS_COMENTARIO = "Comentario";

const COL_RESULTADO_ESTADO = "Estado";
const COL_RESULTADO_FECHA = "Fecha de ejecución";
const COL_RESULTADO_USUARIO = "Usuario";
const COL_RESULTADO_ERROR = "Error";

const HOJA_RESULTADO = "Resultado";
const CELDA_API_BASE = "B1";

const MSG_CREAR_PEDIDO_PRIMERO = "Debe crear primero el pedido.";

async function probarConexion(workbook: ExcelScript.Workbook): Promise<void> {
  const ov = leerPedidoDynamicsDesdeResultados(workbook);
  try {
    const base = leerApiBaseUrl(workbook);

    const ping = (await llamarApi(
      base + "/api/health/ping",
      "GET",
      null
    )) as HealthConexionApi;

    if (!ping.success) {
      escribirResultado(
        workbook,
        "ERROR",
        ping.message || "El servidor no respondió al ping.",
        ov
      );
      return;
    }

    const respuesta = (await llamarApi(
      base + "/api/health/dynamics",
      "GET",
      null
    )) as HealthConexionApi;

    if (respuesta.success) {
      escribirResultado(
        workbook,
        "CONEXION OK",
        respuesta.message || "Conexion OK",
        ov
      );
      return;
    }

    escribirResultado(
      workbook,
      "ERROR",
      "Servidor OK, pero Dynamics/Azure falló: " +
        (respuesta.message || "Fallo de conexion"),
      ov
    );
  } catch (error) {
    escribirResultado(workbook, "ERROR", obtenerMensajeError(error), ov);
  }
}

async function crearPedido(workbook: ExcelScript.Workbook): Promise<void> {
  try {
    const tablaPedido = workbook.getTable(TABLA_PEDIDO);
    if (!tablaPedido) {
      escribirResultado(workbook, "ERROR", "No se encontró la tabla '" + TABLA_PEDIDO + "'.", "");
      return;
    }

    const pedidoData = leerTabla(tablaPedido);
    const errorPedido = validarPedido(pedidoData);
    if (errorPedido !== null) {
      escribirResultado(workbook, "ERROR", errorPedido, "");
      return;
    }

    const body = construirCrearPedidoBody(pedidoData);
    const base = leerApiBaseUrl(workbook);
    const respuesta = (await llamarApi(
      base + "/api/pedidos/crear",
      "POST",
      JSON.stringify(body)
    )) as CrearPedidoApi;

    sincronizarLineasConCabecera(
      workbook,
      body.referenciaCliente,
      respuesta.salesOrderNumber
    );

    escribirResultado(
      workbook,
      "CABECERA OK",
      "Cabecera creada. Referencia: " + body.referenciaCliente,
      respuesta.salesOrderNumber
    );
  } catch (error) {
    escribirResultado(workbook, "ERROR", obtenerMensajeError(error), "");
  }
}

function sincronizarLineasConCabecera(
  workbook: ExcelScript.Workbook,
  referenciaCliente: string,
  ordenVentaDynamics: string
): void {
  const tl = workbook.getTable(TABLA_LINEAS);
  if (!tl) {
    return;
  }
  if (leerTabla(tl).rows.length === 0) {
    return;
  }

  if (referenciaCliente !== "") {
    asegurarColumnaConValor(tl, COL_LINEAS_REF_PEDIDO, referenciaCliente);
  }
  if (ordenVentaDynamics !== "") {
    asegurarColumnaConValor(tl, COL_LINEAS_OV, ordenVentaDynamics);
  }
}

function asegurarColumnaConValor(
  tabla: ExcelScript.Table,
  nombreColumna: string,
  valor: string
): void {
  if (valor === "") {
    return;
  }

  let datos = leerTabla(tabla);
  if (datos.rows.length === 0) {
    return;
  }

  let colIdx = indiceColumna(datos.headers, nombreColumna);
  if (colIdx < 0) {
    tabla.addColumn(-1, null, nombreColumna);
    datos = leerTabla(tabla);
    colIdx = indiceColumna(datos.headers, nombreColumna);
  }
  if (colIdx < 0) {
    return;
  }

  const cuerpo = tabla.getRangeBetweenHeaderAndTotal();
  const allValues = cuerpo.getValues() as (string | number | boolean)[][];
  let cambio = false;

  for (let r = 0; r < allValues.length; r++) {
    const act = aTexto(allValues[r][colIdx] as ValorCelda);
    if (act === "") {
      allValues[r][colIdx] = valor;
      cambio = true;
    }
  }

  if (cambio) {
    cuerpo.setValues(allValues);
  }
}

async function crearLineas(workbook: ExcelScript.Workbook): Promise<void> {
  try {
    const salesOrderNumber = leerPedidoDynamicsDesdeResultados(workbook);
    if (salesOrderNumber === "") {
      escribirResultado(workbook, "ERROR", MSG_CREAR_PEDIDO_PRIMERO, "");
      return;
    }

    const tablaLineas = workbook.getTable(TABLA_LINEAS);
    if (!tablaLineas) {
      escribirResultado(workbook, "ERROR", "No se encontró la tabla '" + TABLA_LINEAS + "'.", salesOrderNumber);
      return;
    }

    const lineasData = leerTabla(tablaLineas);
    const errorLineas = validarLineas(lineasData);
    if (errorLineas !== null) {
      escribirResultado(workbook, "ERROR", errorLineas, salesOrderNumber);
      return;
    }

    const body: CrearLineasBody = {
      salesOrderNumber: salesOrderNumber,
      lineas: construirLineasBody(lineasData),
    };

    const base = leerApiBaseUrl(workbook);
    const respuesta = (await llamarApi(
      base + "/api/pedidos/lineas",
      "POST",
      JSON.stringify(body)
    )) as CrearLineasApi;

    const tablaPedido = workbook.getTable(TABLA_PEDIDO);
    const pedidoData = tablaPedido ? leerTabla(tablaPedido) : { headers: [], rows: [] };
    const lineBodies = construirLineasBody(lineasData);
    if (workbook.getTable(TABLA_HISTORIAL) && pedidoData.rows.length > 0) {
      registrarHistorialLineas(
        workbook,
        new Date(),
        respuesta.salesOrderNumber,
        pedidoData,
        lineBodies
      );
    }

    escribirResultado(
      workbook,
      "LINEAS OK",
      respuesta.mensaje,
      respuesta.salesOrderNumber
    );
  } catch (error) {
    const ov = leerPedidoDynamicsDesdeResultados(workbook);
    escribirResultado(workbook, "ERROR", obtenerMensajeError(error), ov);
  }
}

function registrarHistorialLineas(
  workbook: ExcelScript.Workbook,
  fechaHora: Date,
  ov: string,
  pedidoSnapshot: TablaLeida,
  lineas: LineaBody[]
): void {
  const tabla = workbook.getTable(TABLA_HISTORIAL);
  if (!tabla || pedidoSnapshot.rows.length === 0) {
    return;
  }

  const h = pedidoSnapshot.headers;
  const filaP = pedidoSnapshot.rows[0];
  const ref = aTexto(valorCelda(h, filaP, COL_PEDIDO_REFERENCIA));
  const cli = aTexto(valorCelda(h, filaP, COL_PEDIDO_CLIENTE));
  const desc = aTexto(valorCelda(h, filaP, COL_PEDIDO_DESCRIPCION));

  for (let i = 0; i < lineas.length; i++) {
    const ln = lineas[i];
    agregarFilaHistorial(tabla, fechaHora, ov, ref, cli, desc, ln);
  }
}

function agregarFilaHistorial(
  tabla: ExcelScript.Table,
  fechaHora: Date,
  ov: string,
  ref: string,
  cli: string,
  desc: string,
  ln: LineaBody
): void {
  const datos = leerTabla(tabla);
  const headers = datos.headers;

  tabla.addRow();
  const cuerpo = tabla.getRangeBetweenHeaderAndTotal();
  const r = cuerpo.getRowCount() - 1;

  for (let c = 0; c < headers.length; c++) {
    const nombre = headers[c];
    const celda = cuerpo.getCell(r, c);

    if (nombre === "FechaHora") {
      celda.setValue(fechaHora);
    } else if (nombre === "OrdenVenta") {
      celda.setValue(ov);
    } else if (nombre === "ReferenciaCliente") {
      celda.setValue(ref);
    } else if (nombre === "Cliente") {
      celda.setValue(cli);
    } else if (nombre === "DescripcionPedido") {
      celda.setValue(desc);
    } else if (nombre === "Codigo_Articulo") {
      celda.setValue(ln.codigoArticulo);
    } else if (nombre === "Cantidad") {
      celda.setValue(ln.cantidad);
    } else if (nombre === "PrecioUnitario") {
      celda.setValue(ln.precioUnitario);
    } else if (nombre === "Porcentaje de descuento") {
      celda.setValue(ln.porcentajeDescuento);
    } else if (nombre === "Fecha de envio") {
      celda.setValue(ln.fechaEnvio);
    } else if (nombre === "Comentario") {
      celda.setValue(ln.comentario);
    } else {
      celda.setValue("");
    }
  }
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

function construirCrearPedidoBody(pedido: TablaLeida): CrearPedidoBody {
  const fila = pedido.rows[0];
  const h = pedido.headers;

  return {
    cliente: aTexto(valorCelda(h, fila, COL_PEDIDO_CLIENTE)),
    referenciaCliente: aTexto(valorCelda(h, fila, COL_PEDIDO_REFERENCIA)),
    descripcionPedido: aTexto(valorCelda(h, fila, COL_PEDIDO_DESCRIPCION)),
    fechaEnvioSolicitada: convertirFecha(valorCelda(h, fila, COL_PEDIDO_FECHA_ENVIO)),
    fechaRecepcionSolicitada: convertirFecha(valorCelda(h, fila, COL_PEDIDO_FECHA_RECEPCION)),
  };
}

function construirLineasBody(lineas: TablaLeida): LineaBody[] {
  const hl = lineas.headers;
  const resultado: LineaBody[] = [];

  for (let i = 0; i < lineas.rows.length; i++) {
    const fila = lineas.rows[i];
    resultado.push({
      codigoArticulo: aTexto(valorCelda(hl, fila, COL_LINEAS_CODIGO)),
      cantidad: aNumero(valorCelda(hl, fila, COL_LINEAS_CANTIDAD)),
      precioUnitario: aNumero(valorCelda(hl, fila, COL_LINEAS_PRECIO)),
      porcentajeDescuento: aPorcentaje(valorCelda(hl, fila, COL_LINEAS_DESCUENTO)),
      fechaEnvio: convertirFecha(valorCelda(hl, fila, COL_LINEAS_FECHA_ENVIO)),
      comentario: aTexto(
        valorCelda(hl, fila, COL_LINEAS_COMENTARIO) ??
          valorCelda(hl, fila, "Comentarios")
      ),
    });
  }

  return resultado;
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

function headersApi(): Record<string, string> {
  return {
    "Content-Type": "application/json",
    "ngrok-skip-browser-warning": "true",
  };
}

async function llamarApi(
  url: string,
  metodo: string,
  cuerpo: string | null
): Promise<object> {
  const headers = headersApi();
  const response =
    cuerpo !== null && cuerpo !== ""
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

  try {
    return JSON.parse(textoRespuesta) as object;
  } catch (parseError) {
    const inicio = textoRespuesta.substring(0, 120).replace(/\s+/g, " ");
    if (
      textoRespuesta.indexOf("<!DOCTYPE") >= 0 ||
      textoRespuesta.indexOf("<html") >= 0
    ) {
      throw new Error(
        "La URL respondió con una página web (HTML), no con la API. " +
          "Revisa Resultado!B1 (URL https de ngrok actual), que Spring Boot y ngrok estén corriendo, " +
          "y vuelve a pegar el script ProbarConexion en Excel."
      );
    }
    throw new Error("Respuesta no es JSON: " + inicio);
  }
}

function obtenerMensajeError(error: string | number | boolean | object): string {
  if (typeof error === "object" && error !== null) {
    const errObj = error as { message?: string };
    if (errObj.message) {
      return String(errObj.message);
    }
  }
  return String(error);
}

function leerTabla(tabla: ExcelScript.Table): TablaLeida {
  const headerRange = tabla.getHeaderRowRange();
  const headerValues = headerRange.getValues();
  const headers = headerValues[0].map((celda) => String(celda).trim());

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

function valorCelda(
  headers: string[],
  fila: ValorCelda[],
  nombreColumna: string
): ValorCelda | undefined {
  const indice = indiceColumna(headers, nombreColumna);
  if (indice < 0) {
    return undefined;
  }
  return fila[indice];
}

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

function aNumero(valor: ValorCelda | undefined): number {
  if (esVacio(valor)) {
    return 0;
  }
  const numero = Number(valor);
  return Number.isFinite(numero) ? numero : NaN;
}

function aPorcentaje(valor: ValorCelda | undefined): number {
  const numero = aNumero(valor);
  if (Number.isNaN(numero)) {
    return 0;
  }
  if (numero <= 1) {
    return numero * 100;
  }
  return numero;
}

function convertirFecha(valor: ValorCelda | undefined): string {
  if (esVacio(valor)) {
    return "";
  }

  if (typeof valor === "number") {
    return serialExcelAFecha(valor);
  }

  if (typeof valor === "string") {
    const texto = valor.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) {
      return texto;
    }
    const fechaDmY = parsearFechaDdMmYyyy(texto);
    if (fechaDmY !== "") {
      return fechaDmY;
    }
    return "";
  }

  return "";
}

function parsearFechaDdMmYyyy(texto: string): string {
  const partes = texto.split("/");
  if (partes.length !== 3) {
    return "";
  }

  const dia = Number(partes[0]);
  const mes = Number(partes[1]);
  const anio = Number(partes[2]);

  if (
    !Number.isFinite(dia) ||
    !Number.isFinite(mes) ||
    !Number.isFinite(anio) ||
    anio < 1000 ||
    mes < 1 ||
    mes > 12 ||
    dia < 1 ||
    dia > 31
  ) {
    return "";
  }

  return (
    String(anio) +
    "-" +
    rellenarDosDigitos(mes) +
    "-" +
    rellenarDosDigitos(dia)
  );
}

function serialExcelAFecha(serial: number): string {
  const fecha = new Date((serial - 25569) * 86400 * 1000);
  const yyyy = fecha.getFullYear();
  const mm = rellenarDosDigitos(fecha.getMonth() + 1);
  const dd = rellenarDosDigitos(fecha.getDate());
  return yyyy + "-" + mm + "-" + dd;
}

function rellenarDosDigitos(valor: number): string {
  return valor < 10 ? "0" + String(valor) : String(valor);
}

function validarPedido(pedido: TablaLeida): string | null {
  if (pedido.rows.length === 0) {
    return "El pedido debe contener exactamente una fila de captura.";
  }
  if (pedido.rows.length > 1) {
    return "tblPedido solo debe tener una fila de datos.";
  }

  const fila = pedido.rows[0];
  const headers = pedido.headers;

  if (indiceColumna(headers, COL_PEDIDO_CLIENTE) < 0) {
    return "Falta la columna '" + COL_PEDIDO_CLIENTE + "' en tblPedido.";
  }
  if (indiceColumna(headers, COL_PEDIDO_REFERENCIA) < 0) {
    return "Falta la columna '" + COL_PEDIDO_REFERENCIA + "' en tblPedido.";
  }
  if (indiceColumna(headers, COL_PEDIDO_DESCRIPCION) < 0) {
    return "Falta la columna '" + COL_PEDIDO_DESCRIPCION + "' en tblPedido.";
  }

  if (esVacio(valorCelda(headers, fila, COL_PEDIDO_CLIENTE))) {
    return "Cliente es obligatorio.";
  }
  if (esVacio(valorCelda(headers, fila, COL_PEDIDO_REFERENCIA))) {
    return "Referencia de cliente es obligatoria.";
  }
  if (esVacio(valorCelda(headers, fila, COL_PEDIDO_DESCRIPCION))) {
    return "Descripción de pedido es obligatoria.";
  }

  return null;
}

function validarLineas(lineas: TablaLeida): string | null {
  if (lineas.rows.length === 0) {
    return "Debe existir al menos una línea en tblLineas.";
  }

  const headers = lineas.headers;

  if (indiceColumna(headers, COL_LINEAS_CODIGO) < 0) {
    return "Falta la columna '" + COL_LINEAS_CODIGO + "' en tblLineas.";
  }
  if (indiceColumna(headers, COL_LINEAS_CANTIDAD) < 0) {
    return "Falta la columna '" + COL_LINEAS_CANTIDAD + "' en tblLineas.";
  }

  for (let i = 0; i < lineas.rows.length; i++) {
    const fila = lineas.rows[i];
    const numeroFila = i + 1;

    if (esVacio(valorCelda(headers, fila, COL_LINEAS_CODIGO))) {
      return "Línea " + numeroFila + ": Codigo_Articulo es obligatorio.";
    }

    const cantidad = aNumero(valorCelda(headers, fila, COL_LINEAS_CANTIDAD));
    if (Number.isNaN(cantidad) || cantidad <= 0) {
      return "Línea " + numeroFila + ": Cantidad debe ser mayor que 0.";
    }
  }

  return null;
}

function escribirResultado(
  workbook: ExcelScript.Workbook,
  estado: string,
  mensajeError: string,
  pedidoDynamics: string
): void {
  const tabla = workbook.getTable(TABLA_RESULTADOS);
  if (!tabla) {
    throw new Error(
      "No existe la tabla '" +
        TABLA_RESULTADOS +
        "'. Debe estar en la hoja Resultado con columnas Estado, Error, etc."
    );
  }

  if (tabla.getRowCount() === 0) {
    tabla.addRow();
  }

  const datos = leerTabla(tabla);
  const headers = datos.headers;

  const idxEstado = indiceColumna(headers, COL_RESULTADO_ESTADO);
  const idxPedido = indiceColumnaPedidoDynamics(headers);
  const idxFecha = indiceColumna(headers, COL_RESULTADO_FECHA);
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

async function main(workbook: ExcelScript.Workbook): Promise<void> {
  await crearLineas(workbook);
}
