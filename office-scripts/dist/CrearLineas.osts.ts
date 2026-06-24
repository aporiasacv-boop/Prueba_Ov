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
const HOJA_HISTORIAL = "Historial";
const CELDA_API_BASE = "B1";
const SCRIPT_LINEAS_MARKETING_VERSION = "2026-06-09-lineas-hist";

const HISTORIAL_COLUMNAS = [
  "FechaHora",
  "OrdenVenta",
  "ReferenciaCliente",
  "Cliente",
  "DescripcionPedido",
  "Codigo_Articulo",
  "Cantidad",
  "PrecioUnitario",
  "Porcentaje de descuento",
  "Fecha de envio",
  "Comentario",
];

const MSG_CREAR_PEDIDO_PRIMERO = "Debe crear primero el pedido.";

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
      respuesta.message || "Fallo de conexion",
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

async function subirLote(workbook: ExcelScript.Workbook): Promise<void> {
  try {
    const tablaPedido = workbook.getTable(TABLA_PEDIDO);
    const tablaLineas = workbook.getTable(TABLA_LINEAS);
    const tablaHist = workbook.getTable(TABLA_HISTORIAL);

    if (!tablaPedido || !tablaLineas) {
      escribirResultado(workbook, "ERROR", "Faltan tblPedido o tblLineas.", "");
      return;
    }
    if (!tablaHist) {
      escribirResultado(
        workbook,
        "ERROR",
        "Cree la hoja Historial con la tabla tblHistorial (ver OPERACION.md).",
        ""
      );
      return;
    }

    const pedidoData = leerTabla(tablaPedido);
    const lineasData = leerTabla(tablaLineas);
    const errLote = validarLote(pedidoData, lineasData);
    if (errLote !== null) {
      escribirResultado(workbook, "ERROR", errLote, "");
      return;
    }

    const base = leerApiBaseUrl(workbook);
    const ahora = new Date();
    let totalLineas = 0;
    const unSoloPedido = pedidoData.rows.length === 1;
    const refUnica = aTexto(
      valorCelda(pedidoData.headers, pedidoData.rows[0], COL_PEDIDO_REFERENCIA)
    );

    for (let r = 0; r < pedidoData.rows.length; r++) {
      const bodyP = construirCrearPedidoBodyFila(pedidoData, r);
      const respP = (await llamarApi(
        base + "/api/pedidos/crear",
        "POST",
        JSON.stringify(bodyP)
      )) as CrearPedidoApi;

      const ov = respP.salesOrderNumber;
      const ref = bodyP.referenciaCliente;

      const subLineas = filtrarLineasPorReferencia(lineasData, ref, unSoloPedido, refUnica);
      if (subLineas.rows.length === 0) {
        escribirResultado(
          workbook,
          "ERROR",
          "Sin lineas para la referencia: " + ref,
          ov
        );
        return;
      }

      const errL = validarLineas(subLineas);
      if (errL !== null) {
        escribirResultado(workbook, "ERROR", errL + " Ref: " + ref, ov);
        return;
      }

      const lineBodies = construirLineasBody(subLineas);
      const bodyL: CrearLineasBody = {
        salesOrderNumber: ov,
        lineas: lineBodies,
      };

      const respL = (await llamarApi(
        base + "/api/pedidos/lineas",
        "POST",
        JSON.stringify(bodyL)
      )) as CrearLineasApi;

      totalLineas += respL.lineasCreadas;

      const pedidoUnaFila: TablaLeida = {
        headers: pedidoData.headers,
        rows: [pedidoData.rows[r]],
      };
      registrarHistorialLineas(workbook, ahora, ov, pedidoUnaFila, lineBodies);
    }

    limpiarTablaDatos(tablaPedido);
    limpiarTablaDatos(tablaLineas);

    escribirResultado(
      workbook,
      "LOTE OK",
      "Pedidos: " +
        pedidoData.rows.length +
        ". Lineas: " +
        totalLineas +
        ". Captura vaciada.",
      ""
    );
  } catch (error) {
    escribirResultado(
      workbook,
      "ERROR",
      obtenerMensajeError(error),
      leerPedidoDynamicsDesdeResultados(workbook)
    );
  }
}

function asegurarTablaHistorial(workbook: ExcelScript.Workbook): ExcelScript.Table | null {
  const existente = workbook.getTable(TABLA_HISTORIAL);
  if (existente) {
    return existente;
  }

  let hoja = workbook.getWorksheet(HOJA_HISTORIAL);
  if (!hoja) {
    hoja = workbook.addWorksheet(HOJA_HISTORIAL);
  }

  const ultimaCol = "K";
  const encabezado = hoja.getRange("A1:" + ultimaCol + "1");
  const primera = aTexto(hoja.getRange("A1").getValue() as ValorCelda);
  if (primera === "") {
    encabezado.setValues([HISTORIAL_COLUMNAS]);
  }

  const filasUsadas = hoja.getUsedRange();
  let direccionTabla = "A1:" + ultimaCol + "1";
  if (filasUsadas && filasUsadas.getRowCount() > 1) {
    const filas = filasUsadas.getRowCount();
    direccionTabla = "A1:" + ultimaCol + filas;
  }

  const tabla = hoja.addTable(direccionTabla, true);
  tabla.setName(TABLA_HISTORIAL);
  return tabla;
}

function escribirCeldaHistorial(
  cuerpo: ExcelScript.Range,
  fila: number,
  headers: string[],
  nombres: string[],
  valor: ValorCelda | Date
): void {
  const idx = indiceColumnaFlexible(headers, nombres);
  if (idx >= 0) {
    cuerpo.getCell(fila, idx).setValue(valor as string | number | boolean | Date);
  }
}

function registrarHistorialLineas(
  workbook: ExcelScript.Workbook,
  fechaHora: Date,
  ov: string,
  pedidoSnapshot: TablaLeida,
  lineas: LineaBody[],
  tablaHistorial: ExcelScript.Table | null = null
): number {
  const tabla = tablaHistorial ?? asegurarTablaHistorial(workbook);
  if (!tabla || pedidoSnapshot.rows.length === 0 || lineas.length === 0) {
    return 0;
  }

  const h = pedidoSnapshot.headers;
  const filaP = pedidoSnapshot.rows[0];
  const ref = aTexto(valorCelda(h, filaP, COL_PEDIDO_REFERENCIA));
  const cli = aTexto(valorCelda(h, filaP, COL_PEDIDO_CLIENTE));
  const desc = aTexto(valorCelda(h, filaP, COL_PEDIDO_DESCRIPCION));

  for (let i = 0; i < lineas.length; i++) {
    agregarFilaHistorial(tabla, fechaHora, ov, ref, cli, desc, lineas[i]);
  }

  return lineas.length;
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

  escribirCeldaHistorial(cuerpo, r, headers, ["FechaHora"], fechaHora);
  escribirCeldaHistorial(cuerpo, r, headers, ["OrdenVenta", "Orden Venta"], ov);
  escribirCeldaHistorial(
    cuerpo,
    r,
    headers,
    ["ReferenciaCliente", "Referencia Cliente"],
    ref
  );
  escribirCeldaHistorial(cuerpo, r, headers, ["Cliente"], cli);
  escribirCeldaHistorial(
    cuerpo,
    r,
    headers,
    ["DescripcionPedido", "Descripcion Pedido", "Descripción pedido"],
    desc
  );
  escribirCeldaHistorial(
    cuerpo,
    r,
    headers,
    ["Codigo_Articulo", "Codigo Articulo", "Dynamics"],
    ln.codigoArticulo
  );
  escribirCeldaHistorial(cuerpo, r, headers, ["Cantidad"], ln.cantidad);
  escribirCeldaHistorial(
    cuerpo,
    r,
    headers,
    ["PrecioUnitario", "Precio Unitario", "Costo"],
    ln.precioUnitario
  );
  escribirCeldaHistorial(
    cuerpo,
    r,
    headers,
    ["Porcentaje de descuento", "Porcentaje descuento"],
    ln.porcentajeDescuento
  );
  escribirCeldaHistorial(
    cuerpo,
    r,
    headers,
    ["Fecha de envio", "Fecha envio", "Fecha", "Fecha de entrega"],
    ln.fechaEnvio
  );
  escribirCeldaHistorial(cuerpo, r, headers, ["Comentario"], ln.comentario);
}

function limpiarTablaDatos(tabla: ExcelScript.Table): void {
  while (tabla.getRowCount() > 0) {
    tabla.deleteRowsAt(0, 1);
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
  return construirCrearPedidoBodyFila(pedido, 0);
}

function construirCrearPedidoBodyFila(pedido: TablaLeida, rowIndex: number): CrearPedidoBody {
  const fila = pedido.rows[rowIndex];
  const h = pedido.headers;

  return {
    cliente: aTexto(valorCelda(h, fila, COL_PEDIDO_CLIENTE)),
    referenciaCliente: aTexto(valorCelda(h, fila, COL_PEDIDO_REFERENCIA)),
    descripcionPedido: aTexto(valorCelda(h, fila, COL_PEDIDO_DESCRIPCION)),
    fechaEnvioSolicitada: convertirFecha(valorCelda(h, fila, COL_PEDIDO_FECHA_ENVIO)),
    fechaRecepcionSolicitada: convertirFecha(valorCelda(h, fila, COL_PEDIDO_FECHA_RECEPCION)),
  };
}

function filtrarLineasPorReferencia(
  todas: TablaLeida,
  refPedido: string,
  unSoloPedido: boolean,
  refUnica: string
): TablaLeida {
  const hl = todas.headers;
  const idxRef = indiceColumna(hl, COL_LINEAS_REF_PEDIDO);
  const rows: ValorCelda[][] = [];

  for (let i = 0; i < todas.rows.length; i++) {
    const fila = todas.rows[i];
    let refLinea = "";
    if (idxRef >= 0) {
      refLinea = aTexto(valorCelda(hl, fila, COL_LINEAS_REF_PEDIDO));
      if (unSoloPedido && refLinea === "") {
        refLinea = refUnica;
      }
    }
    if (unSoloPedido && idxRef < 0) {
      refLinea = refUnica;
    }
    if (refLinea === refPedido) {
      rows.push(fila);
    }
  }

  return { headers: hl, rows: rows };
}

function validarLote(pedido: TablaLeida, lineas: TablaLeida): string | null {
  if (pedido.rows.length === 0) {
    return "tblPedido debe tener al menos un pedido.";
  }

  const h = pedido.headers;
  if (indiceColumna(h, COL_PEDIDO_CLIENTE) < 0) {
    return "Falta columna Cliente en tblPedido.";
  }
  if (indiceColumna(h, COL_PEDIDO_REFERENCIA) < 0) {
    return "Falta columna Referencia de cliente en tblPedido.";
  }
  if (indiceColumna(h, COL_PEDIDO_DESCRIPCION) < 0) {
    return "Falta columna Descripción de pedido en tblPedido.";
  }

  const refsPedido: string[] = [];
  for (let i = 0; i < pedido.rows.length; i++) {
    const fila = pedido.rows[i];
    const n = i + 1;
    if (esVacio(valorCelda(h, fila, COL_PEDIDO_CLIENTE))) {
      return "Pedido fila " + n + ": Cliente obligatorio.";
    }
    if (esVacio(valorCelda(h, fila, COL_PEDIDO_REFERENCIA))) {
      return "Pedido fila " + n + ": Referencia obligatoria.";
    }
    if (esVacio(valorCelda(h, fila, COL_PEDIDO_DESCRIPCION))) {
      return "Pedido fila " + n + ": Descripción obligatoria.";
    }
    const ref = aTexto(valorCelda(h, fila, COL_PEDIDO_REFERENCIA));
    if (refsPedido.indexOf(ref) >= 0) {
      return "Referencia duplicada en tblPedido: " + ref;
    }
    refsPedido.push(ref);
  }

  const idxRefLinea = indiceColumna(lineas.headers, COL_LINEAS_REF_PEDIDO);
  if (pedido.rows.length > 1 && idxRefLinea < 0) {
    return "Con varios pedidos, agregue en tblLineas la columna Referencia de cliente (misma que en tblPedido).";
  }

  const errLineasBase = validarLineas(lineas);
  if (errLineasBase !== null) {
    return errLineasBase;
  }

  for (let i = 0; i < pedido.rows.length; i++) {
    const ref = aTexto(valorCelda(h, pedido.rows[i], COL_PEDIDO_REFERENCIA));
    const sub = filtrarLineasPorReferencia(
      lineas,
      ref,
      pedido.rows.length === 1,
      refsPedido[0]
    );
    if (sub.rows.length === 0) {
      return "Ninguna linea para el pedido con referencia: " + ref;
    }
  }

  for (let i = 0; i < lineas.rows.length; i++) {
    const fila = lineas.rows[i];
    const n = i + 1;
    let refL = "";
    if (idxRefLinea >= 0) {
      refL = aTexto(valorCelda(lineas.headers, fila, COL_LINEAS_REF_PEDIDO));
      if (pedido.rows.length === 1 && refL === "") {
        refL = refsPedido[0];
      }
    } else if (pedido.rows.length === 1) {
      refL = refsPedido[0];
    }
    if (refsPedido.indexOf(refL) < 0) {
      return "Linea " + n + ": Referencia no coincide con ningun pedido: " + refL;
    }
  }

  return null;
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

function headersApi(metodo: string, conCuerpo: boolean): Record<string, string> {
  const headers: Record<string, string> = {
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
  const headers = headersApi(metodo, conCuerpo);
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
    const ovEnError = extraerOvDeTexto(textoRespuesta);
    if (ovEnError !== "" && textoRespuesta.indexOf("\"success\":true") >= 0) {
      return { success: true, salesOrderNumber: ovEnError };
    }
    throw new Error("HTTP " + response.status + ": " + textoRespuesta);
  }

  if (textoRespuesta === "") {
    return {};
  }

  return JSON.parse(textoRespuesta) as object;
}

function extraerOvDeTexto(texto: string): string {
  const m = texto.match(/OV\d{4,}/i);
  if (m) {
    return m[0].toUpperCase();
  }
  return "";
}

function extraerOvDeRespuesta(resp: object): string {
  const r = resp as Record<string, ValorCelda>;
  const ov =
    aTexto(r.salesOrderNumber) ||
    aTexto(r.SalesOrderNumber) ||
    extraerOvDeTexto(JSON.stringify(resp));
  return ov;
}

function extraerMensajeDynamics(texto: string): string {
  const lower = texto.toLowerCase();

  if (lower.indexOf("salesordernumber no viene") >= 0) {
    const ov = extraerOvDeTexto(texto);
    if (ov !== "") {
      return "Orden creada (" + ov + "). Si T3 esta vacia, pegue ese numero manualmente.";
    }
    return "La orden pudo crearse en Dynamics pero el servidor no devolvio el numero OV.";
  }

  if (
    lower.indexOf("accountnum") >= 0 &&
    (lower.indexOf("cliente (") >= 0 || lower.indexOf("custtableorderingcustomer") >= 0)
  ) {
    return (
      "V3 no es valida. Use solo la cuenta (ej. C0010), no el texto de la etiqueta."
    );
  }

  if (lower.indexOf("failed to fetch") >= 0) {
    return "No se pudo conectar con la API. Revise ngrok en Resultado!B1 y que el backend este activo.";
  }

  if (lower.indexOf("http 401") >= 0 || lower.indexOf("http 403") >= 0) {
    return "Sin autorizacion con Dynamics. Revise el token Azure en el backend.";
  }

  const ovEnTexto = extraerOvDeTexto(texto);
  if (ovEnTexto !== "" && lower.indexOf("\"success\":true") >= 0) {
    return "";
  }

  return texto.length > 300 ? texto.substring(0, 300) + "..." : texto;
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
  if (msg.indexOf("HTTP 4") >= 0 || msg.indexOf("HTTP 5") >= 0) {
    const claro = extraerMensajeDynamics(msg);
    return claro !== "" ? claro : msg;
  }
  return msg;
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

function normalizarEncabezado(texto: string): string {
  return texto.trim().toLowerCase();
}

function indiceColumnaFlexible(headers: string[], nombres: string[]): number {
  const norm = headers.map((h) => normalizarEncabezado(h));
  for (let i = 0; i < nombres.length; i++) {
    const idx = norm.indexOf(normalizarEncabezado(nombres[i]));
    if (idx >= 0) {
      return idx;
    }
  }
  return -1;
}

function valorCeldaFlexible(
  headers: string[],
  fila: ValorCelda[],
  nombres: string[]
): ValorCelda | undefined {
  const indice = indiceColumnaFlexible(headers, nombres);
  if (indice < 0) {
    return undefined;
  }
  return fila[indice];
}

function primerTextoNoVacio(
  headers: string[],
  fila: ValorCelda[],
  nombres: string[]
): string {
  for (let i = 0; i < nombres.length; i++) {
    const texto = aTexto(valorCeldaFlexible(headers, fila, [nombres[i]]));
    if (texto !== "") {
      return texto;
    }
  }
  return "";
}

function unirDescripcion(
  headers: string[],
  fila: ValorCelda[],
  nombres: string[]
): string {
  const partes: string[] = [];
  for (let i = 0; i < nombres.length; i++) {
    const texto = aTexto(valorCeldaFlexible(headers, fila, [nombres[i]]));
    if (texto !== "") {
      partes.push(texto);
    }
  }
  return partes.join(" - ");
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

function aCodigoArticulo(valor: ValorCelda | undefined): string {
  if (esVacio(valor)) {
    return "";
  }
  if (typeof valor === "number" && Number.isFinite(valor)) {
    return String(Math.trunc(valor));
  }
  const texto = aTexto(valor);
  if (/^\d+(\.\d+)?e\+\d+$/i.test(texto)) {
    const numero = Number(texto);
    if (Number.isFinite(numero)) {
      return String(Math.trunc(numero));
    }
  }
  return texto;
}

function aNumero(valor: ValorCelda | undefined): number {
  if (esVacio(valor)) {
    return 0;
  }
  if (typeof valor === "number") {
    return Number.isFinite(valor) ? valor : NaN;
  }
  const texto = String(valor).trim();
  if (texto === "") {
    return 0;
  }
  const limpio = texto.replace(/[^\d.,-]/g, "").replace(/,/g, "");
  const numero = Number(limpio);
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
  const wholeDays = Math.floor(serial);
  const ms = Date.UTC(1899, 11, 30) + wholeDays * 86400000;
  const fecha = new Date(ms);
  const yyyy = fecha.getUTCFullYear();
  const mm = rellenarDosDigitos(fecha.getUTCMonth() + 1);
  const dd = rellenarDosDigitos(fecha.getUTCDate());
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
    return;
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

interface FuenteMarketing {
  hoja: string;
  nombreTabla: string;
  etiqueta: string;
  celdaOvCabecera: string;
  celdaCuenta: string;
  celdaNombreCliente: string;
  celdaDescripcion: string;
  celdaReferencia: string;
  celdaFechaEnvio: string;
  celdaFechaRecepcion: string;
  colOv: string[];
  colsReferencia: string[];
}

interface FilasPendientesMarketing {
  headers: string[];
  filas: ValorCelda[][];
  indices: number[];
  idxOv: number;
}

const MSG_CABECERA_PRIMERO = "Debe crear primero la orden (celda T3 con la OV).";
const MSG_CUENTA_V3 =
  "Falta la cuenta cliente. Escriba el codigo de Dynamics en V3, U3 o W3 (ej. C001, C0010).";
const MSG_NOMBRE_V4 =
  "Falta el nombre del cliente en V4. Escriba el nombre comercial del cliente.";
const MSG_DESCRIPCION_V2 =
  "Falta la descripcion de la orden en V2. No se enviara sin ese dato.";
const MSG_CLIENTE_TABLA =
  "Indique la cuenta cliente en la columna Cliente (ej. C0010) en al menos una fila.";

const COL_MKT_CLIENTE = ["Cliente", "Cuenta Dynamics"];
const COL_MKT_DESCRIPCION = ["Descripción pedido", "Descripcion pedido"];
const COL_MKT_FECHA_ENVIO = ["Fecha"];
const COL_MKT_FECHA_RECEPCION = ["Fecha de compra"];
const COL_MKT_CODIGO = ["Dynamics"];
const COL_MKT_CANTIDAD = ["Cantidad"];
const COL_MKT_PRECIO = ["Costo", "Precio unitario", "PrecioUnitario"];

const FUENTES_MARKETING: FuenteMarketing[] = [
  {
    hoja: "Costco",
    nombreTabla: "Tabla3",
    etiqueta: "Costco",
    celdaOvCabecera: "T3",
    celdaCuenta: "V3",
    celdaNombreCliente: "V4",
    celdaDescripcion: "V2",
    celdaReferencia: "V6",
    celdaFechaEnvio: "V7",
    celdaFechaRecepcion: "V8",
    colOv: ["Ov"],
    colsReferencia: ["PO", "Orden"],
  },
  {
    hoja: "E-commerce",
    nombreTabla: "Tabla4",
    etiqueta: "E-commerce",
    celdaOvCabecera: "T3",
    celdaCuenta: "V3",
    celdaNombreCliente: "V4",
    celdaDescripcion: "V2",
    celdaReferencia: "V6",
    celdaFechaEnvio: "V7",
    celdaFechaRecepcion: "V8",
    colOv: ["Orden de venta"],
    colsReferencia: ["N° venta", "N venta", "Orden"],
  },
];

function fuenteMarketingActiva(workbook: ExcelScript.Workbook): FuenteMarketing | null {
  const activa = workbook.getActiveWorksheet().getName();
  for (let i = 0; i < FUENTES_MARKETING.length; i++) {
    if (FUENTES_MARKETING[i].hoja === activa) {
      return FUENTES_MARKETING[i];
    }
  }

  for (let i = 0; i < FUENTES_MARKETING.length; i++) {
    const fuente = FUENTES_MARKETING[i];
    const tabla = workbook.getTable(fuente.nombreTabla);
    if (tabla && tabla.getRowCount() > 0) {
      return fuente;
    }
  }

  return null;
}

function leerOvParaLineasMarketing(
  workbook: ExcelScript.Workbook,
  fuente: FuenteMarketing
): string {
  const ovPanel = leerOvCabeceraPanel(workbook, fuente);
  if (ovPanel !== "") {
    return ovPanel;
  }
  return leerPedidoDynamicsDesdeResultados(workbook);
}

function leerCeldaPanel(
  workbook: ExcelScript.Workbook,
  fuente: FuenteMarketing,
  celda: string
): string {
  const hoja = workbook.getWorksheet(fuente.hoja);
  if (!hoja) {
    return "";
  }
  return aTexto(hoja.getRange(celda).getValue() as ValorCelda);
}

function escribirCeldaPanel(
  workbook: ExcelScript.Workbook,
  fuente: FuenteMarketing,
  celda: string,
  valor: string
): void {
  const hoja = workbook.getWorksheet(fuente.hoja);
  if (!hoja) {
    return;
  }
  hoja.getRange(celda).setValue(valor);
}

function leerOvCabeceraPanel(
  workbook: ExcelScript.Workbook,
  fuente: FuenteMarketing
): string {
  return leerCeldaPanel(workbook, fuente, fuente.celdaOvCabecera);
}

function leerClienteFila(headers: string[], fila: ValorCelda[]): string {
  return primerTextoNoVacio(headers, fila, COL_MKT_CLIENTE);
}

function leerCabeceraDesdeFila(
  headers: string[],
  fila: ValorCelda[],
  colsReferencia: string[]
): CrearPedidoBody {
  const cliente = leerClienteFila(headers, fila);
  let referencia = primerTextoNoVacio(headers, fila, colsReferencia);
  if (referencia === "") {
    referencia = "MKT-" + cliente;
  }

  let descripcion = aTexto(
    valorCeldaFlexible(headers, fila, COL_MKT_DESCRIPCION)
  );
  if (descripcion === "") {
    descripcion = "Pedido " + referencia;
  }

  return {
    cliente: cliente,
    referenciaCliente: referencia,
    descripcionPedido: descripcion,
    fechaEnvioSolicitada: convertirFecha(
      valorCeldaFlexible(headers, fila, COL_MKT_FECHA_ENVIO)
    ),
    fechaRecepcionSolicitada: convertirFecha(
      valorCeldaFlexible(headers, fila, COL_MKT_FECHA_RECEPCION)
    ),
  };
}

function construirLineaDesdeFila(headers: string[], fila: ValorCelda[]): LineaBody | null {
  const codigo = aTexto(valorCeldaFlexible(headers, fila, COL_MKT_CODIGO));
  const cantidad = aNumero(valorCeldaFlexible(headers, fila, COL_MKT_CANTIDAD));
  if (codigo === "" || cantidad <= 0) {
    return null;
  }

  const precio = aNumero(valorCeldaFlexible(headers, fila, COL_MKT_PRECIO));

  return {
    codigoArticulo: codigo,
    cantidad: cantidad,
    precioUnitario: Number.isNaN(precio) ? 0 : precio,
    porcentajeDescuento: 0,
    fechaEnvio: convertirFecha(
      valorCeldaFlexible(headers, fila, COL_MKT_FECHA_ENVIO)
    ),
    comentario: "",
  };
}

interface DiagnosticoFilasMarketing {
  pendientes: FilasPendientesMarketing;
  sinArticulo: number;
  sinCantidad: number;
  totalDatos: number;
}

function diagnosticarFilasMarketing(
  workbook: ExcelScript.Workbook,
  fuente: FuenteMarketing
): DiagnosticoFilasMarketing | string {
  const tabla = workbook.getTable(fuente.nombreTabla);
  if (!tabla) {
    return "No se encontro la tabla " + fuente.nombreTabla + ".";
  }

  const datos = leerTabla(tabla);
  const idxOv = indiceColumnaFlexible(datos.headers, fuente.colOv);

  const filas: ValorCelda[][] = [];
  const indices: number[] = [];
  let sinArticulo = 0;
  let sinCantidad = 0;

  for (let r = 0; r < datos.rows.length; r++) {
    const fila = datos.rows[r];
    const codigo = aTexto(valorCeldaFlexible(datos.headers, fila, COL_MKT_CODIGO));
    const cantidad = aNumero(valorCeldaFlexible(datos.headers, fila, COL_MKT_CANTIDAD));

    if (codigo === "" && cantidad <= 0) {
      continue;
    }

    if (codigo === "") {
      sinArticulo++;
      continue;
    }
    if (cantidad <= 0) {
      sinCantidad++;
      continue;
    }

    filas.push(fila);
    indices.push(r);
  }

  return {
    pendientes: {
      headers: datos.headers,
      filas: filas,
      indices: indices,
      idxOv: idxOv,
    },
    sinArticulo: sinArticulo,
    sinCantidad: sinCantidad,
    totalDatos: datos.rows.length,
  };
}

function mensajeSinFilasPendientes(diag: DiagnosticoFilasMarketing): string {
  if (diag.pendientes.filas.length > 0) {
    return "";
  }

  const partes: string[] = [];
  partes.push("No hay filas con Dynamics y Cantidad en la tabla.");

  if (diag.sinArticulo > 0) {
    partes.push(diag.sinArticulo + " fila(s) sin codigo en Dynamics.");
  }
  if (diag.sinCantidad > 0) {
    partes.push(diag.sinCantidad + " fila(s) sin Cantidad.");
  }
  if (diag.sinArticulo === 0 && diag.sinCantidad === 0) {
    partes.push("Agregue al menos una fila con articulo y cantidad.");
  }

  return partes.join(" ");
}

function limpiarCapturaMarketing(
  workbook: ExcelScript.Workbook,
  fuente: FuenteMarketing
): void {
  const tabla = workbook.getTable(fuente.nombreTabla);
  if (tabla) {
    limpiarTablaDatos(tabla);
  }
  escribirCeldaPanel(workbook, fuente, fuente.celdaOvCabecera, "");
}

function extraerCuentaDynamics(texto: string): string {
  const t = aTexto(texto as ValorCelda);
  if (t === "") {
    return "";
  }

  const entreParentesis = t.match(/\(([A-Za-z0-9_-]+)\)/);
  if (entreParentesis) {
    return entreParentesis[1].toUpperCase();
  }

  const limpio = t.replace(/\s+/g, "");
  if (/^[A-Za-z][A-Za-z0-9_-]{1,19}$/.test(limpio)) {
    return limpio.toUpperCase();
  }

  return "";
}

function esEtiquetaPanelCuenta(texto: string): boolean {
  const lower = texto.toLowerCase().trim();
  if (lower === "") {
    return false;
  }
  return (
    (lower.indexOf("cliente") >= 0 || lower.indexOf("cuenta") >= 0) &&
    texto.indexOf("(") < 0 &&
    extraerCuentaDynamics(texto) === ""
  );
}

function leerCuentaPanelMarketing(
  workbook: ExcelScript.Workbook,
  fuente: FuenteMarketing
): string {
  const celdasBusqueda = ["V3", "U3", "W3"];

  for (let i = 0; i < celdasBusqueda.length; i++) {
    const raw = leerCeldaPanel(workbook, fuente, celdasBusqueda[i]);
    const cuenta = extraerCuentaDynamics(raw);
    if (cuenta !== "" && !esEtiquetaPanelCuenta(raw)) {
      return cuenta;
    }
  }

  return "";
}

function validarCuentaV3(cuenta: string): string | null {
  if (cuenta === "") {
    return MSG_CUENTA_V3;
  }
  if (!/^[A-Za-z][A-Za-z0-9_-]{1,19}$/.test(cuenta)) {
    return (
      "El codigo de cuenta no parece valido (" +
      cuenta +
      "). Use el formato de Dynamics (ej. C001, C0010)."
    );
  }
  return null;
}

function validarNombreV4(nombre: string): string | null {
  if (nombre === "") {
    return MSG_NOMBRE_V4;
  }
  if (/^C\d{3,}$/i.test(nombre)) {
    return "V4 debe ser el nombre del cliente. La cuenta FO va en V3 (ej. C0010).";
  }
  if (
    nombre.toLowerCase().indexOf("nombre cliente") >= 0 &&
    nombre.indexOf("(") >= 0
  ) {
    return "V4 debe tener el nombre del cliente, no el texto de la etiqueta del panel.";
  }
  return null;
}

function validarDescripcionV2(descripcion: string): string | null {
  if (descripcion === "") {
    return MSG_DESCRIPCION_V2;
  }
  if (
    descripcion.toLowerCase().indexOf("descripcion") >= 0 &&
    descripcion.indexOf("orden") >= 0 &&
    descripcion.indexOf("(") >= 0
  ) {
    return "V2 debe tener la descripcion de la orden, no el texto de la etiqueta del panel.";
  }
  return null;
}

function validarPanelMarketing(
  workbook: ExcelScript.Workbook,
  fuente: FuenteMarketing
): string | null {
  const descripcion = leerCeldaPanel(workbook, fuente, fuente.celdaDescripcion);
  const errDescripcion = validarDescripcionV2(descripcion);
  if (errDescripcion !== null) {
    return errDescripcion;
  }

  const cuenta = leerCuentaPanelMarketing(workbook, fuente);
  const errCuenta = validarCuentaV3(cuenta);
  if (errCuenta !== null) {
    return errCuenta;
  }

  const nombre = leerCeldaPanel(workbook, fuente, fuente.celdaNombreCliente);
  const errNombre = validarNombreV4(nombre);
  if (errNombre !== null) {
    return errNombre;
  }

  return null;
}

function leerCabeceraDesdePanel(
  workbook: ExcelScript.Workbook,
  fuente: FuenteMarketing
): CrearPedidoBody | string {
  const errPanel = validarPanelMarketing(workbook, fuente);
  if (errPanel !== null) {
    return errPanel;
  }

  const cuenta = leerCuentaPanelMarketing(workbook, fuente);
  const descripcion = leerCeldaPanel(workbook, fuente, fuente.celdaDescripcion);

  let referencia = leerCeldaPanel(workbook, fuente, fuente.celdaReferencia);
  if (referencia === "") {
    referencia = "MKT-" + fuente.etiqueta + "-" + cuenta;
  }

  const fechaEnvioRaw = leerCeldaPanel(workbook, fuente, fuente.celdaFechaEnvio);
  const fechaRecepcionRaw = leerCeldaPanel(
    workbook,
    fuente,
    fuente.celdaFechaRecepcion
  );

  return {
    cliente: cuenta,
    referenciaCliente: referencia,
    descripcionPedido: descripcion,
    fechaEnvioSolicitada:
      fechaEnvioRaw !== "" ? convertirFecha(fechaEnvioRaw as ValorCelda) : "",
    fechaRecepcionSolicitada:
      fechaRecepcionRaw !== ""
        ? convertirFecha(fechaRecepcionRaw as ValorCelda)
        : "",
  };
}

async function crearCabeceraMarketing(workbook: ExcelScript.Workbook): Promise<void> {
  const fuente = fuenteMarketingActiva(workbook);
  try {
    if (fuente === null) {
      escribirResultado(
        workbook,
        "ERROR",
        "Active la hoja Costco o E-commerce antes de crear la orden.",
        ""
      );
      return;
    }

    const bodyOError = leerCabeceraDesdePanel(workbook, fuente);
    if (typeof bodyOError === "string") {
      escribirResultado(workbook, "ERROR", bodyOError, "");
      return;
    }

    const base = leerApiBaseUrl(workbook);
    const respP = await llamarApi(
      base + "/api/pedidos/crear",
      "POST",
      JSON.stringify(bodyOError)
    );

    const ov = extraerOvDeRespuesta(respP);
    if (ov === "") {
      escribirResultado(
        workbook,
        "ERROR",
        "La API respondio pero no trajo el numero OV. Revise Dynamics o reinicie el backend.",
        ""
      );
      return;
    }

    escribirCeldaPanel(workbook, fuente, fuente.celdaOvCabecera, ov);

    escribirResultado(
      workbook,
      "CABECERA OK",
      "Orden creada. OV en T3: " + ov,
      ov
    );
  } catch (error) {
    const raw = String(error);
    const ovRescatada = extraerOvDeTexto(raw);
    const msg = obtenerMensajeError(error);
    if (ovRescatada !== "" && fuente !== null) {
      escribirCeldaPanel(workbook, fuente, fuente.celdaOvCabecera, ovRescatada);
      escribirResultado(
        workbook,
        "CABECERA OK",
        "Orden creada. OV en T3: " + ovRescatada,
        ovRescatada
      );
      return;
    }
    escribirResultado(workbook, "ERROR", msg !== "" ? msg : raw, ovRescatada);
  }
}

async function crearOrdenMarketing(workbook: ExcelScript.Workbook): Promise<void> {
  await crearCabeceraMarketing(workbook);
}

async function crearLineasMarketing(workbook: ExcelScript.Workbook): Promise<void> {
  const fuente = fuenteMarketingActiva(workbook);
  try {
    if (fuente === null) {
      escribirResultado(
        workbook,
        "ERROR",
        "Active la hoja Costco o E-commerce antes de crear lineas.",
        ""
      );
      return;
    }

    const ov = leerOvParaLineasMarketing(workbook, fuente);
    if (ov === "") {
      escribirResultado(
        workbook,
        "ERROR",
        MSG_CABECERA_PRIMERO + " Cree la orden primero o deje la OV en T3.",
        ""
      );
      return;
    }

    const diag = diagnosticarFilasMarketing(workbook, fuente);
    if (typeof diag === "string") {
      escribirResultado(workbook, "ERROR", diag, ov);
      return;
    }
    if (diag.pendientes.filas.length === 0) {
      escribirResultado(
        workbook,
        "ERROR",
        mensajeSinFilasPendientes(diag) +
          " [" +
          SCRIPT_LINEAS_MARKETING_VERSION +
          "]",
        ov
      );
      return;
    }

    const lineas: LineaBody[] = [];
    for (let i = 0; i < diag.pendientes.filas.length; i++) {
      const linea = construirLineaDesdeFila(
        diag.pendientes.headers,
        diag.pendientes.filas[i]
      );
      if (linea !== null) {
        lineas.push(linea);
      }
    }

    if (lineas.length === 0) {
      escribirResultado(
        workbook,
        "ERROR",
        "Ninguna fila tiene Dynamics y Cantidad validos.",
        ov
      );
      return;
    }

    const tablaHist = asegurarTablaHistorial(workbook);
    if (tablaHist === null) {
      escribirResultado(
        workbook,
        "ERROR",
        "No se pudo crear la hoja Historial con tblHistorial. La captura no se vaciara.",
        ov
      );
      return;
    }

    const filasHistAntes = tablaHist.getRowCount();

    const base = leerApiBaseUrl(workbook);
    const bodyL: CrearLineasBody = {
      salesOrderNumber: ov,
      lineas: lineas,
    };

    const respL = (await llamarApi(
      base + "/api/pedidos/lineas",
      "POST",
      JSON.stringify(bodyL)
    )) as CrearLineasApi;

    const cuenta = leerCuentaPanelMarketing(workbook, fuente);
    const nombre = leerCeldaPanel(workbook, fuente, fuente.celdaNombreCliente);
    const descripcion = leerCeldaPanel(workbook, fuente, fuente.celdaDescripcion);
    let referencia = leerCeldaPanel(workbook, fuente, fuente.celdaReferencia);
    if (referencia === "") {
      referencia = "MKT-" + ov;
    }

    const pedidoSnapshot: TablaLeida = {
      headers: [
        COL_PEDIDO_CLIENTE,
        COL_PEDIDO_REFERENCIA,
        COL_PEDIDO_DESCRIPCION,
      ],
      rows: [[cuenta !== "" ? cuenta : nombre, referencia, descripcion]],
    };

    const filasGuardadas = registrarHistorialLineas(
      workbook,
      new Date(),
      ov,
      pedidoSnapshot,
      lineas
    );

    const filasHistDespues = tablaHist.getRowCount();
    if (filasGuardadas === 0 || filasHistDespues <= filasHistAntes) {
      escribirResultado(
        workbook,
        "ERROR",
        "Las lineas se crearon en Dynamics (" +
          ov +
          ") pero no se guardaron en Historial. La tabla de captura NO se vacio. Revise la hoja Historial / tblHistorial.",
        ov
      );
      return;
    }

    limpiarCapturaMarketing(workbook, fuente);

    escribirResultado(
      workbook,
      "LINEAS OK",
      (respL.mensaje || "Lineas creadas en " + ov) +
        ". " +
        lineas.length +
        " linea(s) en Historial. Tabla vaciada. [" +
        SCRIPT_LINEAS_MARKETING_VERSION +
        "]",
      ov
    );
  } catch (error) {
    escribirResultado(
      workbook,
      "ERROR",
      obtenerMensajeError(error),
      fuente !== null ? leerOvCabeceraPanel(workbook, fuente) : ""
    );
  }
}

async function subirMarketing(workbook: ExcelScript.Workbook): Promise<void> {
  const fuente = fuenteMarketingActiva(workbook);
  if (fuente === null) {
    escribirResultado(
      workbook,
      "ERROR",
      "Active la hoja Costco o E-commerce.",
      ""
    );
    return;
  }

  const ovActual = leerOvCabeceraPanel(workbook, fuente);
  if (ovActual === "") {
    await crearCabeceraMarketing(workbook);
    const ovNueva = leerOvCabeceraPanel(workbook, fuente);
    if (ovNueva === "") {
      return;
    }
  }

  await crearLineasMarketing(workbook);
}

function escribirCeldaTabla(
  tabla: ExcelScript.Table,
  filaIndice: number,
  colIndice: number,
  valor: string
): void {
  const cuerpo = tabla.getRangeBetweenHeaderAndTotal();
  cuerpo.getCell(filaIndice, colIndice).setValue(valor);
}

const TABLA_CAPTURA_COMERCIAL = "tblCapturaComercial";
const HOJA_CAPTURA_COMERCIAL = "Captura";
const TABLA_HISTORIAL_COMERCIAL = "tbl_historial";
const NOMBRES_TABLA_HISTORIAL_COMERCIAL = ["tbl_historial", "tblHistorial"];
const SCRIPT_COMERCIAL_VERSION = "2026-06-19-comercial-v3";

const COL_COM_CLIENTE = "Cliente";
const COL_COM_CODIGO = ["Código", "Codigo", "Codigo_Articulo"];
const COL_COM_PIEZAS = ["Piezas", "Cantidad"];
const COL_COM_PRECIO = ["Precio", "PrecioUnitario"];
const COL_COM_OC = ["Órden de compra", "Orden de compra", "Orden de compra"];
const COL_COM_FECHA = ["Fecha de entrega", "Fecha de envio"];
const COL_COM_OV = ["OrdenVenta", "Orden Venta"];
const COL_COM_PRODUCTO = ["Producto"];

const MSG_CABECERA_COMERCIAL_PRIMERO =
  "Debe crear primero la cabecera (boton Generar cabecera).";

function tablaCapturaComercial(
  workbook: ExcelScript.Workbook
): ExcelScript.Table | null {
  const tabla = workbook.getTable(TABLA_CAPTURA_COMERCIAL);
  if (tabla) {
    return tabla;
  }
  const hoja = workbook.getWorksheet(HOJA_CAPTURA_COMERCIAL);
  if (!hoja) {
    return null;
  }
  const tablas = hoja.getTables();
  return tablas.length > 0 ? tablas[0] : null;
}

function resolverTablaHistorialComercial(
  workbook: ExcelScript.Workbook
): ExcelScript.Table | null {
  for (let i = 0; i < NOMBRES_TABLA_HISTORIAL_COMERCIAL.length; i++) {
    const tabla = workbook.getTable(NOMBRES_TABLA_HISTORIAL_COMERCIAL[i]);
    if (tabla) {
      return tabla;
    }
  }

  const hoja = workbook.getWorksheet(HOJA_HISTORIAL);
  if (hoja) {
    const tablas = hoja.getTables();
    if (tablas.length > 0) {
      return tablas[0];
    }
  }

  return null;
}

function asegurarTablaHistorialComercial(
  workbook: ExcelScript.Workbook
): ExcelScript.Table | null {
  const existente = resolverTablaHistorialComercial(workbook);
  if (existente) {
    return existente;
  }

  let hoja = workbook.getWorksheet(HOJA_HISTORIAL);
  if (!hoja) {
    hoja = workbook.addWorksheet(HOJA_HISTORIAL);
  }

  const ultimaCol = "K";
  const encabezado = hoja.getRange("A1:" + ultimaCol + "1");
  const primera = aTexto(hoja.getRange("A1").getValue() as ValorCelda);
  if (primera === "") {
    encabezado.setValues([HISTORIAL_COLUMNAS]);
  }

  const filasUsadas = hoja.getUsedRange();
  let direccionTabla = "A1:" + ultimaCol + "1";
  if (filasUsadas && filasUsadas.getRowCount() > 1) {
    const filas = filasUsadas.getRowCount();
    direccionTabla = "A1:" + ultimaCol + filas;
  }

  const tabla = hoja.addTable(direccionTabla, true);
  tabla.setName(TABLA_HISTORIAL_COMERCIAL);
  return tabla;
}

function leerOvDesdeCapturaComercial(workbook: ExcelScript.Workbook): string {
  const tabla = tablaCapturaComercial(workbook);
  if (!tabla) {
    return "";
  }

  const datos = leerTabla(tabla);
  const idxOv = indiceColumnaFlexible(datos.headers, COL_COM_OV);
  if (idxOv < 0) {
    return "";
  }

  for (let i = 0; i < datos.rows.length; i++) {
    const ov = aTexto(datos.rows[i][idxOv]);
    if (ov !== "") {
      return ov;
    }
  }

  return "";
}

function leerOvComercial(workbook: ExcelScript.Workbook): string {
  const ovCaptura = leerOvDesdeCapturaComercial(workbook);
  if (ovCaptura !== "") {
    return ovCaptura;
  }
  return leerPedidoDynamicsDesdeResultados(workbook);
}

function filaComercialTieneDatos(headers: string[], fila: ValorCelda[]): boolean {
  const cliente = aTexto(valorCeldaFlexible(headers, fila, [COL_COM_CLIENTE]));
  const codigo = aTexto(valorCeldaFlexible(headers, fila, COL_COM_CODIGO));
  const piezas = aNumero(valorCeldaFlexible(headers, fila, COL_COM_PIEZAS));
  const oc = aTexto(valorCeldaFlexible(headers, fila, COL_COM_OC));
  return cliente !== "" || codigo !== "" || piezas > 0 || oc !== "";
}

function validarCuentaComercial(cuenta: string): string | null {
  if (cuenta === "") {
    return "Cliente es obligatorio (codigo Dynamics, ej. C0010).";
  }
  if (/^\d{6,}$/.test(cuenta)) {
    return (
      "La columna Cliente parece una orden de compra (" +
      cuenta +
      "). Use el codigo Dynamics (ej. C0010) en Cliente y la OC solo en Orden de compra."
    );
  }
  if (!/^[A-Za-z][A-Za-z0-9_-]{1,19}$/.test(cuenta)) {
    return (
      "El codigo de cliente no parece valido (" +
      cuenta +
      "). Use el formato de Dynamics (ej. C0010)."
    );
  }
  return null;
}

function normalizarOrdenCompra(oc: string): string {
  let texto = aTexto(oc as ValorCelda).trim();
  if (/^OC[\s.\-_]*/i.test(texto)) {
    texto = texto.replace(/^OC[\s.\-_]*/i, "").trim();
  }
  return texto;
}

function leerClienteComercial(headers: string[], fila: ValorCelda[]): string {
  const raw = aTexto(valorCeldaFlexible(headers, fila, [COL_COM_CLIENTE]));
  const cuenta = extraerCuentaDynamics(raw);
  if (cuenta !== "") {
    return cuenta;
  }
  return raw;
}

function validarCabeceraComercial(datos: TablaLeida): string | null {
  if (datos.rows.length === 0) {
    return "Agregue al menos una fila en la tabla de captura.";
  }

  if (indiceColumnaFlexible(datos.headers, [COL_COM_CLIENTE]) < 0) {
    return "Falta la columna Cliente.";
  }
  if (indiceColumnaFlexible(datos.headers, COL_COM_OC) < 0) {
    return "Falta la columna Orden de compra.";
  }
  if (indiceColumnaFlexible(datos.headers, COL_COM_FECHA) < 0) {
    return "Falta la columna Fecha de entrega.";
  }

  const idxOv = indiceColumnaFlexible(datos.headers, COL_COM_OV);
  let primeraFila = -1;
  let clienteRef = "";
  let ocRef = "";
  let fechaRef = "";

  for (let i = 0; i < datos.rows.length; i++) {
    const fila = datos.rows[i];
    if (!filaComercialTieneDatos(datos.headers, fila)) {
      continue;
    }

    const cliente = leerClienteComercial(datos.headers, fila);
    const oc = normalizarOrdenCompra(
      aTexto(valorCeldaFlexible(datos.headers, fila, COL_COM_OC))
    );
    const fecha = convertirFecha(
      valorCeldaFlexible(datos.headers, fila, COL_COM_FECHA)
    );
    const ov =
      idxOv >= 0 ? aTexto(datos.rows[i][idxOv]) : "";

    if (ov !== "") {
      return "Ya existe OrdenVenta (" + ov + "). Vacie la tabla o use un archivo nuevo.";
    }

    if (primeraFila < 0) {
      primeraFila = i;
      clienteRef = cliente;
      ocRef = oc;
      fechaRef = fecha;
    } else if (cliente !== "" && cliente !== clienteRef) {
      return "Todas las filas deben tener el mismo Cliente (" + clienteRef + ").";
    }
  }

  if (primeraFila < 0) {
    return "Agregue al menos una fila con datos en la tabla.";
  }

  const errCuenta = validarCuentaComercial(clienteRef);
  if (errCuenta !== null) {
    return errCuenta;
  }
  if (ocRef === "") {
    return "Orden de compra es obligatoria.";
  }
  const ocNormalizada = normalizarOrdenCompra(ocRef);
  if (ocNormalizada === "") {
    return "Orden de compra invalida despues de quitar prefijo OC.";
  }
  if (fechaRef === "") {
    return "Fecha de entrega es obligatoria.";
  }

  return null;
}

function construirCabeceraComercial(datos: TablaLeida): CrearPedidoBody {
  for (let i = 0; i < datos.rows.length; i++) {
    const fila = datos.rows[i];
    if (!filaComercialTieneDatos(datos.headers, fila)) {
      continue;
    }

    const cliente = leerClienteComercial(datos.headers, fila);
    const oc = normalizarOrdenCompra(
      aTexto(valorCeldaFlexible(datos.headers, fila, COL_COM_OC))
    );
    const fecha = convertirFecha(
      valorCeldaFlexible(datos.headers, fila, COL_COM_FECHA)
    );

    return {
      cliente: cliente,
      referenciaCliente: oc,
      descripcionPedido: "Pedido " + cliente,
      fechaEnvioSolicitada: fecha,
      fechaRecepcionSolicitada: "",
    };
  }

  return {
    cliente: "",
    referenciaCliente: "",
    descripcionPedido: "",
    fechaEnvioSolicitada: "",
    fechaRecepcionSolicitada: "",
  };
}

function diagnosticarLineasComercial(datos: TablaLeida): {
  headers: string[];
  filas: ValorCelda[][];
} {
  const filas: ValorCelda[][] = [];

  for (let i = 0; i < datos.rows.length; i++) {
    const fila = datos.rows[i];
    const codigo = aCodigoArticulo(valorCeldaFlexible(datos.headers, fila, COL_COM_CODIGO));
    const piezas = aNumero(valorCeldaFlexible(datos.headers, fila, COL_COM_PIEZAS));
    if (codigo !== "" && piezas > 0) {
      filas.push(fila);
    }
  }

  return { headers: datos.headers, filas: filas };
}

function validarLineasComercial(datos: TablaLeida): string | null {
  const pendientes = diagnosticarLineasComercial(datos);
  if (pendientes.filas.length === 0) {
    return "No hay lineas con Codigo y Piezas mayores que 0.";
  }

  if (indiceColumnaFlexible(datos.headers, COL_COM_CODIGO) < 0) {
    return "Falta la columna Codigo.";
  }
  if (indiceColumnaFlexible(datos.headers, COL_COM_PIEZAS) < 0) {
    return "Falta la columna Piezas.";
  }

  for (let i = 0; i < pendientes.filas.length; i++) {
    const fila = pendientes.filas[i];
    const codigo = aCodigoArticulo(valorCeldaFlexible(datos.headers, fila, COL_COM_CODIGO));
    const piezas = aNumero(valorCeldaFlexible(datos.headers, fila, COL_COM_PIEZAS));
    if (codigo === "") {
      return "Linea " + (i + 1) + ": Codigo es obligatorio.";
    }
    if (piezas <= 0) {
      return "Linea " + (i + 1) + ": Piezas debe ser mayor que 0.";
    }
  }

  return null;
}

function construirLineaComercial(headers: string[], fila: ValorCelda[]): LineaBody {
  const codigo = aCodigoArticulo(valorCeldaFlexible(headers, fila, COL_COM_CODIGO));
  const piezas = aNumero(valorCeldaFlexible(headers, fila, COL_COM_PIEZAS));
  const precio = aNumero(valorCeldaFlexible(headers, fila, COL_COM_PRECIO));
  const fecha = convertirFecha(valorCeldaFlexible(headers, fila, COL_COM_FECHA));
  const producto = aTexto(valorCeldaFlexible(headers, fila, COL_COM_PRODUCTO));

  return {
    codigoArticulo: codigo,
    cantidad: piezas,
    precioUnitario: Number.isNaN(precio) ? 0 : precio,
    porcentajeDescuento: 0,
    fechaEnvio: fecha,
    comentario: producto,
  };
}

function escribirOvEnCapturaComercial(tabla: ExcelScript.Table, ov: string): void {
  if (ov === "") {
    return;
  }

  let datos = leerTabla(tabla);
  let idxOv = indiceColumnaFlexible(datos.headers, COL_COM_OV);
  if (idxOv < 0) {
    tabla.addColumn(-1, null, "OrdenVenta");
    datos = leerTabla(tabla);
    idxOv = indiceColumnaFlexible(datos.headers, COL_COM_OV);
  }
  if (idxOv < 0) {
    return;
  }

  const cuerpo = tabla.getRangeBetweenHeaderAndTotal();
  const valores = cuerpo.getValues() as (string | number | boolean)[][];
  for (let r = 0; r < valores.length; r++) {
    if (aTexto(valores[r][idxOv] as ValorCelda) === "") {
      valores[r][idxOv] = ov;
    }
  }
  cuerpo.setValues(valores);
}

function limpiarCapturaComercial(workbook: ExcelScript.Workbook): void {
  const tabla = tablaCapturaComercial(workbook);
  if (tabla) {
    limpiarTablaDatos(tabla);
  }
}

async function crearCabeceraComercial(workbook: ExcelScript.Workbook): Promise<void> {
  try {
    const tabla = tablaCapturaComercial(workbook);
    if (!tabla) {
      escribirResultado(
        workbook,
        "ERROR",
        "No se encontro la tabla " + TABLA_CAPTURA_COMERCIAL + " en la hoja Captura.",
        ""
      );
      return;
    }

    const datos = leerTabla(tabla);
    const err = validarCabeceraComercial(datos);
    if (err !== null) {
      escribirResultado(workbook, "ERROR", err, "");
      return;
    }

    const body = construirCabeceraComercial(datos);
    const base = leerApiBaseUrl(workbook);
    const respP = await llamarApi(
      base + "/api/pedidos/crear",
      "POST",
      JSON.stringify(body)
    );

    const ov = extraerOvDeRespuesta(respP);
    if (ov === "") {
      escribirResultado(
        workbook,
        "ERROR",
        "La API respondio pero no trajo el numero OV. Revise Dynamics o reinicie el backend.",
        ""
      );
      return;
    }

    escribirOvEnCapturaComercial(tabla, ov);

    escribirResultado(
      workbook,
      "CABECERA OK",
      "Cabecera creada. OV: " + ov + " [" + SCRIPT_COMERCIAL_VERSION + "]",
      ov
    );
  } catch (error) {
    const raw = String(error);
    const ovRescatada = extraerOvDeTexto(raw);
    const msg = obtenerMensajeError(error);
    if (ovRescatada !== "") {
      const tabla = tablaCapturaComercial(workbook);
      if (tabla) {
        escribirOvEnCapturaComercial(tabla, ovRescatada);
      }
      escribirResultado(
        workbook,
        "CABECERA OK",
        "Cabecera creada. OV: " + ovRescatada,
        ovRescatada
      );
      return;
    }
    escribirResultado(workbook, "ERROR", msg !== "" ? msg : raw, ovRescatada);
  }
}

async function crearLineasComercial(workbook: ExcelScript.Workbook): Promise<void> {
  try {
    const tabla = tablaCapturaComercial(workbook);
    if (!tabla) {
      escribirResultado(
        workbook,
        "ERROR",
        "No se encontro la tabla " + TABLA_CAPTURA_COMERCIAL + ".",
        ""
      );
      return;
    }

    const ov = leerOvDesdeCapturaComercial(workbook);
    if (ov === "") {
      escribirResultado(
        workbook,
        "ERROR",
        MSG_CABECERA_COMERCIAL_PRIMERO +
          " La OV debe estar en la columna OrdenVenta de Captura.",
        ""
      );
      return;
    }

    const datos = leerTabla(tabla);
    const err = validarLineasComercial(datos);
    if (err !== null) {
      escribirResultado(workbook, "ERROR", err, ov);
      return;
    }

    const pendientes = diagnosticarLineasComercial(datos);
    const lineas: LineaBody[] = [];
    for (let i = 0; i < pendientes.filas.length; i++) {
      lineas.push(construirLineaComercial(pendientes.headers, pendientes.filas[i]));
    }

    const tablaHist = asegurarTablaHistorialComercial(workbook);
    if (tablaHist === null) {
      escribirResultado(
        workbook,
        "ERROR",
        "No se encontro la tabla tbl_historial en la hoja Historial. La captura no se vaciara.",
        ov
      );
      return;
    }

    const filasHistAntes = tablaHist.getRowCount();
    const cabecera = construirCabeceraComercial(datos);
    const pedidoSnapshot: TablaLeida = {
      headers: [
        COL_PEDIDO_CLIENTE,
        COL_PEDIDO_REFERENCIA,
        COL_PEDIDO_DESCRIPCION,
      ],
      rows: [[cabecera.cliente, cabecera.referenciaCliente, cabecera.descripcionPedido]],
    };

    const base = leerApiBaseUrl(workbook);
    const bodyL: CrearLineasBody = {
      salesOrderNumber: ov,
      lineas: lineas,
    };

    const respL = (await llamarApi(
      base + "/api/pedidos/lineas",
      "POST",
      JSON.stringify(bodyL)
    )) as CrearLineasApi;

    const filasGuardadas = registrarHistorialLineas(
      workbook,
      new Date(),
      ov,
      pedidoSnapshot,
      lineas,
      tablaHist
    );

    const filasHistDespues = tablaHist.getRowCount();
    if (filasGuardadas === 0 || filasHistDespues <= filasHistAntes) {
      escribirResultado(
        workbook,
        "ERROR",
        "Las lineas se crearon en Dynamics (" +
          ov +
          ") pero no se guardaron en tbl_historial. La tabla de captura NO se vacio.",
        ov
      );
      return;
    }

    limpiarCapturaComercial(workbook);

    escribirResultado(
      workbook,
      "LINEAS OK",
      (respL.mensaje || "Lineas creadas en " + ov) +
        ". " +
        lineas.length +
        " linea(s) en tbl_historial. Captura vaciada. [" +
        SCRIPT_COMERCIAL_VERSION +
        "]",
      ov
    );
  } catch (error) {
    escribirResultado(
      workbook,
      "ERROR",
      obtenerMensajeError(error),
      leerOvDesdeCapturaComercial(workbook)
    );
  }
}

async function main(workbook: ExcelScript.Workbook): Promise<void> {
  await crearLineas(workbook);
}
