## 0. Secret de Azure (solo la primera vez)

Copie `dynamics-integration/src/main/resources/application.example.yml` a `application.yml` (misma carpeta). En **client-secret** pegue el valor real (entre comillas simples `'...'` si lleva `~`). Guarde. Ese archivo no se sube a GitHub.

## 1. Arrancar el programa (Terminal en VS Code)

Menú **Terminal → New Terminal**.

```powershell
cd dynamics-integration
.\run.ps1
```

Espere hasta que aparezca algo como **Tomcat started on port 8080** (o “Started DynamicsIntegrationApplication”). Deje esa ventana abierta.

## 2. ngrok (otra terminal)

Menú **Terminal → New Terminal** (segunda pestaña).

```powershell
ngrok http 8080
```

Copie la dirección **https** que muestra (por ejemplo `https://xxxx.ngrok-free.dev`). Deje esta ventana abierta también.

## 3. Excel — URL

En la hoja **Resultado**, celda **B1**, pegue esa dirección **https** (puede quitar la `/` del final). Guarde el libro.

## 4. Tres scripts en Excel (un archivo por botón)

Ejecute `office-scripts\build-office-scripts.ps1`. En la carpeta `office-scripts\dist\` salen **cuatro** archivos `.osts.ts`; los **tres primeros** son los botones habituales:

| Botón en Excel | Pegar el archivo completo |
|----------------|---------------------------|
| Probar conexión | `dist/ProbarConexion.osts.ts` |
| Crear pedido (cabecera) | `dist/CrearPedido.osts.ts` |
| Líneas de pedido | `dist/CrearLineas.osts.ts` |

Cada archivo es **un script independiente** en Automatizar: cree un script nuevo, pegue **todo** el contenido del archivo y guarde. El `main` de cada uno solo ejecuta esa acción (no hace falta elegir función en un desplegable).

**Opcional (varios pedidos a la vez):** `dist/SubirLote.osts.ts` — ver sección 6 más abajo.

### Tras crear la cabecera (botón Crear pedido)

Cuando Dynamics devuelve la OV, el script:

1. Rellena celdas vacías de la columna **Referencia de cliente** en `tblLineas` con la misma referencia del pedido.
2. Si hace falta, **crea** esa columna al final de la tabla.
3. Igual con la columna **OrdenVenta** (número OV, ej. `OV0027999`).

Así las líneas quedan enlazadas a la cabecera antes de pulsar **Crear lineas**.

## 5. Flujo de un solo pedido (como hasta ahora)

1. **Probar Conexion**
2. **Crear Pedido**
3. **Crear Lineas**

Si existe la tabla **tblHistorial** (hoja **Historial**), al terminar **Crear Lineas** se guarda una fila por cada línea en el historial.

## 6. Varios pedidos de una vez (**Subir lote**)

1. Hoja **Pedido**: en **tblPedido** varias filas (un pedido por fila). Cada **Referencia de cliente** debe ser **distinta**.
2. Hoja **Lineas**: en **tblLineas** todas las líneas de todos los pedidos.
   - Si hay **más de un pedido**, agregue la columna **Referencia de cliente** en **tblLineas** y repita en cada línea la misma referencia que en el pedido al que pertenece.
   - Si hay **un solo pedido**, la columna **Referencia de cliente** en líneas es opcional (todas las líneas van a ese pedido).
3. Cree la hoja **Historial** con una tabla llamada **tblHistorial** y estas columnas **exactas** (en el orden que prefiera; el script rellena por nombre):

| Columna |
|---------|
| FechaHora |
| OrdenVenta |
| ReferenciaCliente |
| Cliente |
| DescripcionPedido |
| Codigo_Articulo |
| Cantidad |
| PrecioUnitario |
| Porcentaje de descuento |
| Fecha de envio |
| Comentario |

4. Ejecute **Subir lote**. Se suben los pedidos **en el orden de las filas** de **tblPedido** (cabecera + líneas de esa referencia, uno tras otro). Si todo sale bien, se escribe el historial y se **vacían** **tblPedido** y **tblLineas** para la siguiente captura.

Si falta **tblHistorial**, **Subir lote** no se ejecuta (evita subir sin guardar historial).

## 7. Prueba rápida sin Excel

Con el paso 1 en marcha:

```powershell
cd dynamics-integration
.\test-api.ps1
```
