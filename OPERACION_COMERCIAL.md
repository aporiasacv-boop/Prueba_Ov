## Comercial — operacion diaria

Comercial comparte **el mismo servidor Java** y **la misma URL ngrok** que marketing. Solo cambia el Excel y los scripts de los botones.

| Excel | Scripts en botones |
|-------|-------------------|
| Marketing (`Prueba_OV marketing - Preparado.xlsx`) | ProbarConexion, CrearCabeceraMarketing, CrearLineasMarketing |
| Comercial (`Layout Comercial - Preparado.xlsx`) | ProbarConexion, CrearCabeceraComercial, CrearLineasComercial |

---

## Preparar el Excel comercial (una vez por copia)

```powershell
cd "C:\ruta\al\repo\Prueba_Ov"
python prepare_comercial_excel.py
```

Archivo generado: `Layout Comercial - Preparado.xlsx`

Subir a OneDrive/SharePoint y abrir en **Excel Online**.

---

## Generar scripts para botones

```powershell
cd office-scripts
.\build-office-scripts.ps1
```

| Boton | Archivo a pegar completo |
|-------|---------------------------|
| Probar conexion | `office-scripts/dist/ProbarConexion.osts.ts` |
| Generar cabecera | `office-scripts/dist/CrearCabeceraComercial.osts.ts` |
| Generar lineas | `office-scripts/dist/CrearLineasComercial.osts.ts` |

En **Automatizar → Editor de scripts**: script nuevo → pegar todo → guardar → asignar al boton.

---

## URL ngrok (B1)

Hoja **Resultado**, celda **B1**: misma URL https que marketing, sin `/` al final.

Si reinician ngrok y cambia la URL, actualicen B1 en **ambos** Excel.

---

## Flujo de un pedido comercial

1. Llenar hoja **Captura** (`tblCapturaComercial`): una fila por producto, mismo Cliente en todas.
2. **Probar conexion** → `CONEXION OK`
3. **Generar cabecera** → `CABECERA OK`, columna OrdenVenta rellena
4. **Generar lineas** → `LINEAS OK`, filas en `tbl_historial`, tabla Captura vacia

Columnas obligatorias: Cliente (C0010), Codigo, Piezas, Orden de cliente, Fecha de envio.
Opcional: Fecha de recepcion (cabecera en Dynamics). OrdenVenta la llena el script.

---

## Despliegue en PC servidor (marketing + comercial)

1. En su PC: commit + push del repo.
2. En la PC servidor: `git pull`
3. Reiniciar Spring Boot **una vez** (marketing y comercial usan el mismo proceso):

```powershell
cd dynamics-integration
.\run.ps1
```

4. ngrok sigue igual: `ngrok http 8080`
5. Marketing: sin cambios si ya tenia sus scripts.
6. Comercial: subir `Layout Comercial - Preparado.xlsx` a OneDrive y asignar los 3 scripts comerciales.

Tras el pull, marketing y comercial funcionan en paralelo con la misma URL en B1.

---

## Prueba local sin Excel

```powershell
cd dynamics-integration
.\verify-servidor.ps1
.\test-api.ps1
```
