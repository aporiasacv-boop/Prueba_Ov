## Quien hace que

| Donde | Que se configura | Una sola vez o por Excel |
|-------|------------------|--------------------------|
| **PC servidor** (escritorio siempre encendido) | Java, Spring Boot, ngrok, `application.yml` con secret de Azure | Una sola vez en esa PC |
| **Cada Excel** | URL de ngrok en **Resultado!B1**, tablas, 3 scripts en botones | B1 en cada copia del libro; scripts al crear el Excel |

Spring Boot **no sabe** la URL de ngrok. Solo escucha en `http://localhost:8080`. La URL publica la muestra **ngrok** en su ventana (linea `Forwarding https://.... -> http://localhost:8080`).

---

## PC servidor — configuracion (una vez)

1. Clonar el repo.
2. Instalar **Java 17+** y **ngrok** (`ngrok config add-authtoken ...`).
3. Copiar `dynamics-integration/src/main/resources/application.example.yml` a `application.yml`. Pegar el **client-secret** real de Azure (comillas simples si lleva `~`). No subir a GitHub.
4. Desactivar suspension de Windows en esa PC.

### Cada vez que reinicie la PC (o tras un corte de luz)

**Ventana 1 — Spring Boot** (dejar abierta):

```powershell
cd dynamics-integration
.\run.ps1
```

Espere: `Started DynamicsIntegrationApplication`.

**Ventana 2 — API Olnatura** (doble clic en `Excel_Restringido-\1_API_OLNATURA.bat`, puerto 8011).

**Ventana 3 — ngrok** (doble clic en `2_NGROK.bat` en esta carpeta).

Copie las 2 URLs **https**:
- **8080** → Excel Ordenes, celda **Resultado!B1**
- **8011** → Excel NIKZON (equipo actualiza codigo y `2_INSTALAR_EXCEL.bat`)

**Verificar** (ventana 3, opcional):

```powershell
cd dynamics-integration
.\verify-servidor.ps1
.\verify-servidor.ps1 -NgrokUrl "https://SU-URL-NGROK"
```

Si los 4 pasos dan OK, el servidor esta listo. **Cualquier Excel** con esa URL en B1 puede conectar.

---

## Cada Excel — que necesita (credencial = solo la URL)

1. Archivo guardado en **OneDrive o SharePoint** (no solo en disco local).
2. Abrir en **Excel Online** (navegador), no en la app de escritorio.
3. Hoja **Resultado**, celda **B1**: URL `https://....ngrok-free.dev` **sin** `/` al final.
4. Tabla **tblResultados** con columnas: **Estado**, **Error** (recomendado tambien: Fecha de ejecucion, PedidoDynamics / OrdenVenta).
5. Tablas **tblPedido** y **tblLineas** con las columnas del libro plantilla.
6. Tres scripts en **Automatizar** (ver seccion 4).

No hace falta poner el secret de Azure en el Excel. Eso vive solo en `application.yml` de la PC servidor.

---

## 1. Arrancar Spring Boot

```powershell
cd dynamics-integration
.\run.ps1
```

## 2. Arrancar ngrok

Doble clic en `2_NGROK.bat` (no usar `ngrok http 8080` suelto).

## 3. Excel — URL en B1

Pegue la URL https de ngrok en **Resultado!B1**. Guarde el libro.

## 4. Tres botones (un script por boton)

En la PC servidor (o donde tenga el repo):

```powershell
cd office-scripts
.\build-office-scripts.ps1
```

Pegue en Excel el contenido **completo** de cada archivo:

| Boton en Excel | Archivo |
|----------------|---------|
| Probar conexion | `office-scripts/dist/ProbarConexion.osts.ts` |
| Crear pedido | `office-scripts/dist/CrearPedido.osts.ts` |
| Crear lineas | `office-scripts/dist/CrearLineas.osts.ts` |

En **Automatizar → Editor de scripts**: script nuevo → pegar todo → guardar → asignar al boton.

### Tras crear la cabecera (Crear pedido)

El script rellena en `tblLineas` las columnas **Referencia de cliente** y **OrdenVenta** si estaban vacias.

## 5. Flujo de un pedido

1. **Probar conexion** → Estado `CONEXION OK`
2. **Crear pedido** → Estado `CABECERA OK`, numero OV en tblResultados
3. **Crear lineas** → Estado `LINEAS OK`

Si existe **tblHistorial** (opcional), al terminar Crear lineas se guarda historial.

## 6. Prueba sin Excel

```powershell
cd dynamics-integration
.\verify-servidor.ps1 -NgrokUrl "https://SU-URL-NGROK"
```

O flujo completo de pedido de prueba:

```powershell
.\test-api.ps1
```

## 7. Si Excel falla

| Mensaje en Error | Causa |
|------------------|-------|
| Escribe la URL ngrok en Resultado!B1 | B1 vacia |
| HTML / pagina web | ngrok apagado, URL vieja, o script desactualizado |
| HTTP 404 en ping | Spring Boot sin actualizar o URL mal escrita |
| Servidor OK, pero Dynamics/Azure fallo | `application.yml` en la PC servidor |
| Failed to fetch | ngrok apagado o firewall |
| No existe tblResultados | Falta la tabla en hoja Resultado |

**Importante:** Si reinician ngrok, la URL puede cambiar → actualizar B1 en **todos** los Excel que usen el sistema.
