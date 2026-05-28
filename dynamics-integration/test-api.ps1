param([string]$BaseUrl = "http://localhost:8080")

function Invoke-Api {
    param([string]$Url, [string]$HttpMethod = "Get", [string]$Body = $null)
    if ($Body) {
        return Invoke-RestMethod -Uri $Url -Method $HttpMethod -Body $Body -ContentType "application/json"
    }
    return Invoke-RestMethod -Uri $Url -Method $HttpMethod
}

Invoke-Api -Url "$BaseUrl/api/token/refresh" -HttpMethod Post | Out-Null
$pedido = @{
    cliente = "C0010"
    referenciaCliente = "OC12345"
    descripcionPedido = "Pedido farmacia"
    fechaEnvioSolicitada = "2026-05-25"
    fechaRecepcionSolicitada = "2026-05-30"
} | ConvertTo-Json
$respPedido = Invoke-Api -Url "$BaseUrl/api/pedidos/crear" -HttpMethod Post -Body $pedido
$lineas = @{
    salesOrderNumber = $respPedido.salesOrderNumber
    lineas = @(@{
        codigoArticulo = "501034001500"
        cantidad = 7000
        precioUnitario = 200
        porcentajeDescuento = 5
        fechaEnvio = "2026-05-25"
        comentario = ""
    })
} | ConvertTo-Json -Depth 5
Invoke-Api -Url "$BaseUrl/api/pedidos/lineas" -HttpMethod Post -Body $lineas
