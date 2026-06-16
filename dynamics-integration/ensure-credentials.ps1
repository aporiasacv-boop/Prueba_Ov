function Test-DynamicsCredentialsConfigured {
    param([string]$ConfigPath)

    if (-not (Test-Path $ConfigPath)) {
        return $false
    }

    $content = Get-Content -Path $ConfigPath -Raw -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($content)) {
        return $false
    }

    $fields = @{
        "tenant-id"     = $false
        "client-id"     = $false
        "client-secret" = $false
        "resource"      = $false
    }

    foreach ($key in $fields.Keys) {
        if ($content -match "(?m)^\s*$key\s*:\s*(.+)\s*$") {
            $value = $matches[1].Trim().Trim("'").Trim('"')
            if ($value -and $value -notmatch '^\$\{') {
                $fields[$key] = $true
            }
        }
    }

    return ($fields.Values -notcontains $false)
}

function Get-DynamicsCredentialValue {
    param(
        [string]$Content,
        [string]$Key
    )

    if ($Content -match "(?m)^\s*$Key\s*:\s*(.+)\s*$") {
        $value = $matches[1].Trim()
        if ($value.StartsWith("'") -and $value.EndsWith("'")) {
            return $value.Substring(1, $value.Length - 2).Replace("''", "'")
        }
        if ($value.StartsWith('"') -and $value.EndsWith('"')) {
            return $value.Substring(1, $value.Length - 2)
        }
        if ($value -notmatch '^\$\{') {
            return $value
        }
    }

    return ""
}

function Show-DynamicsCredentialsPanel {
    param(
        [string]$TenantId = "",
        [string]$ClientId = "",
        [string]$ClientSecret = ""
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Credenciales Azure - Dynamics"
    $form.Size = New-Object System.Drawing.Size(520, 320)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $info = New-Object System.Windows.Forms.Label
    $info.Location = New-Object System.Drawing.Point(16, 12)
    $info.Size = New-Object System.Drawing.Size(470, 40)
    $info.Text = "Pegue las credenciales de Azure App Registration.`nSe guardan en application.yml (local, no se sube a GitHub)."
    $form.Controls.Add($info)

    $lblTenant = New-Object System.Windows.Forms.Label
    $lblTenant.Location = New-Object System.Drawing.Point(16, 58)
    $lblTenant.Size = New-Object System.Drawing.Size(120, 20)
    $lblTenant.Text = "Tenant ID"
    $form.Controls.Add($lblTenant)

    $txtTenant = New-Object System.Windows.Forms.TextBox
    $txtTenant.Location = New-Object System.Drawing.Point(16, 78)
    $txtTenant.Size = New-Object System.Drawing.Size(470, 24)
    $txtTenant.Text = $TenantId
    $form.Controls.Add($txtTenant)

    $lblClient = New-Object System.Windows.Forms.Label
    $lblClient.Location = New-Object System.Drawing.Point(16, 110)
    $lblClient.Size = New-Object System.Drawing.Size(120, 20)
    $lblClient.Text = "Client ID"
    $form.Controls.Add($lblClient)

    $txtClient = New-Object System.Windows.Forms.TextBox
    $txtClient.Location = New-Object System.Drawing.Point(16, 130)
    $txtClient.Size = New-Object System.Drawing.Size(470, 24)
    $txtClient.Text = $ClientId
    $form.Controls.Add($txtClient)

    $lblSecret = New-Object System.Windows.Forms.Label
    $lblSecret.Location = New-Object System.Drawing.Point(16, 162)
    $lblSecret.Size = New-Object System.Drawing.Size(120, 20)
    $lblSecret.Text = "Client Secret"
    $form.Controls.Add($lblSecret)

    $txtSecret = New-Object System.Windows.Forms.TextBox
    $txtSecret.Location = New-Object System.Drawing.Point(16, 182)
    $txtSecret.Size = New-Object System.Drawing.Size(470, 24)
    $txtSecret.UseSystemPasswordChar = $true
    $txtSecret.Text = $ClientSecret
    $form.Controls.Add($txtSecret)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Location = New-Object System.Drawing.Point(300, 230)
    $btnOk.Size = New-Object System.Drawing.Size(85, 28)
    $btnOk.Text = "Guardar"
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $btnOk
    $form.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(401, 230)
    $btnCancel.Size = New-Object System.Drawing.Size(85, 28)
    $btnCancel.Text = "Cancelar"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancel
    $form.Controls.Add($btnCancel)

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    return [PSCustomObject]@{
        TenantId     = $txtTenant.Text.Trim()
        ClientId     = $txtClient.Text.Trim()
        ClientSecret = $txtSecret.Text.Trim()
    }
}

function Write-DynamicsApplicationYml {
    param(
        [string]$ConfigPath,
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $escapedSecret = $ClientSecret.Replace("'", "''")
    $yaml = @"
server:
  port: 8080
  error:
    include-message: always
    include-exception: false

spring:
  application:
    name: dynamics-integration

azure:
  tenant-id: $TenantId
  client-id: $ClientId
  client-secret: '$escapedSecret'
  token-version: v1
  resource: https://olnatura-produccion.operations.dynamics.com
  scope: https://olnatura-produccion.operations.dynamics.com/.default

dynamics:
  base-url: https://olnatura-produccion.operations.dynamics.com
  data-area-id: olna
  default-currency: MXN
  line:
    sales-unit-symbol: PZA
    shipping-site-id: OLNATURA
    shipping-warehouse-id: PTM

management:
  endpoints:
    web:
      exposure:
        include: health,info

logging:
  level:
    root: WARN
    com.olnatura.dynamics: INFO
"@

    $dir = Split-Path $ConfigPath -Parent
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }

    Set-Content -Path $ConfigPath -Value $yaml -Encoding UTF8
}

function Ensure-DynamicsCredentials {
    param(
        [string]$ConfigPath,
        [switch]$ForcePrompt
    )

    if (-not $ForcePrompt -and (Test-DynamicsCredentialsConfigured -ConfigPath $ConfigPath)) {
        return
    }

    $existing = ""
    if (Test-Path $ConfigPath) {
        $existing = Get-Content -Path $ConfigPath -Raw
    }

    $prefill = [PSCustomObject]@{
        TenantId     = Get-DynamicsCredentialValue -Content $existing -Key "tenant-id"
        ClientId     = Get-DynamicsCredentialValue -Content $existing -Key "client-id"
        ClientSecret = Get-DynamicsCredentialValue -Content $existing -Key "client-secret"
    }

    $creds = Show-DynamicsCredentialsPanel `
        -TenantId $prefill.TenantId `
        -ClientId $prefill.ClientId `
        -ClientSecret $prefill.ClientSecret

    if ($null -eq $creds) {
        Write-Host "Arranque cancelado: faltan credenciales Azure." -ForegroundColor Yellow
        exit 1
    }

    foreach ($field in @(
            @{ Name = "Tenant ID"; Value = $creds.TenantId },
            @{ Name = "Client ID"; Value = $creds.ClientId },
            @{ Name = "Client Secret"; Value = $creds.ClientSecret }
        )) {
        if ([string]::IsNullOrWhiteSpace($field.Value)) {
            Write-Host "$($field.Name) es obligatorio." -ForegroundColor Red
            exit 1
        }
    }

    Write-DynamicsApplicationYml `
        -ConfigPath $ConfigPath `
        -TenantId $creds.TenantId `
        -ClientId $creds.ClientId `
        -ClientSecret $creds.ClientSecret

    Write-Host "Credenciales guardadas en application.yml" -ForegroundColor Green
}
