# =========================
# Backup perfil + rclone
# =========================
# Execute como Administrador

$ErrorActionPreference = "Stop"

# === Ajuste apenas este usuário (mantive fixo como você pediu) ===
$UserName   = "nome.sobrenome"
$ProfileSrc = "C:\Users\$UserName"

# === Pastas de trabalho ===
$Stamp      = (Get-Date).ToString("yyyyMMdd-HHmmss")
$WorkRoot   = "C:\BackupPerfis"
$CloneDir   = Join-Path $WorkRoot "$UserName-clone-$Stamp"
$ZipPath    = Join-Path $WorkRoot "$UserName-$Stamp.zip"

# === Onde instalar rclone ===
$ToolsDir   = "C:\Tools\rclone"
$rcloneExe  = Join-Path $ToolsDir "rclone.exe"
$rcloneZip  = Join-Path $ToolsDir "rclone-current-windows-amd64.zip"

# === Nome do remoto e conteúdo do token ===
$RemoteName = "onedrive_nw"
$RcloneConfigDir = Join-Path $env:APPDATA "rclone"
$RcloneConfig    = Join-Path $RcloneConfigDir "rclone.conf"

# === Token fornecido (copiado exatamente como enviado) ===
$TokenJson = @'
{"access_token":"eyJ0eXAiOiJKV1QiLCJub25jZSI6IkU2eTIzcE1ZYVZlTHpuc2hGTWF5N3lkTWNWN0hfeG0xV2ZpUE9ta1JxMDQiLCJhbGciOiJSUzI1NiIsIng1dCI6IkpZaEFjVFBNWl9MWDZEQmxPV1E3SG4wTmVYRSIsImtpZCI6IkpZaEFjVFBNWl9MWDZEQmxPV1E3SG4wTmVYRSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC84ZGM0YjIwYi01YjE5LTQ2YjQtOGUwNS00ZTBiNjEyNzFmZTEvIiwiaWF0IjoxNzU0OTIzMjgyLCJuYmYiOjE3NTQ5MjMyODIsImV4cCI6MTc1NDkyODYxOSwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsicDEiXSwiYWlvIjoiQVpRQWEvOFpBQUFBZlZuSll2Unh6N1NJTGJmbkphZGRZOGhLYktodE9RT00rVzJ4c1RzUHh5SUUvYzJiMXJQQUZPcnIzRThGRUwrdXMwbUlNVUUxR2djZWIwaUltZ3BlSUVDbmpiNmpCY0Q1NkxFV1daTzZGTjhlVWZrWkNKSlpnY2M2Z01Xb1dMa29IUmtqU1Fka1M1N1hwNlRXV3NDZHBWbHNiczZWY203b2w2dW9hUWQ2RGgwV1BaZ2ZNQWtrNDBpQi9jQnNlQlB1IiwiYW1yIjpbInB3ZCIsIm1mYSJdLCJhcHBfZGlzcGxheW5hbWUiOiJyY2xvbmUiLCJhcHBpZCI6ImIxNTY2NWQ5LWVkYTYtNDA5Mi04NTM5LTBlZWMzNzZhZmQ1OSIsImFwcGlkYWNyIjoiMSIsImZhbWlseV9uYW1lIjoiT2xpdmVpcmEgUmVpcyBEZSBTb3VzYSIsImdpdmVuX25hbWUiOiJWaW5pY2l1cyIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjE4Ny4xMDIuMTU1LjMwIiwibmFtZSI6IlZpbmljaXVzIE9saXZlaXJhIFJlaXMgRGUgU291c2EiLCJvaWQiOiI0NGViMWEyNC00MWU2LTQzNjYtYWYzMS00ZjU3MjgwZWVhODQiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzIwMDNFNDRFNzkwNyIsInJoIjoiMS5BVmdBQzdMRWpSbGJ0RWFPQlU0TFlTY2Y0UU1BQUFBQUFBQUF3QUFBQUFBQUFBRHlBSDVZQUEuIiwic2NwIjoiRmlsZXMuUmVhZCBGaWxlcy5SZWFkLkFsbCBGaWxlcy5SZWFkV3JpdGUgRmlsZXMuUmVhZFdyaXRlLkFsbCBTaXRlcy5SZWFkLkFsbCBwcm9maWxlIG9wZW5pZCBlbWFpbCIsInNpZCI6IjAwNmQyMTk5LTIzZmMtYjJmNy1jODkxLTgzYTc2ZmM5ZThmZiIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IkxDVTMzRDYxMDdOSWlUbUU5b09MM1o4cWIxZWdvUnNxb0V4ZzR2ejZLRnMiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiU0EiLCJ0aWQiOiI4ZGM0YjIwYi01YjE5LTQ2YjQtOGUwNS00ZTBiNjEyNzFmZTEiLCJ1bmlxdWVfbmFtZSI6InZpbmljaXVzLnNvdXNhQG53YWR2LmNvbS5iciIsInVwbiI6InZpbmljaXVzLnNvdXNhQG53YWR2LmNvbS5iciIsInV0aSI6IlN5VU5nbUtxN1VtRW5zcXpsV0FaQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjdiZTQ0YzhhLWFkYWYtNGUyYS04NGQ2LWFiMjY0OWUwOGExMyIsImJhZjM3YjNhLTYxMGUtNDVkYS05ZTYyLWQ5ZDFlNWU4OTE0YiIsImM0MzBiMzk2LWU2OTMtNDZjYy05NmYzLWRiMDFiZjhiYjYyYSIsIjgzMjkxNTNiLTMxZDAtNDcyNy1iOTQ1LTc0NWViM2JjNWYzMSIsImIwZjU0NjYxLTJkNzQtNGM1MC1hZmEzLTFlYzgwM2YxMmVmZSIsIjliODk1ZDkyLTJjZDMtNDRjNy05ZDAyLWE2YWMyZDVlYTVjMyIsImU4NjExYWI4LWMxODktNDZlOC05NGUxLTYwMjEzYWIxZjgxNCIsIjE3MzE1Nzk3LTEwMmQtNDBiNC05M2UwLTQzMjA2MmNhY2ExOCIsImYwMjNmZDgxLWE2MzctNGI1Ni05NWZkLTc5MWFjMDIyNjAzMyIsIjljNmRmMGYyLTFlN2MtNGRjMy1iMTk1LTY2ZGZiZDI0YWE4ZiIsIjExNDUxZDYwLWFjYjItNDVlYi1hN2Q2LTQzZDBmMDEyNWMxMyIsImQzN2M4YmVkLTA3MTEtNDQxNy1iYTM4LWI0YWJlNjZjZTRjMiIsIjVjNGY5ZGNkLTQ3ZGMtNGNmNy04YzlhLTllNDIwN2NiZmM5MSIsIjc0OTVmZGM0LTM0YzQtNGQxNS1hMjg5LTk4Nzg4Y2UzOTlmZCIsIjg5MmM1ODQyLWE5YTYtNDYzYS04MDQxLTcyYWEwOGNhM2NmNiIsIjY5MDkxMjQ2LTIwZTgtNGE1Ni1hYTRkLTA2NjA3NWIyYTdhOCIsIjVmMjIyMmIxLTU3YzMtNDhiYS04YWQ1LWQ0NzU5ZjFmZGU2ZiIsImZkZDdhNzUxLWI2MGItNDQ0YS05ODRjLTAyNjUyZmU4ZmExYyIsIjI1YTUxNmVkLTJmYTAtNDBlYS1hMmQwLTEyOTIzYTIxNDczYSIsImZlOTMwYmU3LTVlNjItNDdkYi05MWFmLTk4YzNhNDlhMzhiMSIsImYyOGExZjUwLWY2ZTctNDU3MS04MThiLTZhMTJmMmFmNmI2YyIsIjA1MjY3MTZiLTExM2QtNGMxNS1iMmM4LTY4ZTNjMjJiOWY4MCIsIjE3MDcxMjVlLTBhYTItNGQ0ZC04NjU1LWE3Yzc4NmM3NmEyNSIsIjQ1ZDhkM2M1LWM4MDItNDVjNi1iMzJhLTFkNzBiNWUxZTg2ZSIsImU2ZDFhMjNhLWRhMTEtNGJlNC05NTcwLWJlZmM4NmQwNjdhNyIsIjMxMzkyZmZiLTU4NmMtNDJkMS05MzQ2LWU1OTQxNWEyY2M0ZSIsImFhMzgwMTRmLTA5OTMtNDZlOS05YjQ1LTMwNTAxYTIwOTA5ZCIsIjE5NGFlNGNiLWIxMjYtNDBiMi1iZDUiL..."
,"token_type":"Bearer","refresh_token":"1TRUNCADO.AVgAC7LEjRlbtEaOBU4LYScf4dllVrGm7ZJAhTkO7Ddq_VnyAH5YAA.AgABAwEAAABVrSpeuWamR..."
,"expiry":"2025-08-11T13:10:18.4465725-03:00","expires_in":5036}
'@

# =========================
# 1) Pré-checagens
# =========================
if (-not (Test-Path $ProfileSrc)) {
    throw "Path do perfil não encontrado: $ProfileSrc"
}

# Cria diretórios
New-Item -Path $WorkRoot -ItemType Directory -Force | Out-Null
New-Item -Path $CloneDir -ItemType Directory -Force | Out-Null
New-Item -Path $ToolsDir -ItemType Directory -Force | Out-Null
New-Item -Path $RcloneConfigDir -ItemType Directory -Force | Out-Null

Write-Host "Clonando perfil de $UserName..."

# =========================
# 2) Clone do perfil (Robocopy)
#    - XJ evita loops em junctions de AppData
#    - /COPY:DATSO mantém o máximo de metadados possível
# =========================
$robolog = Join-Path $WorkRoot "robocopy-$UserName-$Stamp.log"
$roboArgs = @(
    "`"$ProfileSrc`"", "`"$CloneDir`"",
    "/MIR", "/R:1", "/W:1",
    "/XJ",
    "/COPY:DATSO", "/DCOPY:DAT",
    "/NFL","/NDL","/NP","/TEE","/LOG:`"$robolog`""
)

# Excluir alguns caches grandes e inúteis no zip
$excludeDirs = @(
    "AppData\Local\Temp",
    "AppData\Local\Microsoft\Edge\User Data\Default\Cache",
    "AppData\Local\Google\Chrome\User Data\Default\Cache",
    "AppData\Local\Packages\*\AC",
    "AppData\Local\Packages\*\TempState",
    "AppData\Local\CrashDumps"
)
foreach ($d in $excludeDirs) { $roboArgs += @("/XD", "`"$ProfileSrc\$d`"") }

Start-Process -FilePath "robocopy.exe" -ArgumentList $roboArgs -Wait -NoNewWindow

Write-Host "Clone concluído em: $CloneDir"

# =========================
# 3) Zip do clone
# =========================
Write-Host "Compactando para: $ZipPath"
if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
Compress-Archive -Path (Join-Path $CloneDir "*") -DestinationPath $ZipPath -CompressionLevel Optimal -Force
Write-Host "ZIP gerado."

# =========================
# 4) Baixar e instalar rclone
# =========================
if (-not (Test-Path $rcloneExe)) {
    Write-Host "Baixando rclone..."
    $url = "https://downloads.rclone.org/rclone-current-windows-amd64.zip"
    Invoke-WebRequest -Uri $url -OutFile $rcloneZip

    Write-Host "Extraindo rclone..."
    # Expand-Archive mantém a estrutura; copiar apenas o rclone.exe
    $tempExtract = Join-Path $ToolsDir "extract-$Stamp"
    New-Item -Path $tempExtract -ItemType Directory -Force | Out-Null
    Expand-Archive -Path $rcloneZip -DestinationPath $tempExtract -Force

    $exeFound = Get-ChildItem -Path $tempExtract -Recurse -Filter "rclone.exe" | Select-Object -First 1
    if (-not $exeFound) { throw "rclone.exe não encontrado dentro do zip." }
    Copy-Item $exeFound.FullName $rcloneExe -Force
    Remove-Item $tempExtract -Recurse -Force
    Write-Host "rclone instalado em $rcloneExe"
}

# Opcional: colocar na PATH da sessão
$env:Path = "$ToolsDir;$env:Path"

# =========================
# 5) Configurar rclone com o token fornecido
# =========================
# Vamos criar/atualizar a seção do remoto no rclone.conf
# Para OneDrive corporativo, o tipo é "onedrive" e drive_type "business"
# Evita diálogo interativo usando arquivo de config.
Write-Host "Configurando remoto $RemoteName no rclone..."

# Carrega config existente (se houver)
$rconf = if (Test-Path $RcloneConfig) { Get-Content $RcloneConfig -Raw } else { "" }

# Remove bloco antigo do mesmo remoto
if ($rconf -match "^\[$RemoteName\](?:\r?\n.*?)(?=^\[|\Z)" -im) {
    $rconf = [regex]::Replace($rconf, "^\[$RemoteName\](?:\r?\n.*?)(?=^\[|\Z)", "", "Singleline, IgnoreCase, Multiline")
}

# Adiciona novo bloco
$block = @"
[$RemoteName]
type = onedrive
drive_type = business
token = $TokenJson
"@

$rconf = ($rconf.Trim() + "`r`n`r`n" + $block.Trim() + "`r`n")

# Garante diretório e grava
New-Item -Path $RcloneConfigDir -ItemType Directory -Force | Out-Null
Set-Content -Path $RcloneConfig -Value $rconf -Encoding UTF8

Write-Host "Remoto $RemoteName configurado em: $RcloneConfig"
Write-Host "Teste rápido de 'rclone about' (pode demorar alguns segundos e falhar se o token já estiver expirado):"
try {
    & $rcloneExe about "$RemoteName:" | Write-Host
} catch {
    Write-Warning "Falha no 'rclone about'. Verifique se o token ainda está válido."
}

# =========================
# 6) Saída final
# =========================
Write-Host ""
Write-Host "=== RESUMO ==="
Write-Host ("Perfil origem: {0}" -f $ProfileSrc)
Write-Host ("Clone em:      {0}" -f $CloneDir)
Write-Host ("ZIP em:        {0}" -f $ZipPath)
Write-Host ("rclone:        {0}" -f $rcloneExe)
Write-Host ("Config:        {0}" -f $RcloneConfig)
Write-Host ("Remoto:        {0} (onedrive / business)" -f $RemoteName)
Write-Host "Pronto."

