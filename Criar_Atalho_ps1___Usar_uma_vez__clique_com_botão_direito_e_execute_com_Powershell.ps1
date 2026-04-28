# =========================================
# CRIAR ATALHO NO DESKTOP - ORGANIZADOR DOCS
# =========================================

# FIX CRITICO: $PSScriptRoot fica vazio ao executar via clique direito
# em unidades de rede. Usamos multiplos fallbacks em ordem de confiabilidade.
$ScriptRoot = $null

# Tentativa 1 — $PSScriptRoot (confiavel em PS5+ chamado via -File)
if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $ScriptRoot = $PSScriptRoot
}

# Tentativa 2 — caminho direto do comando em execucao
if ([string]::IsNullOrWhiteSpace($ScriptRoot)) {
    try {
        $def = $MyInvocation.MyCommand.Definition
        if (-not [string]::IsNullOrWhiteSpace($def) -and (Test-Path $def)) {
            $ScriptRoot = Split-Path -Parent $def
        }
    } catch {}
}

# Tentativa 3 — MyCommand.Path
if ([string]::IsNullOrWhiteSpace($ScriptRoot)) {
    try {
        $p = $MyInvocation.MyCommand.Path
        if (-not [string]::IsNullOrWhiteSpace($p) -and (Test-Path $p)) {
            $ScriptRoot = Split-Path -Parent $p
        }
    } catch {}
}

# Tentativa 4 — diretorio atual (ultimo recurso)
if ([string]::IsNullOrWhiteSpace($ScriptRoot)) {
    $ScriptRoot = (Get-Location).Path
}

Write-Host ""
Write-Host "  Pasta detectada: $ScriptRoot" -ForegroundColor DarkGray

$ScriptMenu = Join-Path $ScriptRoot "Menu_Operadoras.ps1"

if (-not (Test-Path $ScriptMenu)) {
    Write-Host ""
    Write-Host "  ERRO: Arquivo 'Menu_Operadoras.ps1' nao encontrado em:" -ForegroundColor Red
    Write-Host "    $ScriptRoot" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Solucoes possiveis:" -ForegroundColor White
    Write-Host "  1) Certifique-se de que todos os arquivos estao na mesma pasta" -ForegroundColor Gray
    Write-Host "  2) Abra o PowerShell manualmente, navegue ate a pasta e execute:" -ForegroundColor Gray
    Write-Host "     .\Criar_Atalho.ps1" -ForegroundColor Cyan
    Write-Host ""
    Read-Host "  Pressione ENTER para sair"
    exit 1
}

$Desktop = [Environment]::GetFolderPath("Desktop")
$Atalho  = Join-Path $Desktop "Organizar Docs Clientes.lnk"

# FIX: Usar caminho completo do powershell.exe — evita falhas quando
#      o PATH do sistema nao esta disponivel no contexto do atalho
$PsExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
if (-not (Test-Path $PsExe)) {
    $PsExe = (Get-Command powershell.exe -ErrorAction SilentlyContinue).Source
}
if ([string]::IsNullOrWhiteSpace($PsExe)) {
    $PsExe = "powershell.exe"
}

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($Atalho)

$Shortcut.TargetPath       = $PsExe
$Shortcut.Arguments        = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptMenu`""
$Shortcut.WorkingDirectory = $ScriptRoot
$Shortcut.IconLocation     = "$PsExe,0"
$Shortcut.WindowStyle      = 1
$Shortcut.Description      = "Organizador de Documentos de Clientes"

$Shortcut.Save()

Write-Host ""
Write-Host "  OK Atalho criado com sucesso no Desktop!" -ForegroundColor Green
Write-Host "     Local  : $Atalho" -ForegroundColor Gray
Write-Host "     Script : $ScriptMenu" -ForegroundColor Gray
Write-Host "     PS.exe : $PsExe" -ForegroundColor Gray
Write-Host ""
Read-Host "  Pressione ENTER para sair"
