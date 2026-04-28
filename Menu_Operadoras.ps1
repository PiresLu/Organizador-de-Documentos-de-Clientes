# ================= CONFIGURACAO =================
$Raiz           = "T:\Cadastro\Grupo Regional 3"
$ScriptOperacao = "$PSScriptRoot\Organizar_DOCS_Clientes.ps1"
# ================================================

if (-not (Test-Path $ScriptOperacao)) {
    Write-Host ""
    Write-Host "  ERRO: Script de operacao nao encontrado:" -ForegroundColor Red
    Write-Host "    $ScriptOperacao" -ForegroundColor Yellow
    Write-Host ""
    Pause; exit 1
}

# ----------------------------------------------------------------
# Funcoes
# ----------------------------------------------------------------
function Get-Operadoras {
    $lista = Get-ChildItem $Raiz -Directory -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -match '^\d{4}\s+-\s+.+$' } |
             Sort-Object Name
    if (-not $lista -or @($lista).Count -eq 0) {
        Write-Host ""
        Write-Host "  Nenhuma operadora encontrada em: $Raiz" -ForegroundColor Red
        Pause; return $null
    }
    return $lista
}

function Show-MenuPrincipal {
    Clear-Host
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "         ORGANIZADOR DE DOCUMENTOS - OPERADORAS                " -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  -- Operadora individual ----------------------------------------" -ForegroundColor DarkGray
    Write-Host "  [1]  Selecionar operadora  >>  Somente pasta 'Enviados'  (padrao)" -ForegroundColor White
    Write-Host "  [2]  Selecionar operadora  >>  Varredura Completa               " -ForegroundColor White
    Write-Host "  [3]  Selecionar operadora  >>  Simulacao (sem alteracoes)       " -ForegroundColor Magenta
    Write-Host ""
    Write-Host "  -- Todas as operadoras -----------------------------------------" -ForegroundColor DarkGray
    Write-Host "  [4]  Todas as operadoras   >>  Somente pasta 'Enviados'  (padrao)" -ForegroundColor Yellow
    Write-Host "  [5]  Todas as operadoras   >>  Varredura Completa               " -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [0]  Sair" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGray
    return (Read-Host "  Escolha uma opcao").Trim()
}

function Show-MenuOperadora {
    param ([object[]]$Operadoras)
    Clear-Host
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "                 SELECIONE A OPERADORA                         " -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    for ($i = 0; $i -lt $Operadoras.Count; $i++) {
        Write-Host ("  [{0,3}]  {1}" -f ($i + 1), $Operadoras[$i].Name)
    }
    Write-Host ""
    Write-Host "  [  0]  Voltar" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGray
    return (Read-Host "  Escolha a operadora").Trim()
}

# Processa uma unica operadora
# -PularMesclagem: usado no modo "todas" para adiar a Etapa 3 ao final
function Invoke-Processar {
    param (
        [string]$Caminho,
        [string]$Modo,
        [switch]$Sim,
        [switch]$PularMesclagem
    )
    Write-Host ""
    Write-Host "  > Operadora : " -NoNewline -ForegroundColor Gray
    Write-Host (Split-Path $Caminho -Leaf) -ForegroundColor Cyan
    Write-Host "  > Modo      : $Modo" -ForegroundColor Gray
    if ($Sim)            { Write-Host "  > SIMULACAO : Nenhuma alteracao sera feita" -ForegroundColor Magenta }
    if ($PularMesclagem) { Write-Host "  > Mesclagem adiada para execucao unica no final" -ForegroundColor DarkYellow }
    Write-Host ""

    $args = @(
        "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", $ScriptOperacao,
        "-DiretorioOperadora", $Caminho,
        "-ModoVarredura", $Modo
    )
    if ($Sim)            { $args += "-Simulacao" }
    if ($PularMesclagem) { $args += "-PularMesclagem" }

    powershell.exe @args

    Write-Host ""
    Write-Host ("  OK Concluido: {0}" -f (Split-Path $Caminho -Leaf)) -ForegroundColor Green
}

# Executa apenas a Etapa 3 (mesclagem) uma unica vez no final
# passando o diretorio raiz como referencia (nao e usado na Etapa 3,
# que sempre varre R:\ diretamente)
function Invoke-MeslagemFinal {
    param ([switch]$Sim)
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "   MESCLAGEM FINAL - Consolidando duplicatas em R:\            " -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Executando varredura unica de duplicatas em R:\" -ForegroundColor White
    Write-Host ""

    $args = @(
        "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", $ScriptOperacao,
        "-DiretorioOperadora", $Raiz,
        "-SomenteMesclagem"
    )
    if ($Sim) { $args += "-Simulacao" }

    powershell.exe @args
}

# ----------------------------------------------------------------
# Loop principal
# ----------------------------------------------------------------
while ($true) {

    $opcao = Show-MenuPrincipal

    switch ($opcao) {

        "0" { exit }

        # Operadora individual — mesclagem incluida normalmente
        { $_ -in "1","2","3" } {
            $operadoras = Get-Operadoras
            if (-not $operadoras) { break }

            $sel = Show-MenuOperadora -Operadoras $operadoras
            if ($sel -eq "0") { break }

            [int]$idx = 0
            if (-not [int]::TryParse($sel, [ref]$idx) -or $idx -lt 1 -or $idx -gt $operadoras.Count) {
                Write-Host ""; Write-Host "  Opcao invalida." -ForegroundColor Red; Start-Sleep 2; break
            }

            $modo = switch ($opcao) { "2" { "Completa" } default { "Enviados" } }
            $sim  = ($opcao -eq "3")
            Invoke-Processar -Caminho $operadoras[$idx - 1].FullName -Modo $modo -Sim:$sim
            Pause
        }

        # Todas as operadoras — mesclagem executada UMA UNICA VEZ no final
        { $_ -in "4","5" } {
            $operadoras = Get-Operadoras
            if (-not $operadoras) { break }

            $modo  = if ($opcao -eq "5") { "Completa" } else { "Enviados" }
            $aviso = if ($opcao -eq "5") { "  ATENCAO: Varredura Completa selecionada!" } else { "" }

            Write-Host ""
            if ($aviso) { Write-Host $aviso -ForegroundColor Red }
            Write-Host ("  Processar TODAS as {0} operadora(s) no modo: {1}" -f $operadoras.Count, $modo) -ForegroundColor Yellow
            Write-Host "  A busca por duplicatas sera feita UMA UNICA VEZ ao final." -ForegroundColor DarkYellow
            Write-Host ""
            $confirm = Read-Host "  Confirma? (S = Sim / qualquer tecla = Cancelar)"
            if ($confirm -notin "S","s") {
                Write-Host "  Cancelado." -ForegroundColor DarkGray; Start-Sleep 1; break
            }

            $total = $operadoras.Count
            $atual = 0

            # Etapas 1, 2 e 4 para cada operadora (mesclagem pulada)
            foreach ($op in $operadoras) {
                $atual++
                Write-Host ""
                Write-Host ("  == [{0}/{1}] {2}" -f $atual, $total, $op.Name) -ForegroundColor Cyan
                Invoke-Processar -Caminho $op.FullName -Modo $modo -PularMesclagem
            }

            # Etapa 3 — mesclagem unica apos processar todas as operadoras
            Invoke-MeslagemFinal

            Write-Host ""
            Write-Host "  OK Todas as $total operadoras foram processadas." -ForegroundColor Green
            Write-Host "  OK Mesclagem final concluida."                    -ForegroundColor Green
            Pause
        }

        default {
            Write-Host ""; Write-Host "  Opcao invalida." -ForegroundColor Red; Start-Sleep 1
        }
    }
}
