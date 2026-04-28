<#
.SYNOPSIS
    Organizador de Documentos de Clientes v4.1 - MODO SEGURO

.DESCRIPTION
    Organiza PDFs de clientes (6 digitos) em pastas, move para R:\
    e mescla duplicadas usando SEMPRE a pasta mais antiga como principal.

    SEGURANCA - NENHUM ARQUIVO/PASTA EH EXCLUIDO PERMANENTEMENTE:
    - Tudo que seria "removido" vai para R:\_LIXEIRA_SCRIPT\<data>\<operadora>\
    - A Lixeira do Windows NAO eh usada (nao funciona em unidades de rede T:\ R:\)
    - Um LOG completo eh gravado a cada execucao em R:\_LIXEIRA_SCRIPT\Logs\
    - Modo -Simulacao disponivel: mostra o que SERIA feito, sem alterar nada

    REGRAS:
    - Apenas PDFs com padrao 000000 - nome.pdf (6 digitos)
    - Ignora PDFs com 5 digitos (contratos)
    - Pastas com (cancelamento), (DEP) ou DEP isolado sao secundarias
    - Ao mesclar duplicadas: a pasta MAIS ANTIGA e sempre a principal
    - Modo "Enviados": processa apenas documentos dentro da pasta "Enviados"
    - Modo "Completa": processa todos os documentos sem restricao de pasta

    PARAMETROS DE CONTROLE DE FLUXO:
    - PularMesclagem  : Pula a Etapa 3 (mesclagem). Usado quando o menu
                        processa todas as operadoras e faz a mesclagem
                        uma unica vez no final, evitando varreduras repetidas.
    - SomenteMesclagem: Executa APENAS a Etapa 3 (mesclagem em R:\).
                        Usado pelo menu apos processar todas as operadoras.

.PARAMETER DiretorioOperadora
    Diretorio da operadora. Obrigatorio exceto no modo SomenteMesclagem.

.PARAMETER ModoVarredura
    "Enviados" (padrao) ou "Completa".

.PARAMETER Simulacao
    Exibe o que SERIA feito. Nenhum arquivo e alterado.

.PARAMETER PularMesclagem
    Pula a Etapa 3 (mesclagem de duplicatas em R:\).

.PARAMETER SomenteMesclagem
    Executa apenas a Etapa 3. Ignora DiretorioOperadora.

.EXAMPLE
    # Uso normal - pasta Enviados
    .\Organizar_DOCS_Clientes.ps1 -DiretorioOperadora "T:\...\0041 - Jau"

    # Varredura completa
    .\Organizar_DOCS_Clientes.ps1 -DiretorioOperadora "T:\...\0041 - Jau" -ModoVarredura Completa

    # Sem mesclagem (usado pelo menu no modo "todas as operadoras")
    .\Organizar_DOCS_Clientes.ps1 -DiretorioOperadora "T:\...\0041 - Jau" -PularMesclagem

    # Apenas mesclagem final (chamado uma unica vez pelo menu)
    .\Organizar_DOCS_Clientes.ps1 -DiretorioOperadora "T:\..." -SomenteMesclagem

    # Simulacao
    .\Organizar_DOCS_Clientes.ps1 -DiretorioOperadora "T:\...\0041 - Jau" -Simulacao
#>

param (
    [Parameter(Mandatory = $false)]
    [string]$DiretorioOperadora = "",

    [Parameter(Mandatory = $false)]
    [ValidateSet("Enviados", "Completa")]
    [string]$ModoVarredura = "Enviados",

    [Parameter(Mandatory = $false)]
    [switch]$Simulacao,

    [Parameter(Mandatory = $false)]
    [switch]$PularMesclagem,

    [Parameter(Mandatory = $false)]
    [switch]$SomenteMesclagem
)

# Validar parametros essenciais
if (-not $SomenteMesclagem -and [string]::IsNullOrWhiteSpace($DiretorioOperadora)) {
    Write-Host "  ERRO: -DiretorioOperadora e obrigatorio." -ForegroundColor Red
    exit 1
}

# =====================================================================
# CONFIGURACOES GLOBAIS
# =====================================================================
$DestinoFinal      = "R:\"
$NomePastaEnviados = "Enviados"
$NomePastaBackup   = "_LIXEIRA_SCRIPT"
$Timestamp         = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$NomeOperadora     = if ($SomenteMesclagem) { "MeslagemFinal" } else { Split-Path $DiretorioOperadora -Leaf }

$PastaBackupExecucao = Join-Path (Join-Path (Join-Path $DestinoFinal $NomePastaBackup) $Timestamp) $NomeOperadora
$PastaLog            = Join-Path (Join-Path $DestinoFinal $NomePastaBackup) "Logs"
$ArquivoLog          = Join-Path $PastaLog ("LOG_{0}_{1}.txt" -f $Timestamp, ($NomeOperadora -replace '[\\/:*?"<>|]', '_'))

# Regex
$RegexPdfCliente   = '^\d{6}\s+-\s+.+\.pdf$'
$RegexPastaCliente = '^\d{6}\s+-\s+.+'
$RegexCodigo       = '^(\d{6})'

$IdentificadoresSecundarios = @('(?i)\(cancelamento\)', '(?i)\(DEP\)', '(?i)\bDEP\b')

# Contadores
$script:ContadorOrganizados = 0
$script:ContadorMovidas     = 0
$script:ContadorMesclados   = 0
$script:ContadorBackup      = 0
$script:ContadorVazias      = 0
$script:Erros               = [System.Collections.Generic.List[string]]::new()
$script:LogLinhas           = [System.Collections.Generic.List[string]]::new()

# =====================================================================
# FUNCOES
# =====================================================================

function Write-Log {
    param ([string]$Mensagem, [string]$Nivel = "INFO")
    $Linha = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Nivel.PadRight(8), $Mensagem
    $script:LogLinhas.Add($Linha)
    switch ($Nivel) {
        "ERRO"      { Write-Host "  X $Mensagem" -ForegroundColor Red }
        "AVISO"     { Write-Host "  ! $Mensagem" -ForegroundColor Yellow }
        "ACAO"      { Write-Host "  > $Mensagem" -ForegroundColor Green }
        "SIMULACAO" { Write-Host "  [SIM] $Mensagem" -ForegroundColor Magenta }
        default     { Write-Host "    $Mensagem" -ForegroundColor Gray }
    }
}

function Save-Log {
    try {
        if ($Simulacao) { return }
        if (-not (Test-Path $PastaLog)) {
            New-Item -ItemType Directory -Path $PastaLog -Force | Out-Null
        }
        $Linhas = @(
            ("=" * 70),
            "  LOG DE EXECUCAO - ORGANIZADOR DE DOCUMENTOS v4.1",
            ("=" * 70),
            "  Data/Hora       : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')",
            "  Operadora       : $NomeOperadora",
            "  Modo            : $ModoVarredura",
            "  Simulacao       : $($Simulacao.IsPresent)",
            "  PularMesclagem  : $($PularMesclagem.IsPresent)",
            "  SomenteMesclagem: $($SomenteMesclagem.IsPresent)",
            "  Backup em       : $PastaBackupExecucao",
            ("=" * 70),
            ""
        ) + $script:LogLinhas.ToArray() + @(
            "",
            ("=" * 70),
            "  RESUMO",
            ("=" * 70),
            "  PDFs organizados   : $($script:ContadorOrganizados)",
            "  Pastas movidas     : $($script:ContadorMovidas)",
            "  Arquivos mesclados : $($script:ContadorMesclados)",
            "  Itens no backup    : $($script:ContadorBackup)",
            "  Pastas vazias      : $($script:ContadorVazias)",
            "  Erros              : $($script:Erros.Count)",
            ("=" * 70)
        )
        $Linhas | Set-Content -Path $ArquivoLog -Encoding UTF8
    } catch {
        Write-Host "  ! Nao foi possivel gravar o log: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}

# -----------------------------------------------------------------------
# Move para backup - substitui TODA e qualquer exclusao do script.
# NUNCA exclui permanentemente. Tudo vai para R:\_LIXEIRA_SCRIPT\
# IMPORTANTE: SendToRecycleBin NAO funciona em unidades de rede (T:\ R:\)
# -----------------------------------------------------------------------
function Move-ToBackup {
    param ([Parameter(Mandatory = $true)][string]$Path, [string]$Motivo = "")
    if ($Simulacao) {
        Write-Log "BACKUP (simulado): '$(Split-Path $Path -Leaf)'  [$Motivo]" "SIMULACAO"
        return $true
    }
    try {
        if (-not (Test-Path $Path)) {
            Write-Log "Caminho nao existe para backup: $Path" "AVISO"
            return $false
        }
        $NomeItem    = Split-Path $Path -Leaf
        $DestinoItem = Join-Path $PastaBackupExecucao $NomeItem
        if (Test-Path $DestinoItem) {
            $i = 1
            do {
                $DestinoItem = Join-Path $PastaBackupExecucao ("{0}_{1:D3}" -f $NomeItem, $i)
                $i++
            } while ((Test-Path $DestinoItem) -and $i -le 999)
        }
        if (-not (Test-Path $PastaBackupExecucao)) {
            New-Item -ItemType Directory -Path $PastaBackupExecucao -Force -ErrorAction Stop | Out-Null
        }
        Move-Item $Path -Destination $DestinoItem -Force -ErrorAction Stop
        $script:ContadorBackup++
        Write-Log "Backup: '$NomeItem'  [$Motivo]  >>  $DestinoItem" "ACAO"
        return $true
    } catch {
        $Msg = "Falha ao mover para backup '$Path': $($_.Exception.Message)"
        Write-Log $Msg "ERRO"; $script:Erros.Add($Msg)
        return $false
    }
}

function Test-PastaSecundaria {
    param ([string]$NomePasta)
    foreach ($Padrao in $IdentificadoresSecundarios) {
        if ($NomePasta -match $Padrao) { return $true }
    }
    return $false
}

function Merge-FoldersSafely {
    param (
        [Parameter(Mandatory = $true)][string]$Origem,
        [Parameter(Mandatory = $true)][string]$Destino
    )
    try {
        $Arquivos = @(Get-ChildItem $Origem -Recurse -File -ErrorAction Stop)
        $Total    = $Arquivos.Count
        Write-Log "Mesclar $Total arquivo(s): '$(Split-Path $Origem -Leaf)' >> '$(Split-Path $Destino -Leaf)'" "INFO"
        if ($Simulacao) {
            Write-Log "Mesclar $Total arquivo(s) (simulado)" "SIMULACAO"
            $script:ContadorMesclados += $Total
            return $true
        }
        foreach ($Arquivo in $Arquivos) {
            try {
                $Relativo       = $Arquivo.FullName.Substring($Origem.Length).TrimStart('\')
                $DestinoArquivo = Join-Path $Destino $Relativo
                $DestinoDir     = Split-Path $DestinoArquivo -Parent
                if (-not (Test-Path $DestinoDir)) {
                    New-Item -ItemType Directory -Path $DestinoDir -Force -ErrorAction Stop | Out-Null
                }
                if (Test-Path $DestinoArquivo) {
                    $i = 1
                    do {
                        $NovoNome       = "{0}_DUP_{1:D3}{2}" -f $Arquivo.BaseName, $i, $Arquivo.Extension
                        $DestinoArquivo = Join-Path $DestinoDir $NovoNome
                        $i++
                        if ($i -gt 999) { throw "Limite de duplicatas excedido para '$($Arquivo.Name)'" }
                    } while (Test-Path $DestinoArquivo)
                    Write-Log "Conflito de nome - renomeado: $NovoNome" "AVISO"
                }
                Move-Item $Arquivo.FullName -Destination $DestinoArquivo -Force -ErrorAction Stop
                $script:ContadorMesclados++
            } catch {
                $Msg = "Erro ao mover '$($Arquivo.Name)': $($_.Exception.Message)"
                Write-Log $Msg "ERRO"; $script:Erros.Add($Msg)
            }
        }
        return $true
    } catch {
        $Msg = "Erro ao mesclar: $($_.Exception.Message)"
        Write-Log $Msg "ERRO"; $script:Erros.Add($Msg)
        return $false
    }
}

# =====================================================================
# CABECALHO
# =====================================================================
Clear-Host
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "     ORGANIZADOR DE DOCUMENTOS DE CLIENTES v4.1                " -ForegroundColor Cyan
if ($Simulacao)        { Write-Host "          *** MODO SIMULACAO - ZERO ALTERACOES ***             " -ForegroundColor Magenta }
if ($SomenteMesclagem) { Write-Host "              *** MESCLAGEM FINAL UNICA ***                    " -ForegroundColor Yellow }
if ($PularMesclagem)   { Write-Host "       (Mesclagem adiada - sera feita ao final do lote)        " -ForegroundColor DarkYellow }
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

if ($SomenteMesclagem) {
    Write-Host "  Modo       : Somente Mesclagem Final em R:\" -ForegroundColor Yellow
} else {
    Write-Host "  Operadora  : " -NoNewline -ForegroundColor Yellow; Write-Host $DiretorioOperadora -ForegroundColor White
    Write-Host "  Modo       : " -NoNewline -ForegroundColor Yellow
    if ($ModoVarredura -eq "Enviados") {
        Write-Host "Somente pasta 'Enviados'" -ForegroundColor Green
    } else {
        Write-Host "Varredura Completa" -ForegroundColor Red
    }
}
Write-Host "  Backup em  : " -NoNewline -ForegroundColor Yellow; Write-Host $PastaBackupExecucao -ForegroundColor DarkGray
Write-Host "  Log em     : " -NoNewline -ForegroundColor Yellow; Write-Host $ArquivoLog -ForegroundColor DarkGray
Write-Host ""
Write-Log "Inicio | Operadora: $NomeOperadora | Modo: $ModoVarredura | Sim: $($Simulacao.IsPresent) | PularMescl: $($PularMesclagem.IsPresent) | SomenteMescl: $($SomenteMesclagem.IsPresent)" "INFO"

# Validacoes
if (-not $SomenteMesclagem -and -not (Test-Path $DiretorioOperadora)) {
    Write-Log "Diretorio de origem nao encontrado: $DiretorioOperadora" "ERRO"; Save-Log; exit 1
}
if (-not (Test-Path $DestinoFinal)) {
    Write-Log "Unidade de destino nao encontrada: $DestinoFinal" "ERRO"; Save-Log; exit 1
}

# =====================================================================
# ETAPAS 1, 2 e 4 — Ignoradas no modo SomenteMesclagem
# =====================================================================
if (-not $SomenteMesclagem) {

    # -------------------------------------------------------------------
    # ETAPA 1 - ORGANIZAR PDFs EM PASTAS
    # -------------------------------------------------------------------
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " [1/4] ORGANIZANDO PDFs EM PASTAS" -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Log "--- ETAPA 1: Organizar PDFs ---" "INFO"

    if ($ModoVarredura -eq "Enviados") {
        $PastasEnviados = @(Get-ChildItem $DiretorioOperadora -Recurse -Directory -ErrorAction SilentlyContinue |
                            Where-Object { $_.Name -ieq $NomePastaEnviados })
        if ($PastasEnviados.Count -eq 0) {
            Write-Log "Nenhuma pasta '$NomePastaEnviados' encontrada. Use Varredura Completa se necessario." "AVISO"
        }
        $TodosPdfs = @(foreach ($PE in $PastasEnviados) {
            Get-ChildItem $PE.FullName -File -Filter "*.pdf" -ErrorAction SilentlyContinue
        })
    } else {
        $TodosPdfs = @(Get-ChildItem $DiretorioOperadora -Recurse -File -Filter "*.pdf" -ErrorAction SilentlyContinue)
    }

    $PdfsClientes  = @($TodosPdfs | Where-Object { $_.Name -match $RegexPdfCliente })
    $PdfsContratos = @($TodosPdfs | Where-Object { $_.Name -match '^\d{5}\s+-\s+.+\.pdf$' })

    Write-Host "  PDFs encontrados              : $($TodosPdfs.Count)" -ForegroundColor White
    Write-Host "  PDFs de clientes (6 digitos)  : $($PdfsClientes.Count)" -ForegroundColor Green
    if ($PdfsContratos.Count -gt 0) {
        Write-Host "  PDFs de contratos (ignorados) : $($PdfsContratos.Count)" -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Log "PDFs total: $($TodosPdfs.Count) | Clientes: $($PdfsClientes.Count) | Contratos: $($PdfsContratos.Count)" "INFO"

    foreach ($Pdf in $PdfsClientes) {
        try {
            if ($Pdf.Directory.Name -eq $Pdf.BaseName) { continue }
            $PastaDestino   = Join-Path $Pdf.Directory.FullName $Pdf.BaseName
            $ArquivoDestino = Join-Path $PastaDestino $Pdf.Name
            if ($Simulacao) {
                Write-Log "Organizar: '$($Pdf.Name)' >> '$PastaDestino'" "SIMULACAO"
                $script:ContadorOrganizados++; continue
            }
            if (-not (Test-Path $PastaDestino)) { New-Item -ItemType Directory -Path $PastaDestino -ErrorAction Stop | Out-Null }
            if (-not (Test-Path $ArquivoDestino)) {
                Move-Item $Pdf.FullName -Destination $ArquivoDestino -Force -ErrorAction Stop
                $script:ContadorOrganizados++
                Write-Log "Organizado: '$($Pdf.Name)'" "ACAO"
            }
        } catch {
            $Msg = "Erro ao organizar '$($Pdf.Name)': $($_.Exception.Message)"
            Write-Log $Msg "ERRO"; $script:Erros.Add($Msg)
        }
    }
    Write-Host ""
    Write-Host "  OK PDFs organizados: $($script:ContadorOrganizados)" -ForegroundColor Green
    Write-Host ""

    # -------------------------------------------------------------------
    # ETAPA 2 - MOVER PASTAS PARA R:\
    # -------------------------------------------------------------------
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " [2/4] MOVENDO PASTAS DE CLIENTES PARA R:\" -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Log "--- ETAPA 2: Mover pastas para R:\ ---" "INFO"

    if ($ModoVarredura -eq "Enviados") {
        Write-Host "  Modo: Somente pasta '$NomePastaEnviados'" -ForegroundColor Green
        $PastasEnviados = @(Get-ChildItem $DiretorioOperadora -Recurse -Directory -ErrorAction SilentlyContinue |
                            Where-Object { $_.Name -ieq $NomePastaEnviados })
        $PastasParaMover = @(foreach ($PE in $PastasEnviados) {
            Get-ChildItem $PE.FullName -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match $RegexPastaCliente }
        })
    } else {
        Write-Host "  Modo: Varredura Completa" -ForegroundColor Red
        $PastasParaMover = @(Get-ChildItem $DiretorioOperadora -Recurse -Directory -ErrorAction SilentlyContinue |
                             Where-Object { $_.Name -match $RegexPastaCliente })
    }

    $PastasParaMover = $PastasParaMover | Sort-Object { $_.FullName.Split('\').Count } -Descending
    Write-Host "  Pastas de clientes encontradas: $($PastasParaMover.Count)" -ForegroundColor White
    Write-Host ""
    Write-Log "Pastas para mover: $($PastasParaMover.Count)" "INFO"

    foreach ($Pasta in $PastasParaMover) {
        if (-not (Test-Path $Pasta.FullName)) { continue }
        try {
            $DestinoCliente = Join-Path $DestinoFinal $Pasta.Name
            if ($Simulacao) {
                $acao = if (Test-Path $DestinoCliente) { "Mover com nome temporario (conflito)" } else { "Mover para R:\" }
                Write-Log "$acao : '$($Pasta.Name)'" "SIMULACAO"
                $script:ContadorMovidas++; continue
            }
            if (Test-Path $DestinoCliente) {
                $NomeTemp    = "$($Pasta.Name)__ENTRADA_$(Get-Date -Format 'yyyyMMdd_HHmmss_fff')"
                $DestinoTemp = Join-Path $DestinoFinal $NomeTemp
                Write-Log "Conflito '$($Pasta.Name)' - movido como '$NomeTemp' para mesclagem posterior" "AVISO"
                Move-Item $Pasta.FullName -Destination $DestinoTemp -Force -ErrorAction Stop
            } else {
                Move-Item $Pasta.FullName -Destination $DestinoFinal -Force -ErrorAction Stop
                Write-Log "Movida: '$($Pasta.Name)' >> R:\" "ACAO"
            }
            $script:ContadorMovidas++
        } catch {
            $Msg = "Erro ao mover '$($Pasta.Name)': $($_.Exception.Message)"
            Write-Log $Msg "ERRO"; $script:Erros.Add($Msg)
        }
    }
    Write-Host ""
    Write-Host "  OK Pastas movidas: $($script:ContadorMovidas)" -ForegroundColor Green
    Write-Host ""

} # fim do bloco -not SomenteMesclagem para Etapas 1 e 2

# =====================================================================
# ETAPA 3 - MESCLAR DUPLICADAS
# Executada normalmente OU pulada (PularMesclagem) OU sozinha (SomenteMesclagem)
# =====================================================================
if ($PularMesclagem) {
    Write-Host "================================================================" -ForegroundColor DarkGray
    Write-Host " [3/4] MESCLAGEM PULADA - sera executada ao final do lote      " -ForegroundColor DarkYellow
    Write-Host "================================================================" -ForegroundColor DarkGray
    Write-Log "Etapa 3 pulada (PularMesclagem ativo)" "AVISO"
    Write-Host ""
} else {
    $EtapaLabel = if ($SomenteMesclagem) { " [MESCLAGEM FINAL]" } else { " [3/4]" }
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "$EtapaLabel MESCLANDO PASTAS DUPLICADAS EM R:\              " -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Log "--- ETAPA 3: Mesclar duplicadas ---" "INFO"

    $PastasR = @(Get-ChildItem $DestinoFinal -Directory -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -match $RegexCodigo -and $_.Name -ne $NomePastaBackup })

    Write-Host "  Pastas em R:\: $($PastasR.Count)" -ForegroundColor White
    Write-Log "Pastas em R:\: $($PastasR.Count)" "INFO"

    $Grupos = $PastasR | Group-Object { if ($_.Name -match $RegexCodigo) { $Matches[1] } }
    $GruposComDuplicatas = @($Grupos | Where-Object { $_.Count -gt 1 })

    if ($GruposComDuplicatas.Count -eq 0) {
        Write-Log "Nenhuma duplicata encontrada" "INFO"
        Write-Host "  OK Nenhuma duplicata encontrada" -ForegroundColor Green
    } else {
        Write-Host "  ! Grupos com duplicatas: $($GruposComDuplicatas.Count)" -ForegroundColor Yellow
        Write-Log "Grupos com duplicatas: $($GruposComDuplicatas.Count)" "AVISO"
        Write-Host ""

        foreach ($Grupo in $GruposComDuplicatas) {
            Write-Host "  ----------------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Codigo: $($Grupo.Name)  |  $($Grupo.Count) pastas" -ForegroundColor Cyan
            Write-Log "Grupo $($Grupo.Name): $($Grupo.Count) pastas" "INFO"
            Write-Host ""

            $Principais  = @()
            $Secundarias = @()

            foreach ($Pasta in $Grupo.Group) {
                if (Test-PastaSecundaria $Pasta.Name) {
                    $Secundarias += $Pasta
                    Write-Host "    [SECUNDARIA] $($Pasta.Name)  ($($Pasta.CreationTime.ToString('dd/MM/yyyy HH:mm')))" -ForegroundColor DarkYellow
                    Write-Log "Secundaria: $($Pasta.Name)" "INFO"
                } else {
                    $Principais += $Pasta
                    Write-Host "    [PRINCIPAL ] $($Pasta.Name)  ($($Pasta.CreationTime.ToString('dd/MM/yyyy HH:mm')))" -ForegroundColor Green
                    Write-Log "Principal candidata: $($Pasta.Name)" "INFO"
                }
            }
            Write-Host ""

            # REGRA: pasta MAIS ANTIGA e sempre a principal
            if ($Principais.Count -gt 0) {
                $PastaPrincipal = $Principais | Sort-Object CreationTime | Select-Object -First 1
            } else {
                $PastaPrincipal = $Grupo.Group | Sort-Object CreationTime | Select-Object -First 1
                Write-Log "Nenhuma principal - usando mais antiga: $($PastaPrincipal.Name)" "AVISO"
                Write-Host "    ! Sem pasta principal identificada - usando a mais antiga" -ForegroundColor Yellow
            }

            Write-Host "    >> Principal: $($PastaPrincipal.Name)" -ForegroundColor Cyan
            Write-Host "       Criada em: $($PastaPrincipal.CreationTime.ToString('dd/MM/yyyy HH:mm:ss'))" -ForegroundColor DarkGray
            Write-Log "Principal definida: $($PastaPrincipal.Name) | Criada: $($PastaPrincipal.CreationTime)" "ACAO"
            Write-Host ""

            foreach ($Pasta in $Grupo.Group) {
                if ($Pasta.FullName -eq $PastaPrincipal.FullName) { continue }
                $Mesclado = Merge-FoldersSafely -Origem $Pasta.FullName -Destino $PastaPrincipal.FullName
                if ($Mesclado) {
                    $ok = Move-ToBackup -Path $Pasta.FullName -Motivo "Pasta secundaria apos mesclagem"
                    if ($ok) { Write-Host "    OK Pasta movida para backup: $($Pasta.Name)" -ForegroundColor Green }
                }
                Write-Host ""
            }
        }
    }
    Write-Host ""
    Write-Host "  OK Mesclagem concluida" -ForegroundColor Green
    Write-Host ""
}

# =====================================================================
# ETAPA 4 - LIMPAR PASTAS VAZIAS (ignorada no modo SomenteMesclagem)
# =====================================================================
if (-not $SomenteMesclagem) {
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " [4/4] LIMPANDO PASTAS VAZIAS" -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Log "--- ETAPA 4: Limpar pastas vazias ---" "INFO"

    $PastasParaLimpar = Get-ChildItem $DiretorioOperadora -Recurse -Directory -ErrorAction SilentlyContinue |
                        Sort-Object { $_.FullName.Length } -Descending

    foreach ($Pasta in $PastasParaLimpar) {
        try {
            if ($Pasta.Name -ieq $NomePastaEnviados) { continue }
            if (-not (Test-Path $Pasta.FullName)) { continue }
            $Conteudo = @(Get-ChildItem $Pasta.FullName -Force -ErrorAction SilentlyContinue)
            if ($null -eq $Conteudo -or $Conteudo.Count -eq 0) {
                $ok = Move-ToBackup -Path $Pasta.FullName -Motivo "Pasta vazia"
                if ($ok) {
                    $script:ContadorVazias++
                    Write-Host "  OK Pasta vazia movida para backup: $($Pasta.Name)" -ForegroundColor DarkGray
                }
            }
        } catch { }
    }
    Write-Host ""
    Write-Host "  OK Pastas vazias processadas: $($script:ContadorVazias)" -ForegroundColor Green
}

# =====================================================================
# SUMARIO FINAL E LOG
# =====================================================================
Save-Log

$CorBorda = if ($script:Erros.Count -gt 0) { "Yellow" } else { "Green" }

Write-Host ""
Write-Host "================================================================" -ForegroundColor $CorBorda
if ($Simulacao)        { Write-Host "    SIMULACAO CONCLUIDA - NENHUMA ALTERACAO REALIZADA          " -ForegroundColor Magenta }
elseif ($SomenteMesclagem) { Write-Host "            MESCLAGEM FINAL CONCLUIDA                         " -ForegroundColor $CorBorda }
else                   { Write-Host "              PROCESSO CONCLUIDO COM SUCESSO                   " -ForegroundColor $CorBorda }
Write-Host "================================================================" -ForegroundColor $CorBorda
Write-Host ""
if (-not $SomenteMesclagem) {
    Write-Host ("  PDFs organizados em pastas : {0}" -f $script:ContadorOrganizados) -ForegroundColor White
    Write-Host ("  Pastas movidas para R:\    : {0}" -f $script:ContadorMovidas)     -ForegroundColor White
}
Write-Host ("  Arquivos mesclados         : {0}" -f $script:ContadorMesclados)   -ForegroundColor White
if (-not $SomenteMesclagem) {
    Write-Host ("  Pastas vazias processadas  : {0}" -f $script:ContadorVazias)   -ForegroundColor White
}
if ($script:ContadorBackup -gt 0) {
    Write-Host ("  Itens movidos para backup  : {0}" -f $script:ContadorBackup) -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Backup disponivel em:" -ForegroundColor DarkYellow
    Write-Host "  $PastaBackupExecucao"  -ForegroundColor DarkYellow
}
if ($script:Erros.Count -gt 0) {
    Write-Host ("  ! Erros registrados        : {0}  (consulte o log)" -f $script:Erros.Count) -ForegroundColor Red
}
if (-not $Simulacao) {
    Write-Host ""
    Write-Host "  Log completo em:" -ForegroundColor DarkGray
    Write-Host "  $ArquivoLog" -ForegroundColor DarkGray
}
Write-Host ""
if ($SomenteMesclagem) {
    Write-Host "Pressione ENTER para continuar..." -ForegroundColor Gray
} else {
    Write-Host "Pressione ENTER para sair..." -ForegroundColor Gray
}
Read-Host
