$ErrorActionPreference = "Stop"

$repo = "C:\Users\ivferreira\Documents\GitHub\produtividade-t3-seco"
$jsonPath = Join-Path $repo "dados_produtividade_individual.json"

function Get-Text($value) {
    if ($null -eq $value) { return "" }
    return ([string]$value).Trim()
}

function Get-Number($value, $default = 0) {
    if ($null -eq $value -or $value -eq "") { return $default }
    try { return [double]$value } catch { return $default }
}

function Find-LabelCell($sheet, [string]$needle, [int]$maxRow = 12, [int]$maxCol = 30) {
    $target = $needle.ToUpperInvariant()
    for ($r = 1; $r -le $maxRow; $r++) {
        for ($c = 1; $c -le $maxCol; $c++) {
            $v = Get-Text $sheet.Cells.Item($r,$c).Value2
            if ($v -and $v.ToUpperInvariant().Contains($target)) {
                return @($r,$c)
            }
        }
    }
    return $null
}

function Get-ValueBelowLabel($sheet, [string]$label, [int]$offsetRows, $fallbackAddress = $null) {
    $p = Find-LabelCell $sheet $label
    if ($p) {
        return $sheet.Cells.Item($p[0] + $offsetRows, $p[1]).Value2
    }
    if ($fallbackAddress) {
        return $sheet.Range($fallbackAddress).Value2
    }
    return $null
}

$excel = $null
$wb = $null

try {
    # IMPORTANTE: só usa o Excel já aberto. Não abre nenhuma planilha.
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
        throw "Abra primeiro a planilha NOVA no Excel (a que contém a aba DASH_INDIVIDUAL) e execute novamente."
    }

    # Procura somente um workbook ABERTO que contenha DASH_INDIVIDUAL.
    foreach ($book in $excel.Workbooks) {
        try {
            $dashTest = $book.Worksheets.Item("DASH_INDIVIDUAL")
            if ($dashTest) {
                $wb = $book
                break
            }
        } catch {}
    }

    if (-not $wb) {
        throw "Nenhuma planilha aberta contém a aba DASH_INDIVIDUAL. Abra a planilha nova e tente novamente."
    }

    try { $excel.Calculate() } catch {}

    $dash = $wb.Worksheets.Item("DASH_INDIVIDUAL")
    $t3 = $wb.Worksheets.Item("T3")

    # ---------- Cabeçalho e KPIs ----------
    $dataRaw = Get-ValueBelowLabel $dash "DATA" 1 "A5"
    if ($dataRaw -is [double] -or $dataRaw -is [int]) {
        $dataTexto = [DateTime]::FromOADate([double]$dataRaw).ToString("dd/MM/yyyy")
    } else {
        $dataTexto = Get-Text $dataRaw
    }

    $turno = Get-Text (Get-ValueBelowLabel $dash "TURNO" 1 "D5")
    $setor = Get-Text (Get-ValueBelowLabel $dash "SETOR" 1 "F5")
    $modalidade = Get-Text (Get-ValueBelowLabel $dash "MODALIDADE" 1 "I5")
    $colaboradorFiltro = Get-Text (Get-ValueBelowLabel $dash "COLABORADOR" 1 "M5")

    $staffPlanejado = [int](Get-Number (Get-ValueBelowLabel $dash "PLANEJADO" 2 "A9"))
    $presentes = [int](Get-Number (Get-ValueBelowLabel $dash "PRESENTES" 2 "D9"))
    $ausentes = [int](Get-Number (Get-ValueBelowLabel $dash "AUSENTES" 2 "G9"))
    $presenca = Get-Number (Get-ValueBelowLabel $dash "% PRESENÇA" 2 "J9")
    # Produção real será calculada pela soma dos colaboradores para evitar leitura incorreta de célula mesclada.
    $producaoReal = 0
    $projecaoTurno = Get-Number (Get-ValueBelowLabel $dash "PROJEÇÃO TURNO" 2 "R9")
    $atingimento = 0

    # ---------- Tabela individual ----------
    $people = @()
    $n = 0
    $blankStreak = 0

    for ($r = 14; $r -le 100; $r++) {
        $nome = Get-Text $dash.Cells.Item($r,3).Value2
        if (-not $nome) {
            $blankStreak++
            if ($blankStreak -ge 5 -and $people.Count -gt 0) { break }
            continue
        }

        $blankStreak = 0
        $n++

        $horas = @()
        for ($c = 6; $c -le 13; $c++) {
            $horas += [int](Get-Number $dash.Cells.Item($r,$c).Value2)
        }

        $atingRaw = $dash.Cells.Item($r,16).Value2
        $atingPessoa = $null
        if ($null -ne $atingRaw -and $atingRaw -ne "") {
            $atingPessoa = Get-Number $atingRaw
        }

        $people += [PSCustomObject]@{
            n = $n
            matricula = [int](Get-Number $dash.Cells.Item($r,2).Value2)
            colaborador = $nome
            tipo = Get-Text $dash.Cells.Item($r,4).Value2
            posicao = Get-Text $dash.Cells.Item($r,5).Value2
            horas = $horas
            total = [int](Get-Number $dash.Cells.Item($r,14).Value2)
            media = Get-Number $dash.Cells.Item($r,15).Value2
            ating = $atingPessoa
            status = Get-Text $dash.Cells.Item($r,17).Value2
            justificativa = Get-Text $dash.Cells.Item($r,18).Value2
            resp = Get-Text $dash.Cells.Item($r,19).Value2
        }
    }

    # Produção real = soma do TOTAL de todos os colaboradores listados.
    $producaoReal = [int](($people | Measure-Object -Property total -Sum).Sum)
    if ($projecaoTurno -gt 0) {
        $atingimento = $producaoReal / $projecaoTurno
    } else {
        $atingimento = 0
    }

    # ---------- Projeção ----------
    # Somente estas tipologias podem aparecer no painel.
    $tipologiasValidas = @(
        "SEPARAÇÃO(SECO)",
        "SEPARAÇÃO(DOCE)",
        "SEPARAÇÃO(PACK)",
        "SEPARAÇÃO(QUIMICO)",
        "REMONTE",
        "EMPILHADOR",
        "APOIO",
        "DEVOLUÇÃO"
    )

    $projection = @()

    # Primeiro tenta ler diretamente da tabela oficial da aba T3 (AX:BA).
    for ($r = 2; $r -le 25; $r++) {
        $tipologia = Get-Text $t3.Cells.Item($r,50).Value2   # AX
        if (-not $tipologia) { continue }

        $tipologiaUpper = $tipologia.ToUpperInvariant()
        $match = $tipologiasValidas | Where-Object { $_.ToUpperInvariant() -eq $tipologiaUpper } | Select-Object -First 1
        if (-not $match) { continue }

        $mediaRaw = $t3.Cells.Item($r,51).Value2             # AY
        $staffRaw = $t3.Cells.Item($r,52).Value2             # AZ
        $projRaw  = $t3.Cells.Item($r,53).Value2             # BA

        $projection += [PSCustomObject]@{
            tipologia = $match
            media_dia = if ($mediaRaw -eq "-" -or $null -eq $mediaRaw -or $mediaRaw -eq "") { "-" } else { Get-Number $mediaRaw }
            staff = if ($staffRaw -eq "-" -or $null -eq $staffRaw -or $staffRaw -eq "") { "-" } else { [int](Get-Number $staffRaw) }
            projecao = if ($projRaw -eq "-" -or $null -eq $projRaw -or $projRaw -eq "") { "-" } else { [int](Get-Number $projRaw) }
        }
    }

    # Segurança: remove duplicados e mantém a ordem oficial.
    $projectionFinal = @()
    foreach ($tv in $tipologiasValidas) {
        $item = $projection | Where-Object { $_.tipologia -eq $tv } | Select-Object -First 1
        if ($item) { $projectionFinal += $item }
    }
    $projection = $projectionFinal

    # Se a projeção do card não estiver disponível, soma somente linhas numéricas válidas.
    if ($projecaoTurno -le 0) {
        $projecaoTurno = [int](($projection | Where-Object { $_.projecao -is [ValueType] } | Measure-Object -Property projecao -Sum).Sum)
        if ($projecaoTurno -gt 0) {
            $atingimento = $producaoReal / $projecaoTurno
        }
    }

    # ---------- Ausências ----------
    # Lê diretamente a lista oficial "AUSENCIAS" da aba T3 (coluna BD),
    # para bater exatamente com o total exibido no painel.
    $absences = @()

    for ($r = 2; $r -le 30; $r++) {
        $nomeAusencia = Get-Text $t3.Cells.Item($r,56).Value2   # BD
        if ($nomeAusencia) {
            $absences += $nomeAusencia
        }
    }

    # ---------- Motivos ----------
    $reasonNames = @(
        "FALTA DE ETIQUETA",
        "OSCILAÇÃO COLETOR",
        "BAIXA SISTÊMICA",
        "PAUSA / CEIA",
        "DIVERGÊNCIA ESTOQUE",
        "RETRABALHO",
        "BAIXA PERFORMANCE"
    )

    $reasons = @()
    foreach ($reason in $reasonNames) {
        $count = @($people | Where-Object { $_.justificativa -eq $reason }).Count
        $reasons += [PSCustomObject]@{
            motivo = $reason
            quantidade = $count
        }
    }

    $justPendentes = @(
        $people | Where-Object {
            ($_.status -eq "CRÍTICO" -or $_.status -eq "ATENÇÃO") -and
            [string]::IsNullOrWhiteSpace($_.justificativa)
        }
    ).Count

    $payload = [PSCustomObject]@{
        atualizadoEm = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        origem = $wb.Name
        meta = [PSCustomObject]@{
            data = $dataTexto
            turno = $turno
            setor = $setor
            modalidade = $(if ($modalidade) { $modalidade } else { "TODOS" })
            colaborador = $(if ($colaboradorFiltro) { $colaboradorFiltro } else { "TODOS" })
            staff_planejado = $staffPlanejado
            presentes = $presentes
            ausentes = $ausentes
            presenca = $presenca
            producao_real = $producaoReal
            projecao_turno = $projecaoTurno
            atingimento = $atingimento
            just_pendentes = $justPendentes
        }
        people = $people
        projection = $projection
        absences = $absences
        reasons = $reasons
    }

    $json = $payload | ConvertTo-Json -Depth 10 -Compress
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($jsonPath, $json, $utf8NoBom)

    Write-Host ""
    Write-Host "OK - JSON gerado usando SOMENTE a planilha aberta:"
    Write-Host $wb.Name
    Write-Host ""
    Write-Host "Colaboradores: $($people.Count)"
    Write-Host "Pendencias: $justPendentes"
    Write-Host "Producao: $producaoReal"
    Write-Host "Projecao: $projecaoTurno"
}
finally {
    # Não fecha Excel e não fecha workbook.
    if ($wb) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {} }
    if ($excel) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
