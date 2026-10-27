$ErrorActionPreference = "Stop"

$repo = "C:\Users\ivferreira\Documents\GitHub\produtividade-t3-seco"
$jsonPath = Join-Path $repo "dados_produtividade_individual.json"

# Localiza a planilha operacional dentro do repositório.
$excelFile = Get-ChildItem -Path $repo -Filter "Produtividade_T3_SECO*.xlsm" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $excelFile) {
    $excelFile = Get-ChildItem -Path $repo -Filter "*.xlsm" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}
if (-not $excelFile) {
    throw "Nenhum arquivo XLSM encontrado em $repo"
}

$excelPath = $excelFile.FullName
$excel = $null
$wb = $null
$openedByScript = $false

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

try {
    # Tenta usar o Excel que já estiver aberto.
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
    }

    foreach ($book in $excel.Workbooks) {
        try {
            if ($book.FullName -eq $excelPath -or $book.Name -like "Produtividade_T3_SECO*.xlsm") {
                $wb = $book
                break
            }
        } catch {}
    }

    if (-not $wb) {
        $wb = $excel.Workbooks.Open($excelPath, 0, $true)
        $openedByScript = $true
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
    $producaoReal = Get-Number (Get-ValueBelowLabel $dash "PRODUÇÃO REAL" 2 "M9")
    $projecaoTurno = Get-Number (Get-ValueBelowLabel $dash "PROJEÇÃO TURNO" 2 "R9")
    $atingimento = Get-Number (Get-ValueBelowLabel $dash "ATINGIMENTO" 2 "V9")

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

    # ---------- Projeção por tipologia ----------
    $projection = @()
    $projBlank = 0
    for ($r = 14; $r -le 30; $r++) {
        $tipologia = Get-Text $dash.Cells.Item($r,20).Value2
        if (-not $tipologia) {
            $projBlank++
            if ($projBlank -ge 3 -and $projection.Count -gt 0) { break }
            continue
        }
        $projBlank = 0

        $mediaRaw = $dash.Cells.Item($r,21).Value2
        $staffRaw = $dash.Cells.Item($r,22).Value2
        $projRaw = $dash.Cells.Item($r,23).Value2

        $media = if ($null -eq $mediaRaw -or $mediaRaw -eq "" -or $mediaRaw -eq "-") { "-" } else { Get-Number $mediaRaw }
        $staff = if ($null -eq $staffRaw -or $staffRaw -eq "" -or $staffRaw -eq "-") { "-" } else { [int](Get-Number $staffRaw) }
        $proj = if ($null -eq $projRaw -or $projRaw -eq "" -or $projRaw -eq "-") { "-" } else { [int](Get-Number $projRaw) }

        $projection += ,@($tipologia,$media,$staff,$proj)
    }

    # ---------- Ausências: apenas FALTA na aba T3 ----------
    $absences = @()
    for ($r = 2; $r -le 150; $r++) {
        $nome = Get-Text $t3.Cells.Item($r,1).Value2
        $pos = (Get-Text $t3.Cells.Item($r,3).Value2).ToUpperInvariant()
        if ($nome -and $nome -ne "OPEN POSITION" -and $pos -eq "FALTA") {
            $absences += $nome
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
        $reasons += ,@($reason,$count)
    }

    $justPendentes = @(
        $people | Where-Object {
            ($_.status -eq "CRÍTICO" -or $_.status -eq "ATENÇÃO") -and
            [string]::IsNullOrWhiteSpace($_.justificativa)
        }
    ).Count

    $payload = [PSCustomObject]@{
        atualizadoEm = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
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

    Write-Host "JSON atualizado: $jsonPath"
    Write-Host "Colaboradores: $($people.Count) | Pendencias: $justPendentes | Producao: $producaoReal | Projecao: $projecaoTurno"
}
finally {
    if ($openedByScript -and $wb) {
        try { $wb.Close($false) } catch {}
    }
    if ($openedByScript -and $excel) {
        try { $excel.Quit() } catch {}
    }
    if ($wb) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {} }
    if ($excel) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
