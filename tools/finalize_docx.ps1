# Доводка DOCX после MarkToGost средствами Microsoft Word:
#  - таблица, перенесённая на следующую страницу, делится по фактической границе страницы,
#    перед продолжением ставится подпись «Продолжение таблицы N» и повторяется строка заголовка;
#  - обновляется оглавление;
#  - при указании -Pdf сохраняется PDF.
# Использование: powershell -File tools\finalize_docx.ps1 output\Файл.docx [-Pdf output\Файл.pdf]
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$Pdf
)

# Вывод в UTF-8: его читает MarkToGost (флаги --finalize и --pdf)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$wdParagraph = 4
$wdActiveEndPageNumber = 3

$Docx = (Resolve-Path $Docx).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($Docx)
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
    $doc.Repaginate()

    $splits = 0
    $i = 1
    while ($i -le $doc.Tables.Count) {
        $table = $doc.Tables.Item($i)
        $caption = $table.Range.Previous($wdParagraph, 1)
        $text = if ($caption) { $caption.Text } else { "" }
        if ($text -notmatch '^(Таблица|Продолжение таблицы) (\d+)') { $i++; continue }
        $number = $Matches[2]

        $startPage = $table.Rows.Item(1).Range.Information($wdActiveEndPageNumber)
        $splitRow = 0
        for ($r = 2; $r -le $table.Rows.Count; $r++) {
            if ($table.Rows.Item($r).Range.Information($wdActiveEndPageNumber) -ne $startPage) { $splitRow = $r; break }
        }
        # Перенос после первой строки данных или раньше не делим: Word сам перенесёт таблицу целиком
        if ($splitRow -le 2) { $i++; continue }

        $rest = $table.Split($splitRow)

        # Между частями Word вставляет пустой абзац — он становится подписью продолжения
        $para = $doc.Range($table.Range.End, $rest.Range.Start)
        $para.InsertBefore("Продолжение таблицы $number")
        $para.ParagraphFormat.Alignment = $caption.ParagraphFormat.Alignment
        $para.ParagraphFormat.FirstLineIndent = 0
        $para.ParagraphFormat.LeftIndent = 0
        $para.ParagraphFormat.SpaceBefore = 0
        $para.ParagraphFormat.SpaceAfter = 0
        $para.ParagraphFormat.PageBreakBefore = $true
        $para.ParagraphFormat.KeepWithNext = $true
        $para.Font.Name = $caption.Font.Name
        $para.Font.Size = $caption.Font.Size
        $para.Font.Bold = $false
        $para.Font.Italic = $false

        # Строка заголовка в начале продолжения: копия ячеек заголовка с форматированием
        $rest.Rows.Add($rest.Rows.Item(1)) | Out-Null
        for ($c = 1; $c -le $table.Rows.Item(1).Cells.Count; $c++) {
            $src = $table.Cell(1, $c).Range
            $src.MoveEnd(1, -1) | Out-Null
            $dst = $rest.Cell(1, $c).Range
            $dst.MoveEnd(1, -1) | Out-Null
            $dst.FormattedText = $src.FormattedText
            $rest.Cell(1, $c).Range.ParagraphFormat.Alignment = $table.Cell(1, $c).Range.ParagraphFormat.Alignment
        }
        $rest.Rows.Item(1).HeadingFormat = $true

        $splits++
        $i++
        $doc.Repaginate()
    }

    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
    $doc.Save()
    if ($Pdf) {
        $pdfPath = [System.IO.Path]::GetFullPath($Pdf)
        $doc.ExportAsFixedFormat($pdfPath, 17)
    }
    "Разделено таблиц: $splits; страниц: $($doc.ComputeStatistics(2))"
    $doc.Close()
}
finally {
    $word.Quit()
}
