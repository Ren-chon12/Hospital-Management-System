$ErrorActionPreference = "Stop"
$root = "C:\Users\shrey\OneDrive\Documents\New project"
$docx = Join-Path $root "scripts\word_test.docx"
$pdf = Join-Path $root "scripts\word_test.pdf"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()
$sel = $word.Selection
$sel.Font.Name = "Times New Roman"
$sel.Font.Size = 16
$sel.Font.Bold = 1
$sel.TypeText("Test Report Title")
$sel.TypeParagraph()
$sel.TypeParagraph()
$sel.Font.Size = 12
$sel.Font.Bold = 0
$sel.ParagraphFormat.Alignment = 3
$sel.ParagraphFormat.LineSpacingRule = 1
$sel.TypeText("This is a short paragraph to verify Word COM export is working correctly.")
$doc.SaveAs([ref]$docx, [ref]16)
$doc.ExportAsFixedFormat($pdf, 17)
$doc.Close()
$word.Quit()
Write-Output "created" 
