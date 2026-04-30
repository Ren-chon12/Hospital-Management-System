$ErrorActionPreference = "Stop"
$root = "C:\Users\shrey\OneDrive\Documents\New project"
$pptPath = Join-Path $root "scripts\ppt_test.pptx"
$pdfPath = Join-Path $root "scripts\ppt_test.pdf"
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 595.3
$pres.PageSetup.SlideHeight = 841.9
$slide = $pres.Slides.Add(1, 12)
$box = $slide.Shapes.AddTextbox(1, 50, 50, 500, 80)
$box.TextFrame.TextRange.Text = "Test Slide"
$box.TextFrame.TextRange.Font.Name = "Times New Roman"
$box.TextFrame.TextRange.Font.Size = 16
$pres.SaveAs($pptPath)
$pres.SaveAs($pdfPath, 32)
$pres.Close()
$ppt.Quit()
Write-Output "created"
