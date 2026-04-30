$ErrorActionPreference = "Stop"
$root = "C:\Users\shrey\OneDrive\Documents\New project"
$outDir = Join-Path $root "report\dfd_images"
$pptPath = Join-Path $outDir "smart_hospital_dfd_level_0.pptx"
$jpgPath = Join-Path $outDir "smart_hospital_dfd_level_0.jpg"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 1280
$pres.PageSetup.SlideHeight = 720
$slide = $pres.Slides.Add(1, 12)
$slide.FollowMasterBackground = 0
$slide.Background.Fill.ForeColor.RGB = 16777215

function Add-TextBox {
    param($text, $left, $top, $width, $height, $size, $bold = $false, $align = 1)
    $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
    $range = $shape.TextFrame.TextRange
    $range.Text = $text
    $range.Font.Name = 'Times New Roman'
    $range.Font.Size = $size
    $range.Font.Color.RGB = 0
    $range.Font.Bold = $(if ($bold) { -1 } else { 0 })
    $range.ParagraphFormat.Alignment = $align
    return $shape
}

function Add-Box {
    param($text, $left, $top, $width, $height)
    $shape = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
    $shape.Fill.ForeColor.RGB = 16777215
    $shape.Line.ForeColor.RGB = 0
    $shape.Line.Weight = 2
    $range = $shape.TextFrame.TextRange
    $range.Text = $text
    $range.Font.Name = 'Times New Roman'
    $range.Font.Size = 18
    $range.Font.Bold = -1
    $range.Font.Color.RGB = 0
    $range.ParagraphFormat.Alignment = 2
    return $shape
}

function Add-Circle {
    param($text, $left, $top, $width, $height)
    $shape = $slide.Shapes.AddShape(9, $left, $top, $width, $height)
    $shape.Fill.ForeColor.RGB = 16777215
    $shape.Line.ForeColor.RGB = 0
    $shape.Line.Weight = 2
    $range = $shape.TextFrame.TextRange
    $range.Text = $text
    $range.Font.Name = 'Times New Roman'
    $range.Font.Size = 18
    $range.Font.Bold = -1
    $range.Font.Color.RGB = 0
    $range.ParagraphFormat.Alignment = 2
    return $shape
}

function Add-Arrow {
    param($x1, $y1, $x2, $y2, $label, $labelLeft, $labelTop, $labelWidth)
    $line = $slide.Shapes.AddLine($x1, $y1, $x2, $y2)
    $line.Line.ForeColor.RGB = 0
    $line.Line.Weight = 2
    $line.Line.EndArrowheadStyle = 3
    Add-TextBox -text $label -left $labelLeft -top $labelTop -width $labelWidth -height 40 -size 12 -align 2 | Out-Null
}

Add-TextBox -text 'LEVEL 0 DFD : SMART HOSPITAL MANAGEMENT SYSTEM' -left 70 -top 40 -width 1140 -height 40 -size 22 -bold $true -align 1 | Out-Null
Add-TextBox -text 'The names of data stores, sources, and destinations are written in capital letters.' -left 160 -top 95 -width 960 -height 26 -size 14 -align 2 | Out-Null
Add-TextBox -text 'Rules for constructing a Data Flow Diagram:' -left 70 -top 145 -width 500 -height 24 -size 18 -bold $true -align 1 | Out-Null
Add-TextBox -text "Arrows should not cross each other.`rSquares, Circles, and Files must bear a name.`rDecomposed data flow squares and circles can have the same names.`rDraw all data flow around the outside of the diagram." -left 130 -top 185 -width 850 -height 120 -size 16 -align 1 | Out-Null

$patient = Add-Box -text 'PATIENT' -left 70 -top 390 -width 170 -height 70
$doctor = Add-Box -text 'DOCTOR' -left 70 -top 565 -width 170 -height 70
$admin = Add-Box -text 'ADMIN' -left 1040 -top 475 -width 170 -height 70
$system = Add-Circle -text "0`rSMART HOSPITAL`rMANAGEMENT SYSTEM" -left 450 -top 430 -width 330 -height 150

Add-Arrow -x1 240 -y1 425 -x2 450 -y2 470 -label "REGISTRATION, LOGIN,`rAPPOINTMENT REQUEST, ORDER DETAILS" -labelLeft 245 -labelTop 360 -labelWidth 210
Add-Arrow -x1 450 -y1 520 -x2 240 -y2 455 -label "APPOINTMENT STATUS,`rORDER CONFIRMATION, NOTIFICATIONS" -labelLeft 245 -labelTop 500 -labelWidth 215

Add-Arrow -x1 240 -y1 600 -x2 450 -y2 545 -label "PRESCRIPTION DETAILS,`rAPPOINTMENT UPDATE, CHAT MESSAGE" -labelLeft 245 -labelTop 610 -labelWidth 220
Add-Arrow -x1 450 -y1 495 -x2 240 -y2 575 -label "PATIENT DETAILS,`rAPPOINTMENT LIST, NOTIFICATIONS" -labelLeft 255 -labelTop 545 -labelWidth 205

Add-Arrow -x1 780 -y1 490 -x2 1040 -y2 500 -label "PATIENT RECORDS, USER DATA,`rAPPOINTMENT DATA, ORDER STATUS" -labelLeft 805 -labelTop 430 -labelWidth 220
Add-Arrow -x1 1040 -y1 530 -x2 780 -y2 535 -label "DASHBOARD DETAILS,`rORDER DETAILS, SYSTEM REPORTS" -labelLeft 805 -labelTop 545 -labelWidth 210

$pres.SaveAs($pptPath)
$slide.Export($jpgPath, 'JPG', 1600, 900)
$pres.Close()
$ppt.Quit()
Write-Output "Saved: $jpgPath"
