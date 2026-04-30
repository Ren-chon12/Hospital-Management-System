$ErrorActionPreference = "Stop"
$root = "C:\Users\shrey\OneDrive\Documents\New project"
$outDir = Join-Path $root "report\dfd_images"
$pptPath = Join-Path $outDir "smart_hospital_sequence_appointment_booking.pptx"
$jpgPath = Join-Path $outDir "smart_hospital_sequence_appointment_booking.jpg"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 1280
$pres.PageSetup.SlideHeight = 1400
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

function Add-Line {
    param($x1, $y1, $x2, $y2, $arrow = $true, $dashed = $false)
    $line = $slide.Shapes.AddLine($x1, $y1, $x2, $y2)
    $line.Line.ForeColor.RGB = 0
    $line.Line.Weight = 2
    if ($arrow) { $line.Line.EndArrowheadStyle = 3 }
    if ($dashed) { $line.Line.DashStyle = 4 }
    return $line
}

function Add-Lifeline {
    param($label, $left)
    $box = $slide.Shapes.AddShape(1, $left, 170, 150, 46)
    $box.Fill.ForeColor.RGB = 16777164
    $box.Line.ForeColor.RGB = 0
    $box.Line.Weight = 2
    $range = $box.TextFrame.TextRange
    $range.Text = $label
    $range.Font.Name = 'Times New Roman'
    $range.Font.Size = 18
    $range.Font.Bold = -1
    $range.Font.Color.RGB = 0
    $range.ParagraphFormat.Alignment = 2
    $line = $slide.Shapes.AddLine($left + 75, 216, $left + 75, 1015)
    $line.Line.ForeColor.RGB = 0
    $line.Line.DashStyle = 4
    return @{ Center = ($left + 75); Left = $left }
}

function Add-Activation {
    param($centerX, $top, $height)
    $rect = $slide.Shapes.AddShape(1, $centerX - 8, $top, 16, $height)
    $rect.Fill.ForeColor.RGB = 16777215
    $rect.Line.ForeColor.RGB = 0
    return $rect
}

function Add-Message {
    param($x1, $y, $x2, $label, $labelLeft, $labelTop, $dashed = $false)
    Add-Line -x1 $x1 -y1 $y -x2 $x2 -y2 $y -arrow $true -dashed:$dashed | Out-Null
    Add-TextBox -text $label -left $labelLeft -top $labelTop -width 240 -height 24 -size 13 -align 1 | Out-Null
}

Add-TextBox -text '5B. SEQUENCE DIAGRAM (Patient Books Appointment)' -left 55 -top 35 -width 720 -height 30 -size 20 -bold $true -align 1 | Out-Null
Add-TextBox -text 'Patient -> UI -> Backend API -> Database -> Backend API -> UI -> Patient' -left 55 -top 85 -width 780 -height 24 -size 16 -align 1 | Out-Null

$patient = Add-Lifeline -label 'Patient' -left 80
$ui = Add-Lifeline -label 'UI' -left 320
$api = Add-Lifeline -label 'Backend API' -left 560
$db = Add-Lifeline -label 'Database' -left 860

# Patient icon
$head = $slide.Shapes.AddShape(9, 118, 120, 20, 20)
$head.Fill.Visible = 0
$head.Line.ForeColor.RGB = 0
foreach ($l in @(
    $slide.Shapes.AddLine(128,140,128,176),
    $slide.Shapes.AddLine(112,152,144,152),
    $slide.Shapes.AddLine(128,176,112,198),
    $slide.Shapes.AddLine(128,176,144,198)
)) { $l.Line.ForeColor.RGB = 0 }

# Main success flow
Add-Message -x1 $patient.Center -y 270 -x2 $ui.Center -label '1: fill appointment form(details)' -labelLeft 150 -labelTop 248
Add-Activation -centerX $ui.Center -top 250 -height 120 | Out-Null
Add-Message -x1 $ui.Center -y 340 -x2 $api.Center -label '1.1: POST /appointments' -labelLeft 400 -labelTop 318
Add-Activation -centerX $api.Center -top 320 -height 250 | Out-Null
Add-Message -x1 $api.Center -y 395 -x2 $api.Center -label '1.1.1: validate()' -labelLeft 610 -labelTop 372
Add-TextBox -text 'alt' -left 70 -top 430 -width 40 -height 24 -size 14 -bold $true -align 2 | Out-Null
$alt1 = $slide.Shapes.AddShape(1, 95, 425, 1085, 220)
$alt1.Fill.Visible = 0
$alt1.Line.ForeColor.RGB = 0
$alt1.Line.Weight = 1.5
Add-TextBox -text '[if all inputs are valid]' -left 125 -top 438 -width 220 -height 20 -size 13 -align 1 | Out-Null
Add-Message -x1 $api.Center -y 485 -x2 $db.Center -label '1.2: insert appointment : void' -labelLeft 650 -labelTop 462
Add-Activation -centerX $db.Center -top 465 -height 95 | Out-Null
Add-TextBox -text '<<create>>' -left 760 -top 437 -width 100 -height 20 -size 12 -align 2 | Out-Null
$created = $slide.Shapes.AddShape(1, 930, 500, 165, 44)
$created.Fill.ForeColor.RGB = 16777164
$created.Line.ForeColor.RGB = 0
$created.TextFrame.TextRange.Text = 'newAppointment: Appointment'
$created.TextFrame.TextRange.Font.Name = 'Times New Roman'
$created.TextFrame.TextRange.Font.Size = 14
$created.TextFrame.TextRange.Font.Bold = -1
$created.TextFrame.TextRange.Font.Color.RGB = 0
$created.TextFrame.TextRange.ParagraphFormat.Alignment = 2
Add-Message -x1 $db.Center -y 580 -x2 $api.Center -label 'return appointment id' -labelLeft 690 -labelTop 557 -dashed $true
Add-Message -x1 $api.Center -y 615 -x2 $ui.Center -label 'return success response' -labelLeft 445 -labelTop 592 -dashed $true
Add-Message -x1 $ui.Center -y 645 -x2 $patient.Center -label 'display confirmation message' -labelLeft 150 -labelTop 622 -dashed $true

# invalid block
Add-TextBox -text 'alt' -left 70 -top 705 -width 40 -height 24 -size 14 -bold $true -align 2 | Out-Null
$alt2 = $slide.Shapes.AddShape(1, 95, 700, 760, 125)
$alt2.Fill.Visible = 0
$alt2.Line.ForeColor.RGB = 0
$alt2.Line.Weight = 1.5
Add-TextBox -text '[if any input is invalid]' -left 125 -top 713 -width 210 -height 20 -size 13 -align 1 | Out-Null
Add-Message -x1 $api.Center -y 770 -x2 $ui.Center -label '2: return validation error' -labelLeft 440 -labelTop 747 -dashed $true
Add-Message -x1 $ui.Center -y 800 -x2 $patient.Center -label '2.1: display error message()' -labelLeft 145 -labelTop 777 -dashed $true

# Notes
Add-TextBox -text '1. Patient fills appointment form in the dashboard.' -left 70 -top 1110 -width 1000 -height 22 -size 16 -align 1 | Out-Null
Add-TextBox -text '2. Frontend sends POST request to backend API.' -left 70 -top 1150 -width 1000 -height 22 -size 16 -align 1 | Out-Null
Add-TextBox -text '3. Backend validates input and stores the appointment in the database.' -left 70 -top 1190 -width 1080 -height 22 -size 16 -align 1 | Out-Null
Add-TextBox -text '4. Success response is sent back to the frontend.' -left 70 -top 1230 -width 1000 -height 22 -size 16 -align 1 | Out-Null
Add-TextBox -text '5. Patient sees confirmation; otherwise an error message is shown.' -left 70 -top 1270 -width 1080 -height 22 -size 16 -align 1 | Out-Null

$pres.SaveAs($pptPath)
$slide.Export($jpgPath, 'JPG', 1400, 1530)
$pres.Close()
$ppt.Quit()
Write-Output "Saved: $jpgPath"
