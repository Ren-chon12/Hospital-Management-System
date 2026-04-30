$ErrorActionPreference = "Stop"
$root = "C:\Users\shrey\OneDrive\Documents\New project"
$outDir = Join-Path $root "report\dfd_images"
$pptPath = Join-Path $outDir "smart_hospital_use_case_diagram.pptx"
$jpgPath = Join-Path $outDir "smart_hospital_use_case_diagram.jpg"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 1280
$pres.PageSetup.SlideHeight = 1500
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

function Add-Bullets {
    param($items, $left, $top, $width, $height)
    $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
    $range = $shape.TextFrame.TextRange
    $range.Text = ($items -join "`r")
    $range.Font.Name = 'Times New Roman'
    $range.Font.Size = 17
    $range.Font.Color.RGB = 0
    $range.ParagraphFormat.Bullet.Visible = -1
    $range.ParagraphFormat.Bullet.Character = 8226
    $range.ParagraphFormat.SpaceAfter = 6
    return $shape
}

function Add-Actor {
    param($label, $left, $top)
    $head = $slide.Shapes.AddShape(9, $left + 18, $top, 22, 22)
    $head.Fill.Visible = 0
    $head.Line.ForeColor.RGB = 0
    foreach ($l in @(
        $slide.Shapes.AddLine($left + 29, $top + 22, $left + 29, $top + 62),
        $slide.Shapes.AddLine($left + 12, $top + 36, $left + 46, $top + 36),
        $slide.Shapes.AddLine($left + 29, $top + 62, $left + 12, $top + 88),
        $slide.Shapes.AddLine($left + 29, $top + 62, $left + 46, $top + 88)
    )) { $l.Line.ForeColor.RGB = 0; $l.Line.Weight = 2 }
    Add-TextBox -text $label -left ($left - 25) -top ($top + 95) -width 120 -height 40 -size 16 -align 2 | Out-Null
    return @{ X = ($left + 29); Y = ($top + 40) }
}

function Add-Oval {
    param($text, $left, $top, $width = 150, $height = 38)
    $shape = $slide.Shapes.AddShape(9, $left, $top, $width, $height)
    $shape.Fill.ForeColor.RGB = 16777215
    $shape.Line.ForeColor.RGB = 0
    $shape.Line.Weight = 1.5
    $range = $shape.TextFrame.TextRange
    $range.Text = $text
    $range.Font.Name = 'Times New Roman'
    $range.Font.Size = 14
    $range.Font.Color.RGB = 0
    $range.ParagraphFormat.Alignment = 2
    return $shape
}

function Add-Line {
    param($x1, $y1, $x2, $y2, $dashed = $false)
    $line = $slide.Shapes.AddLine($x1, $y1, $x2, $y2)
    $line.Line.ForeColor.RGB = 0
    $line.Line.Weight = 1.5
    if ($dashed) { $line.Line.DashStyle = 4 }
    return $line
}

Add-TextBox -text 'Actors:' -left 90 -top 35 -width 180 -height 26 -size 20 -bold $true -align 1 | Out-Null
Add-Bullets -items @('Patient', 'Doctor', 'Admin') -left 105 -top 80 -width 250 -height 130 | Out-Null
Add-TextBox -text 'Use Cases:' -left 90 -top 235 -width 180 -height 26 -size 20 -bold $true -align 1 | Out-Null
Add-Bullets -items @('Register / Login', 'Verify OTP', 'Manage Patient Records', 'Manage Appointments', 'Upload Reports', 'Cart and Place Order', 'View Notifications', 'Real-Time Chat', 'View Dashboard', 'Logout') -left 105 -top 280 -width 330 -height 280 | Out-Null

$boundary = $slide.Shapes.AddShape(1, 120, 620, 1040, 760)
$boundary.Fill.Visible = 0
$boundary.Line.ForeColor.RGB = 0
$boundary.Line.Weight = 2
Add-TextBox -text 'Smart Hospital Management System' -left 450 -top 1335 -width 360 -height 28 -size 18 -align 2 | Out-Null

$patient = Add-Actor -label 'Patient' -left 125 -top 860
$doctor = Add-Actor -label 'Doctor' -left 1000 -top 1030
$admin = Add-Actor -label 'Admin' -left 1000 -top 780

# patient use cases
$uc1 = Add-Oval -text 'Register' -left 260 -top 690 -width 135
$uc2 = Add-Oval -text 'Login' -left 260 -top 740 -width 135
$uc3 = Add-Oval -text 'Verify OTP' -left 260 -top 790 -width 135
$uc4 = Add-Oval -text 'View Profile' -left 260 -top 840 -width 135
$uc5 = Add-Oval -text 'Book Appointment' -left 260 -top 890 -width 160
$uc6 = Add-Oval -text 'Upload Reports' -left 260 -top 940 -width 160
$uc7 = Add-Oval -text 'Cart and Place Order' -left 260 -top 990 -width 175
$uc8 = Add-Oval -text 'View Notifications' -left 260 -top 1040 -width 175
$uc9 = Add-Oval -text 'Real-Time Chat' -left 260 -top 1090 -width 160
$uc10 = Add-Oval -text 'Logout' -left 260 -top 1140 -width 135

# admin/doctor use cases
$uc11 = Add-Oval -text 'Manage Patient Records' -left 715 -top 770 -width 195
$uc12 = Add-Oval -text 'Manage Appointments' -left 715 -top 830 -width 185
$uc13 = Add-Oval -text 'View Dashboard' -left 715 -top 890 -width 155
$uc14 = Add-Oval -text 'Manage Orders' -left 715 -top 950 -width 160
$uc15 = Add-Oval -text 'View Appointment List' -left 715 -top 1030 -width 185
$uc16 = Add-Oval -text 'Update Prescription' -left 715 -top 1090 -width 175
$uc17 = Add-Oval -text 'Review Notifications' -left 715 -top 1150 -width 185

# connectors patient
foreach ($pair in @(
    @($patient.X, $patient.Y, 260, 709),
    @($patient.X, $patient.Y, 260, 759),
    @($patient.X, $patient.Y, 260, 809),
    @($patient.X, $patient.Y, 260, 859),
    @($patient.X, $patient.Y, 260, 909),
    @($patient.X, $patient.Y, 260, 959),
    @($patient.X, $patient.Y, 260, 1009),
    @($patient.X, $patient.Y, 260, 1059),
    @($patient.X, $patient.Y, 260, 1109),
    @($patient.X, $patient.Y, 260, 1159)
)) { Add-Line -x1 $pair[0] -y1 $pair[1] -x2 $pair[2] -y2 $pair[3] | Out-Null }

# connectors admin
foreach ($pair in @(
    @($admin.X, $admin.Y, 910, 789),
    @($admin.X, $admin.Y, 900, 849),
    @($admin.X, $admin.Y, 870, 909),
    @($admin.X, $admin.Y, 875, 969),
    @($admin.X, $admin.Y, 900, 1169)
)) { Add-Line -x1 $pair[0] -y1 $pair[1] -x2 $pair[2] -y2 $pair[3] | Out-Null }

# connectors doctor
foreach ($pair in @(
    @($doctor.X, $doctor.Y, 900, 1049),
    @($doctor.X, $doctor.Y, 890, 1109),
    @($doctor.X, $doctor.Y, 900, 1169),
    @($doctor.X, $doctor.Y, 900, 849)
)) { Add-Line -x1 $pair[0] -y1 $pair[1] -x2 $pair[2] -y2 $pair[3] | Out-Null }

# include / extend notes
Add-Line -x1 395 -y1 708 -x2 260 -y2 759 -dashed $true | Out-Null
Add-TextBox -text '<<extends>>' -left 395 -top 708 -width 85 -height 18 -size 11 -align 1 | Out-Null
Add-Line -x1 395 -y1 758 -x2 260 -y2 809 -dashed $true | Out-Null
Add-TextBox -text '<<extends>>' -left 395 -top 758 -width 85 -height 18 -size 11 -align 1 | Out-Null
Add-Line -x1 420 -y1 908 -x2 260 -y2 1059 -dashed $true | Out-Null
Add-TextBox -text '<<includes>>' -left 430 -top 935 -width 90 -height 18 -size 11 -align 1 | Out-Null
Add-Line -x1 900 -y1 849 -x2 910 -y2 969 -dashed $true | Out-Null
Add-TextBox -text '<<includes>>' -left 915 -top 900 -width 90 -height 18 -size 11 -align 1 | Out-Null
Add-Line -x1 900 -y1 1049 -x2 900 -y2 1169 -dashed $true | Out-Null
Add-TextBox -text '<<includes>>' -left 905 -top 1095 -width 90 -height 18 -size 11 -align 1 | Out-Null

Add-TextBox -text 'Use Case Diagram for Smart Hospital Management System' -left 350 -top 1400 -width 580 -height 24 -size 16 -align 2 | Out-Null

$pres.SaveAs($pptPath)
$slide.Export($jpgPath, 'JPG', 1350, 1580)
$pres.Close()
$ppt.Quit()
Write-Output "Saved: $jpgPath"
