$ErrorActionPreference = "Stop"
$root = "C:\Users\shrey\OneDrive\Documents\New project"
$outDir = Join-Path $root "report\dfd_images"
$pptPath = Join-Path $outDir "smart_hospital_dfd_level_1_set.pptx"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 1280
$pres.PageSetup.SlideHeight = 720

function Add-TextBox {
    param($slide, $text, $left, $top, $width, $height, $size, $bold = $false, $align = 1)
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
    param($slide, $text, $left, $top, $width, $height)
    $shape = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
    $shape.Fill.ForeColor.RGB = 16777215
    $shape.Line.ForeColor.RGB = 0
    $shape.Line.Weight = 2
    $range = $shape.TextFrame.TextRange
    $range.Text = $text
    $range.Font.Name = 'Times New Roman'
    $range.Font.Size = 17
    $range.Font.Bold = -1
    $range.Font.Color.RGB = 0
    $range.ParagraphFormat.Alignment = 2
    return $shape
}

function Add-Process {
    param($slide, $number, $text, $left, $top, $width = 150, $height = 85)
    $shape = $slide.Shapes.AddShape(9, $left, $top, $width, $height)
    $shape.Fill.ForeColor.RGB = 16777215
    $shape.Line.ForeColor.RGB = 0
    $shape.Line.Weight = 2
    $slide.Shapes.AddLine($left + 12, $top + 30, $left + $width - 12, $top + 30).Line.ForeColor.RGB = 0
    $range = $shape.TextFrame.TextRange
    $range.Text = "$number`r$text"
    $range.Font.Name = 'Times New Roman'
    $range.Font.Size = 16
    $range.Font.Bold = -1
    $range.Font.Color.RGB = 0
    $range.ParagraphFormat.Alignment = 2
    return $shape
}

function Add-DataStore {
    param($slide, $text, $left, $top, $width = 150)
    $slide.Shapes.AddLine($left, $top, $left + $width, $top).Line.ForeColor.RGB = 0
    $slide.Shapes.AddLine($left, $top + 22, $left + $width, $top + 22).Line.ForeColor.RGB = 0
    Add-TextBox -slide $slide -text $text -left $left -top ($top - 2) -width $width -height 26 -size 15 -align 2 | Out-Null
}

function Add-Arrow {
    param($slide, $x1, $y1, $x2, $y2, $label, $labelLeft, $labelTop, $labelWidth = 170)
    $line = $slide.Shapes.AddLine($x1, $y1, $x2, $y2)
    $line.Line.ForeColor.RGB = 0
    $line.Line.Weight = 2
    $line.Line.EndArrowheadStyle = 3
    Add-TextBox -slide $slide -text $label -left $labelLeft -top $labelTop -width $labelWidth -height 34 -size 12 -align 2 | Out-Null
}

function Init-Slide {
    param($slide, $titleTop, $titleMain)
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.ForeColor.RGB = 16777215
    Add-TextBox -slide $slide -text $titleTop -left 50 -top 25 -width 430 -height 28 -size 18 -bold $true -align 1 | Out-Null
    Add-TextBox -slide $slide -text $titleMain -left 250 -top 85 -width 780 -height 40 -size 24 -bold $true -align 2 | Out-Null
    $rule = $slide.Shapes.AddLine(40, 145, 1240, 145)
    $rule.Line.ForeColor.RGB = 0
    $rule.Line.Weight = 2
}

# Patient slide
$slide = $pres.Slides.Add(1, 12)
Init-Slide -slide $slide -titleTop 'LEVEL 1 DFD (PATIENT)' -titleMain 'LEVEL 1 DFD: PATIENT'
Add-Box -slide $slide -text 'PATIENT' -left 60 -top 330 -width 180 -height 82 | Out-Null
Add-Process -slide $slide -number '1.0' -text 'Register /`rLogin' -left 450 -top 180 | Out-Null
Add-Process -slide $slide -number '1.1' -text 'Book`rAppointment' -left 450 -top 305 | Out-Null
Add-Process -slide $slide -number '1.2' -text 'Upload`rReport' -left 450 -top 430 | Out-Null
Add-Process -slide $slide -number '1.3' -text 'Cart and`rOrder' -left 450 -top 555 | Out-Null
Add-DataStore -slide $slide -text 'LOGIN DATA' -left 920 -top 215
Add-DataStore -slide $slide -text 'APPOINTMENT DATA' -left 900 -top 340
Add-DataStore -slide $slide -text 'REPORT DATA' -left 925 -top 465
Add-DataStore -slide $slide -text 'ORDER DATA' -left 930 -top 590
Add-Arrow -slide $slide -x1 240 -y1 355 -x2 450 -y2 220 -label 'Manage Register / Request to Login' -labelLeft 255 -labelTop 200
Add-Arrow -slide $slide -x1 450 -y1 250 -x2 240 -y2 382 -label 'Login Result' -labelLeft 290 -labelTop 285 -labelWidth 120
Add-Arrow -slide $slide -x1 600 -y1 220 -x2 920 -y2 220 -label 'Request' -labelLeft 700 -labelTop 192 -labelWidth 90
Add-Arrow -slide $slide -x1 920 -y1 245 -x2 600 -y2 245 -label 'Show Requests' -labelLeft 700 -labelTop 248 -labelWidth 110
Add-Arrow -slide $slide -x1 240 -y1 360 -x2 450 -y2 347 -label 'Request to Book Appointment' -labelLeft 260 -labelTop 335 -labelWidth 170
Add-Arrow -slide $slide -x1 450 -y1 385 -x2 240 -y2 392 -label 'Display Appointment Status' -labelLeft 255 -labelTop 395 -labelWidth 180
Add-Arrow -slide $slide -x1 600 -y1 347 -x2 900 -y2 347 -label 'Request' -labelLeft 700 -labelTop 320 -labelWidth 90
Add-Arrow -slide $slide -x1 900 -y1 372 -x2 600 -y2 372 -label 'Show Status' -labelLeft 700 -labelTop 375 -labelWidth 100
Add-Arrow -slide $slide -x1 240 -y1 375 -x2 450 -y2 472 -label 'Request to Upload Report' -labelLeft 255 -labelTop 448 -labelWidth 170
Add-Arrow -slide $slide -x1 450 -y1 510 -x2 240 -y2 405 -label 'Confirm Upload' -labelLeft 275 -labelTop 505 -labelWidth 120
Add-Arrow -slide $slide -x1 600 -y1 472 -x2 925 -y2 472 -label 'Report Detail' -labelLeft 705 -labelTop 444 -labelWidth 110
Add-Arrow -slide $slide -x1 240 -y1 388 -x2 450 -y2 597 -label 'Request to View / Place Order' -labelLeft 250 -labelTop 585 -labelWidth 180
Add-Arrow -slide $slide -x1 450 -y1 635 -x2 240 -y2 418 -label 'Display Order Status' -labelLeft 255 -labelTop 640 -labelWidth 150
Add-Arrow -slide $slide -x1 600 -y1 597 -x2 930 -y2 597 -label 'Order Details' -labelLeft 710 -labelTop 570 -labelWidth 110
Add-Arrow -slide $slide -x1 930 -y1 622 -x2 600 -y2 622 -label 'Show Order' -labelLeft 715 -labelTop 625 -labelWidth 100

# Doctor slide
$slide = $pres.Slides.Add(2, 12)
Init-Slide -slide $slide -titleTop 'LEVEL 1 DFD (DOCTOR)' -titleMain 'LEVEL 1 DFD: DOCTOR'
Add-Box -slide $slide -text 'DOCTOR' -left 60 -top 330 -width 180 -height 82 | Out-Null
Add-Process -slide $slide -number '2.0' -text 'Doctor`rLogin' -left 450 -top 190 | Out-Null
Add-Process -slide $slide -number '2.1' -text 'View`rAppointments' -left 450 -top 330 | Out-Null
Add-Process -slide $slide -number '2.2' -text 'Update`rPrescription' -left 450 -top 500 | Out-Null
Add-DataStore -slide $slide -text 'DOCTOR LOGIN DATA' -left 885 -top 225
Add-DataStore -slide $slide -text 'APPOINTMENT LIST' -left 915 -top 365
Add-DataStore -slide $slide -text 'PRESCRIPTION DATA' -left 900 -top 535
Add-Arrow -slide $slide -x1 240 -y1 355 -x2 450 -y2 232 -label 'Request to Login' -labelLeft 270 -labelTop 220 -labelWidth 135
Add-Arrow -slide $slide -x1 450 -y1 262 -x2 240 -y2 382 -label 'Login Result' -labelLeft 295 -labelTop 292 -labelWidth 120
Add-Arrow -slide $slide -x1 600 -y1 232 -x2 885 -y2 232 -label 'Request' -labelLeft 700 -labelTop 205 -labelWidth 90
Add-Arrow -slide $slide -x1 885 -y1 257 -x2 600 -y2 257 -label 'Show Requests' -labelLeft 700 -labelTop 260 -labelWidth 110
Add-Arrow -slide $slide -x1 240 -y1 370 -x2 450 -y2 372 -label 'Request to View Appointments' -labelLeft 260 -labelTop 342 -labelWidth 180
Add-Arrow -slide $slide -x1 450 -y1 410 -x2 240 -y2 396 -label 'Display Appointment List' -labelLeft 260 -labelTop 414 -labelWidth 170
Add-Arrow -slide $slide -x1 600 -y1 372 -x2 915 -y2 372 -label 'Appointment Details' -labelLeft 700 -labelTop 344 -labelWidth 130
Add-Arrow -slide $slide -x1 915 -y1 397 -x2 600 -y2 397 -label 'Show Details' -labelLeft 710 -labelTop 400 -labelWidth 100
Add-Arrow -slide $slide -x1 240 -y1 390 -x2 450 -y2 542 -label 'Request to Add Prescription' -labelLeft 255 -labelTop 530 -labelWidth 185
Add-Arrow -slide $slide -x1 450 -y1 580 -x2 240 -y2 412 -label 'Confirm Update' -labelLeft 280 -labelTop 585 -labelWidth 120
Add-Arrow -slide $slide -x1 600 -y1 542 -x2 900 -y2 542 -label 'Prescription Detail' -labelLeft 700 -labelTop 515 -labelWidth 125
Add-Arrow -slide $slide -x1 900 -y1 567 -x2 600 -y2 567 -label 'Show Status' -labelLeft 710 -labelTop 570 -labelWidth 100

# Admin slide
$slide = $pres.Slides.Add(3, 12)
Init-Slide -slide $slide -titleTop 'LEVEL 1 DFD (ADMIN)' -titleMain 'LEVEL 1 DFD: ADMIN'
Add-Box -slide $slide -text 'ADMIN' -left 60 -top 330 -width 180 -height 82 | Out-Null
Add-Process -slide $slide -number '3.0' -text 'Admin`rLogin' -left 450 -top 155 | Out-Null
Add-Process -slide $slide -number '3.1' -text 'Manage`rPatients' -left 450 -top 275 | Out-Null
Add-Process -slide $slide -number '3.2' -text 'Manage`rAppointments' -left 450 -top 395 | Out-Null
Add-Process -slide $slide -number '3.3' -text 'Manage`rOrders' -left 450 -top 515 | Out-Null
Add-DataStore -slide $slide -text 'ADMIN LOGIN DATA' -left 900 -top 190
Add-DataStore -slide $slide -text 'PATIENT DATA' -left 935 -top 310
Add-DataStore -slide $slide -text 'APPOINTMENT DATA' -left 900 -top 430
Add-DataStore -slide $slide -text 'ORDER DATA' -left 930 -top 550
Add-Arrow -slide $slide -x1 240 -y1 350 -x2 450 -y2 197 -label 'Manage Register / Request to Login' -labelLeft 255 -labelTop 180 -labelWidth 180
Add-Arrow -slide $slide -x1 450 -y1 227 -x2 240 -y2 378 -label 'Login Result' -labelLeft 300 -labelTop 255 -labelWidth 110
Add-Arrow -slide $slide -x1 600 -y1 197 -x2 900 -y2 197 -label 'Request' -labelLeft 700 -labelTop 170 -labelWidth 90
Add-Arrow -slide $slide -x1 900 -y1 222 -x2 600 -y2 222 -label 'Show Requests' -labelLeft 700 -labelTop 225 -labelWidth 110
Add-Arrow -slide $slide -x1 240 -y1 365 -x2 450 -y2 317 -label 'Request to Manage Patients' -labelLeft 255 -labelTop 300 -labelWidth 180
Add-Arrow -slide $slide -x1 450 -y1 355 -x2 240 -y2 392 -label 'Display Patient Records' -labelLeft 260 -labelTop 358 -labelWidth 165
Add-Arrow -slide $slide -x1 600 -y1 317 -x2 935 -y2 317 -label 'Patient Details' -labelLeft 710 -labelTop 290 -labelWidth 110
Add-Arrow -slide $slide -x1 935 -y1 342 -x2 600 -y2 342 -label 'Show Data' -labelLeft 720 -labelTop 345 -labelWidth 100
Add-Arrow -slide $slide -x1 240 -y1 382 -x2 450 -y2 437 -label 'Request to Manage Appointments' -labelLeft 250 -labelTop 422 -labelWidth 190
Add-Arrow -slide $slide -x1 450 -y1 475 -x2 240 -y2 405 -label 'Display Appointment Status' -labelLeft 250 -labelTop 478 -labelWidth 180
Add-Arrow -slide $slide -x1 600 -y1 437 -x2 900 -y2 437 -label 'Appointment Details' -labelLeft 700 -labelTop 410 -labelWidth 130
Add-Arrow -slide $slide -x1 900 -y1 462 -x2 600 -y2 462 -label 'Show Status' -labelLeft 710 -labelTop 465 -labelWidth 100
Add-Arrow -slide $slide -x1 240 -y1 398 -x2 450 -y2 557 -label 'Request to Manage Orders' -labelLeft 255 -labelTop 545 -labelWidth 175
Add-Arrow -slide $slide -x1 450 -y1 595 -x2 240 -y2 418 -label 'Display Order Status' -labelLeft 270 -labelTop 598 -labelWidth 150
Add-Arrow -slide $slide -x1 600 -y1 557 -x2 930 -y2 557 -label 'Order Details' -labelLeft 710 -labelTop 530 -labelWidth 110
Add-Arrow -slide $slide -x1 930 -y1 582 -x2 600 -y2 582 -label 'Show Orders' -labelLeft 715 -labelTop 585 -labelWidth 100

$pres.SaveAs($pptPath)
$pres.Slides.Item(1).Export((Join-Path $outDir 'smart_hospital_dfd_level_1_patient.jpg'), 'JPG', 1600, 900)
$pres.Slides.Item(2).Export((Join-Path $outDir 'smart_hospital_dfd_level_1_doctor.jpg'), 'JPG', 1600, 900)
$pres.Slides.Item(3).Export((Join-Path $outDir 'smart_hospital_dfd_level_1_admin.jpg'), 'JPG', 1600, 900)
$pres.Close()
$ppt.Quit()
Write-Output 'Saved Level 1 DFD images.'
