$ErrorActionPreference = "Stop"
$root = "C:\Users\shrey\OneDrive\Documents\New project"
$outDir = Join-Path $root "report\dfd_images"
$pptPath = Join-Path $outDir "smart_hospital_schema_design.pptx"
$jpgPath = Join-Path $outDir "smart_hospital_schema_design.jpg"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 1400
$pres.PageSetup.SlideHeight = 1000
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

function Add-Entity {
    param($title, $rows, $left, $top, $width = 260)
    $headerH = 42
    $rowH = 32
    $height = $headerH + ($rows.Count * $rowH)

    $header = $slide.Shapes.AddShape(1, $left, $top, $width, $headerH)
    $header.Fill.ForeColor.RGB = 14474460
    $header.Line.ForeColor.RGB = 8421504
    $header.Line.Weight = 1.5
    $header.TextFrame.TextRange.Text = $title
    $header.TextFrame.TextRange.Font.Name = 'Times New Roman'
    $header.TextFrame.TextRange.Font.Size = 18
    $header.TextFrame.TextRange.Font.Bold = -1
    $header.TextFrame.TextRange.Font.Color.RGB = 0
    $header.TextFrame.TextRange.ParagraphFormat.Alignment = 2

    for ($i = 0; $i -lt $rows.Count; $i++) {
        $y = $top + $headerH + ($i * $rowH)
        $pkBox = $slide.Shapes.AddShape(1, $left, $y, 44, $rowH)
        $pkBox.Fill.ForeColor.RGB = 16050653
        $pkBox.Line.ForeColor.RGB = 8421504
        $pkBox.Line.Weight = 1
        $pkBox.TextFrame.TextRange.Text = $rows[$i][0]
        $pkBox.TextFrame.TextRange.Font.Name = 'Times New Roman'
        $pkBox.TextFrame.TextRange.Font.Size = 13
        $pkBox.TextFrame.TextRange.Font.Bold = -1
        $pkBox.TextFrame.TextRange.Font.Color.RGB = 0
        $pkBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

        $fieldBox = $slide.Shapes.AddShape(1, $left + 44, $y, $width - 44, $rowH)
        $fieldBox.Fill.ForeColor.RGB = 16777215
        $fieldBox.Line.ForeColor.RGB = 8421504
        $fieldBox.Line.Weight = 1
        $fieldBox.TextFrame.TextRange.Text = $rows[$i][1]
        $fieldBox.TextFrame.TextRange.Font.Name = 'Times New Roman'
        $fieldBox.TextFrame.TextRange.Font.Size = 13
        $fieldBox.TextFrame.TextRange.Font.Color.RGB = 0
        $fieldBox.TextFrame.TextRange.ParagraphFormat.Alignment = 1
    }

    return @{
        Left = $left; Top = $top; Width = $width; Height = $height;
        CenterX = $left + ($width / 2); CenterY = $top + ($height / 2)
    }
}

function Add-Connector {
    param($x1, $y1, $x2, $y2, $dashed = $true)
    $line = $slide.Shapes.AddLine($x1, $y1, $x2, $y2)
    $line.Line.ForeColor.RGB = 0
    $line.Line.Weight = 1.5
    if ($dashed) { $line.Line.DashStyle = 4 }
    $line.Line.BeginArrowheadStyle = 3
    $line.Line.EndArrowheadStyle = 3
    return $line
}

Add-TextBox -text 'Schema Diagram - Smart Hospital Management System' -left 320 -top 25 -width 760 -height 34 -size 24 -bold $false -align 2 | Out-Null

$users = Add-Entity -title 'users' -left 60 -top 110 -rows @(
    @('PK', 'user_id'),
    @('', 'name    VARCHAR(255)'),
    @('', 'email   VARCHAR(255)'),
    @('', 'role    VARCHAR(50)'),
    @('', 'phone   VARCHAR(20)')
)

$patients = Add-Entity -title 'patients' -left 570 -top 110 -rows @(
    @('PK', 'patient_id'),
    @('FK', 'user_id    OBJECTID'),
    @('', 'age        INT'),
    @('', 'blood_group VARCHAR(10)'),
    @('', 'disease    VARCHAR(255)')
)

$appointments = Add-Entity -title 'appointments' -left 1080 -top 110 -rows @(
    @('PK', 'appointment_id'),
    @('FK', 'patient_id    OBJECTID'),
    @('FK', 'doctor_id     OBJECTID'),
    @('', 'date         DATE'),
    @('', 'status       VARCHAR(50)')
)

$orders = Add-Entity -title 'orders' -left 60 -top 470 -rows @(
    @('PK', 'order_id'),
    @('FK', 'user_id      OBJECTID'),
    @('', 'total_amount DECIMAL'),
    @('', 'payment_method VARCHAR(50)'),
    @('', 'order_status  VARCHAR(50)')
)

$messages = Add-Entity -title 'messages' -left 570 -top 470 -rows @(
    @('PK', 'message_id'),
    @('FK', 'sender_id    OBJECTID'),
    @('FK', 'receiver_id  OBJECTID'),
    @('', 'text        TEXT'),
    @('', 'created_at  DATE')
)

$notifications = Add-Entity -title 'notifications' -left 1080 -top 470 -rows @(
    @('PK', 'notification_id'),
    @('FK', 'user_id         OBJECTID'),
    @('', 'title          VARCHAR(255)'),
    @('', 'message        TEXT'),
    @('', 'read           BOOLEAN')
)

$uploads = Add-Entity -title 'uploads / reports' -left 570 -top 775 -rows @(
    @('PK', 'report_id'),
    @('FK', 'patient_id   OBJECTID'),
    @('', 'file_path    VARCHAR(255)'),
    @('', 'uploaded_at  DATE')
)

Add-Connector -x1 ($users.Left + $users.Width) -y1 ($users.Top + 120) -x2 $patients.Left -y2 ($patients.Top + 120) | Out-Null
Add-Connector -x1 ($patients.Left + $patients.Width) -y1 ($patients.Top + 120) -x2 $appointments.Left -y2 ($appointments.Top + 120) | Out-Null
Add-Connector -x1 ($users.Left + 130) -y1 ($users.Top + $users.Height) -x2 ($orders.Left + 130) -y2 $orders.Top | Out-Null
Add-Connector -x1 ($users.Left + $users.Width) -y1 ($users.Top + 160) -x2 $messages.Left -y2 ($messages.Top + 80) | Out-Null
Add-Connector -x1 ($users.Left + $users.Width) -y1 ($users.Top + 200) -x2 $notifications.Left -y2 ($notifications.Top + 90) | Out-Null
Add-Connector -x1 ($patients.Left + 130) -y1 ($patients.Top + $patients.Height) -x2 ($uploads.Left + 130) -y2 $uploads.Top | Out-Null
Add-Connector -x1 ($appointments.Left + 40) -y1 ($appointments.Top + $appointments.Height) -x2 ($notifications.Left + 40) -y2 $notifications.Top | Out-Null
Add-Connector -x1 ($messages.Left + 130) -y1 ($messages.Top + $messages.Height) -x2 ($uploads.Left + 130) -y2 $uploads.Top | Out-Null

Add-TextBox -text 'Schema Diagram - Smart Hospital Management System' -left 360 -top 955 -width 700 -height 26 -size 20 -align 2 | Out-Null

$pres.SaveAs($pptPath)
$slide.Export($jpgPath, 'JPG', 1600, 1140)
$pres.Close()
$ppt.Quit()
Write-Output "Saved: $jpgPath"
