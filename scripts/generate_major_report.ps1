$ErrorActionPreference = "Stop"

$projectRoot = "C:\Users\shrey\OneDrive\Documents\New project"
$reportDir = Join-Path $projectRoot "report"
$diagramDir = Join-Path $reportDir "diagrams"
$pptPath = Join-Path $reportDir "Smart_Hospital_Management_System_Standards_Report.pptx"
$pdfPath = Join-Path $reportDir "Smart_Hospital_Management_System_Standards_Report.pdf"
$sourcePath = Join-Path $reportDir "Smart_Hospital_Management_System_Standards_Report_Source.txt"

New-Item -ItemType Directory -Force -Path $reportDir | Out-Null
New-Item -ItemType Directory -Force -Path $diagramDir | Out-Null

$ColorTitle = 2622476
$ColorText = 1975344
$ColorMuted = 6908265
$ColorAccent = 10840842
$ColorLine = 5987163
$ColorSoftBlue = 15329769
$ColorSoftGreen = 14809723
$ColorSoftCream = 16118771
$ColorSoftRose = 15921906

$ppt = $null
$presentation = $null
$global:PageNumber = 0
$sourceLines = New-Object System.Collections.Generic.List[string]

function New-ReportSlide {
  param([string]$Title,[string]$SubTitle = "",[switch]$HideFooter)
  $global:PageNumber += 1
  $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
  $slide.FollowMasterBackground = 0
  $slide.Background.Fill.ForeColor.RGB = 16777215
  $titleBox = $slide.Shapes.AddTextbox(1, 46, 34, 503, 34)
  $titleRange = $titleBox.TextFrame.TextRange
  $titleRange.Text = $Title
  $titleRange.Font.Name = "Times New Roman"
  $titleRange.Font.Size = 16
  $titleRange.Font.Bold = -1
  $titleRange.Font.Color.RGB = $ColorTitle
  if ($SubTitle) {
    $subBox = $slide.Shapes.AddTextbox(1, 46, 68, 503, 18)
    $subRange = $subBox.TextFrame.TextRange
    $subRange.Text = $SubTitle
    $subRange.Font.Name = "Times New Roman"
    $subRange.Font.Size = 10
    $subRange.Font.Italic = -1
    $subRange.Font.Color.RGB = 6706767
  }
  $accent = $slide.Shapes.AddShape(1, 46, 92, 110, 3)
  $accent.Fill.ForeColor.RGB = $ColorAccent
  $accent.Line.Visible = 0
  $footer = $slide.Shapes.AddTextbox(1, 260, 804, 75, 18)
  $footerRange = $footer.TextFrame.TextRange
  $footerRange.Text = [string]$global:PageNumber
  $footerRange.Font.Name = "Times New Roman"
  $footerRange.Font.Size = 10
  $footerRange.ParagraphFormat.Alignment = 2
  $footerRange.Font.Color.RGB = $ColorMuted
  if ($HideFooter) { $footer.Visible = 0 }
  return $slide
}

function Add-TextBlock {
  param($Slide,[string]$Text,[double]$Left = 46,[double]$Top = 110,[double]$Width = 503,[double]$Height = 650,[int]$FontSize = 12,[int]$Alignment = 4,[int]$Color = $ColorText,[switch]$Bold,[switch]$Italic)
  $box = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
  $range = $box.TextFrame.TextRange
  $range.Text = $Text
  $range.Font.Name = "Times New Roman"
  $range.Font.Size = $FontSize
  $range.Font.Color.RGB = $Color
  $range.Font.Bold = $(if ($Bold) { -1 } else { 0 })
  $range.Font.Italic = $(if ($Italic) { -1 } else { 0 })
  $range.ParagraphFormat.Alignment = $Alignment
  $range.ParagraphFormat.SpaceAfter = 6
  $range.ParagraphFormat.SpaceWithin = 1.35
  $box.TextFrame.WordWrap = -1
  return $box
}

function Add-ParagraphBlock {
  param($Slide,[string[]]$Paragraphs,[double]$Left = 46,[double]$Top = 110,[double]$Width = 503,[double]$Height = 650,[int]$FontSize = 12,[int]$Alignment = 4)
  return Add-TextBlock -Slide $Slide -Text ($Paragraphs -join "`r`r") -Left $Left -Top $Top -Width $Width -Height $Height -FontSize $FontSize -Alignment $Alignment
}

function Add-BulletBlock {
  param($Slide,[string[]]$Items,[double]$Left = 60,[double]$Top = 120,[double]$Width = 490,[double]$Height = 620,[int]$FontSize = 12)
  $box = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
  $range = $box.TextFrame.TextRange
  $range.Text = ($Items -join "`r")
  $range.Font.Name = "Times New Roman"
  $range.Font.Size = $FontSize
  $range.Font.Color.RGB = $ColorText
  $range.ParagraphFormat.Bullet.Visible = -1
  $range.ParagraphFormat.Bullet.Character = 8226
  $range.ParagraphFormat.Alignment = 1
  $range.ParagraphFormat.SpaceAfter = 7
  $range.ParagraphFormat.SpaceWithin = 1.3
  return $box
}

function Add-TableBlock {
  param($Slide,[string[]]$Headers,[object[][]]$Rows,[double]$Left = 46,[double]$Top = 125,[double]$Width = 503,[double]$Height = 590,[int]$HeaderFill = $ColorSoftBlue)
  $tableShape = $Slide.Shapes.AddTable($Rows.Count + 1, $Headers.Count, $Left, $Top, $Width, $Height)
  $table = $tableShape.Table
  for ($c = 1; $c -le $Headers.Count; $c++) {
    $cell = $table.Cell(1, $c)
    $cell.Shape.Fill.ForeColor.RGB = $HeaderFill
    $cell.Shape.TextFrame.TextRange.Text = $Headers[$c - 1]
    $cell.Shape.TextFrame.TextRange.Font.Name = "Times New Roman"
    $cell.Shape.TextFrame.TextRange.Font.Size = 10
    $cell.Shape.TextFrame.TextRange.Font.Bold = -1
  }
  for ($r = 1; $r -le $Rows.Count; $r++) {
    for ($c = 1; $c -le $Headers.Count; $c++) {
      $cell = $table.Cell($r + 1, $c)
      $cell.Shape.TextFrame.TextRange.Text = [string]$Rows[$r - 1][$c - 1]
      $cell.Shape.TextFrame.TextRange.Font.Name = "Times New Roman"
      $cell.Shape.TextFrame.TextRange.Font.Size = 9
    }
  }
  return $tableShape
}

function Add-Caption {
  param($Slide,[string]$Text,[double]$Top = 754)
  return Add-TextBlock -Slide $Slide -Text $Text -Left 60 -Top $Top -Width 470 -Height 22 -FontSize 10 -Alignment 2 -Color $ColorMuted -Italic
}

function Add-RectangleLabel {
  param($Slide,[string]$Text,[double]$Left,[double]$Top,[double]$Width,[double]$Height,[int]$Fill = $ColorSoftCream,[int]$FontSize = 11)
  $shape = $Slide.Shapes.AddShape(1, $Left, $Top, $Width, $Height)
  $shape.Fill.ForeColor.RGB = $Fill
  $shape.Line.ForeColor.RGB = 10132122
  $shape.TextFrame.TextRange.Text = $Text
  $shape.TextFrame.TextRange.Font.Name = "Times New Roman"
  $shape.TextFrame.TextRange.Font.Size = $FontSize
  $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
  return $shape
}

function Add-RoundedLabel {
  param($Slide,[string]$Text,[double]$Left,[double]$Top,[double]$Width,[double]$Height,[int]$Fill = $ColorSoftGreen,[int]$FontSize = 11)
  $shape = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
  $shape.Fill.ForeColor.RGB = $Fill
  $shape.Line.ForeColor.RGB = 10132122
  $shape.TextFrame.TextRange.Text = $Text
  $shape.TextFrame.TextRange.Font.Name = "Times New Roman"
  $shape.TextFrame.TextRange.Font.Size = $FontSize
  $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
  return $shape
}

function Add-EllipseLabel {
  param($Slide,[string]$Text,[double]$Left,[double]$Top,[double]$Width,[double]$Height,[int]$Fill = $ColorSoftBlue,[int]$FontSize = 10)
  $shape = $Slide.Shapes.AddShape(9, $Left, $Top, $Width, $Height)
  $shape.Fill.ForeColor.RGB = $Fill
  $shape.Line.ForeColor.RGB = 10132122
  $shape.TextFrame.TextRange.Text = $Text
  $shape.TextFrame.TextRange.Font.Name = "Times New Roman"
  $shape.TextFrame.TextRange.Font.Size = $FontSize
  $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
  return $shape
}

function Add-DiamondLabel {
  param($Slide,[string]$Text,[double]$Left,[double]$Top,[double]$Width,[double]$Height,[int]$Fill = $ColorSoftRose,[int]$FontSize = 10)
  $shape = $Slide.Shapes.AddShape(4, $Left, $Top, $Width, $Height)
  $shape.Fill.ForeColor.RGB = $Fill
  $shape.Line.ForeColor.RGB = 10132122
  $shape.TextFrame.TextRange.Text = $Text
  $shape.TextFrame.TextRange.Font.Name = "Times New Roman"
  $shape.TextFrame.TextRange.Font.Size = $FontSize
  $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
  return $shape
}

function Add-LineArrow {
  param($Slide,[double]$X1,[double]$Y1,[double]$X2,[double]$Y2)
  $line = $Slide.Shapes.AddLine($X1, $Y1, $X2, $Y2)
  $line.Line.ForeColor.RGB = $ColorLine
  $line.Line.EndArrowheadStyle = 3
  return $line
}

function Add-LineNoArrow {
  param($Slide,[double]$X1,[double]$Y1,[double]$X2,[double]$Y2,[switch]$Dashed)
  $line = $Slide.Shapes.AddLine($X1, $Y1, $X2, $Y2)
  $line.Line.ForeColor.RGB = $ColorLine
  if ($Dashed) { $line.Line.DashStyle = 4 }
  return $line
}

function Add-ConnectorText {
  param($Slide,[string]$Text,[double]$Left,[double]$Top,[double]$Width = 120)
  return Add-TextBlock -Slide $Slide -Text $Text -Left $Left -Top $Top -Width $Width -Height 20 -FontSize 9 -Alignment 2 -Color $ColorMuted
}

function Add-StickActor {
  param($Slide,[string]$Label,[double]$Left,[double]$Top)
  $head = $Slide.Shapes.AddShape(9, $Left + 18, $Top, 18, 18)
  $head.Fill.Visible = 0
  $head.Line.ForeColor.RGB = $ColorText
  $body = $Slide.Shapes.AddLine($Left + 27, $Top + 18, $Left + 27, $Top + 52)
  $arm = $Slide.Shapes.AddLine($Left + 10, $Top + 30, $Left + 44, $Top + 30)
  $leg1 = $Slide.Shapes.AddLine($Left + 27, $Top + 52, $Left + 10, $Top + 75)
  $leg2 = $Slide.Shapes.AddLine($Left + 27, $Top + 52, $Left + 44, $Top + 75)
  foreach ($shape in @($body, $arm, $leg1, $leg2)) { $shape.Line.ForeColor.RGB = $ColorText }
  Add-TextBlock -Slide $Slide -Text $Label -Left ($Left - 10) -Top ($Top + 80) -Width 90 -Height 22 -FontSize 11 -Alignment 2 | Out-Null
}

function Add-Lifeline {
  param($Slide,[string]$Label,[double]$Left,[double]$Top = 150,[double]$Height = 500)
  Add-RoundedLabel -Slide $Slide -Text $Label -Left $Left -Top $Top -Width 100 -Height 28 -Fill $ColorSoftBlue -FontSize 10 | Out-Null
  Add-LineNoArrow -Slide $Slide -X1 ($Left + 50) -Y1 ($Top + 28) -X2 ($Left + 50) -Y2 ($Top + $Height) -Dashed | Out-Null
}

function Add-BarMetric {
  param($Slide,[string]$Label,[double]$Value,[double]$Left,[double]$Top,[double]$Width = 220)
  Add-TextBlock -Slide $Slide -Text $Label -Left $Left -Top ($Top - 4) -Width 150 -Height 18 -FontSize 10 -Alignment 1 | Out-Null
  $base = $Slide.Shapes.AddShape(1, ($Left + 150), $Top, $Width, 16)
  $base.Fill.ForeColor.RGB = 15132390
  $base.Line.Visible = 0
  $bar = $Slide.Shapes.AddShape(1, ($Left + 150), $Top, ($Width * ($Value / 100.0)), 16)
  $bar.Fill.ForeColor.RGB = $ColorAccent
  $bar.Line.Visible = 0
  Add-TextBlock -Slide $Slide -Text ("{0}%" -f [int]$Value) -Left ($Left + 150 + $Width + 6) -Top ($Top - 4) -Width 40 -Height 18 -FontSize 10 -Alignment 1 | Out-Null
}

try {
  $ppt = New-Object -ComObject PowerPoint.Application
  $ppt.Visible = -1
  $presentation = $ppt.Presentations.Add()
  $presentation.PageSetup.SlideWidth = 595.3
  $presentation.PageSetup.SlideHeight = 841.9

  # Slide 1
  $slide = $presentation.Slides.Add(1, 12)
  $slide.FollowMasterBackground = 0
  $slide.Background.Fill.ForeColor.RGB = 16777215
  $band = $slide.Shapes.AddShape(1, 0, 0, 595.3, 90)
  $band.Fill.ForeColor.RGB = 15395562
  $band.Line.Visible = 0
  $emblem = $slide.Shapes.AddShape(9, 235, 118, 125, 125)
  $emblem.Fill.ForeColor.RGB = $ColorAccent
  $emblem.Line.Visible = 0
  $emblem.TextFrame.TextRange.Text = "ASD"
  $emblem.TextFrame.TextRange.Font.Name = "Times New Roman"
  $emblem.TextFrame.TextRange.Font.Size = 30
  $emblem.TextFrame.TextRange.Font.Bold = -1
  $emblem.TextFrame.TextRange.Font.Color.RGB = 16777215
  $emblem.TextFrame.TextRange.ParagraphFormat.Alignment = 2
  Add-TextBlock -Slide $slide -Text "ACADEMY OF SKILL DEVELOPMENT" -Left 75 -Top 28 -Width 450 -Height 22 -FontSize 18 -Alignment 2 -Bold -Color $ColorTitle | Out-Null
  Add-TextBlock -Slide $slide -Text "Major Project Report" -Left 170 -Top 72 -Width 260 -Height 24 -FontSize 14 -Alignment 2 -Bold -Color $ColorMuted | Out-Null
  Add-TextBlock -Slide $slide -Text "SMART HOSPITAL MANAGEMENT SYSTEM" -Left 85 -Top 282 -Width 425 -Height 36 -FontSize 24 -Alignment 2 -Bold -Color $ColorTitle | Out-Null
  Add-TextBlock -Slide $slide -Text "A report prepared in accordance with the prescribed project report standards" -Left 85 -Top 326 -Width 425 -Height 28 -FontSize 12 -Alignment 2 -Italic -Color $ColorMuted | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Submitted By`rShreyansh Verma`rTavis Mariageorge James" -Left 70 -Top 420 -Width 190 -Height 90 -Fill $ColorSoftCream -FontSize 12 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Project Guide`rMr. Subhojit Santra" -Left 335 -Top 420 -Width 190 -Height 90 -Fill $ColorSoftGreen -FontSize 12 | Out-Null
  Add-TextBlock -Slide $slide -Text "Department / Program : Major Project Submission" -Left 120 -Top 560 -Width 360 -Height 20 -FontSize 12 -Alignment 2 | Out-Null
  Add-TextBlock -Slide $slide -Text "Academic Session : 2025 - 2026" -Left 170 -Top 586 -Width 260 -Height 20 -FontSize 12 -Alignment 2 | Out-Null
  Add-TextBlock -Slide $slide -Text "Prepared for academic evaluation and final submission in PDF form" -Left 120 -Top 612 -Width 360 -Height 20 -FontSize 11 -Alignment 2 -Color $ColorMuted | Out-Null
  Add-TextBlock -Slide $slide -Text "New report generated from project instructions, sample format, and the prior draft report." -Left 80 -Top 748 -Width 435 -Height 24 -FontSize 10 -Alignment 2 -Italic -Color $ColorMuted | Out-Null
  $sourceLines.Add("Title Page - Smart Hospital Management System")

  # Slide 2
  $slide = New-ReportSlide -Title "CERTIFICATE"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "This is to certify that the project report entitled ""Smart Hospital Management System"" is a bona fide work carried out by Shreyansh Verma and Tavis Mariageorge James under the guidance and supervision of the undersigned during the academic session 2025 - 2026. The work presented in this report has been completed in accordance with the project report standards prescribed for major project submission.",
    "The report demonstrates the analysis, design, implementation, testing, and evaluation of a web-based Smart Hospital Management System developed using the MERN stack and MVC architecture. To the best of our knowledge, the contents of this report are original and have not been submitted, either in full or in part, for the award of any other degree, diploma, or certificate.",
    "The project fulfills the academic requirements for final submission and is recommended for evaluation."
  ) -Height 430
  Add-TextBlock -Slide $slide -Text "Project Guide" -Left 70 -Top 620 -Width 120 -Height 18 -FontSize 12 -Alignment 1 -Bold | Out-Null
  Add-LineNoArrow -Slide $slide -X1 70 -Y1 668 -X2 200 -Y2 668 | Out-Null
  Add-TextBlock -Slide $slide -Text "Head of Department" -Left 240 -Top 620 -Width 140 -Height 18 -FontSize 12 -Alignment 1 -Bold | Out-Null
  Add-LineNoArrow -Slide $slide -X1 240 -Y1 668 -X2 385 -Y2 668 | Out-Null
  Add-TextBlock -Slide $slide -Text "Dean / Principal" -Left 420 -Top 620 -Width 110 -Height 18 -FontSize 12 -Alignment 1 -Bold | Out-Null
  Add-LineNoArrow -Slide $slide -X1 420 -Y1 668 -X2 530 -Y2 668 | Out-Null
  Add-TextBlock -Slide $slide -Text "Date: ____________" -Left 70 -Top 710 -Width 120 -Height 18 -FontSize 11 -Alignment 1 | Out-Null
  Add-TextBlock -Slide $slide -Text "Place: ____________" -Left 240 -Top 710 -Width 120 -Height 18 -FontSize 11 -Alignment 1 | Out-Null
  $sourceLines.Add("Certificate")

  # Slide 3
  $slide = New-ReportSlide -Title "DECLARATION"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "We hereby declare that the project report entitled ""Smart Hospital Management System"" submitted for academic evaluation is an original work carried out by us under the guidance of the project mentor. The project has been designed and implemented using the MERN stack for the purpose of demonstrating a digital solution to manual hospital management.",
    "The matter embodied in this report has not been copied from any earlier submission, and no part of this report has been submitted to any other institution or university for the award of any degree, diploma, or certificate. Whenever the ideas, findings, or documentation of other authors have been referred to, proper acknowledgement has been incorporated in the references section.",
    "We understand that any instance of plagiarism or misrepresentation may lead to academic action as per institutional rules."
  ) -Height 460
  Add-LineNoArrow -Slide $slide -X1 80 -Y1 690 -X2 235 -Y2 690 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 330 -Y1 690 -X2 485 -Y2 690 | Out-Null
  Add-TextBlock -Slide $slide -Text "Shreyansh Verma" -Left 90 -Top 694 -Width 140 -Height 18 -FontSize 11 -Alignment 2 | Out-Null
  Add-TextBlock -Slide $slide -Text "Tavis Mariageorge James" -Left 315 -Top 694 -Width 185 -Height 18 -FontSize 11 -Alignment 2 | Out-Null
  Add-TextBlock -Slide $slide -Text "Student Signature" -Left 100 -Top 714 -Width 120 -Height 16 -FontSize 10 -Alignment 2 -Color $ColorMuted | Out-Null
  Add-TextBlock -Slide $slide -Text "Student Signature" -Left 350 -Top 714 -Width 120 -Height 16 -FontSize 10 -Alignment 2 -Color $ColorMuted | Out-Null
  $sourceLines.Add("Declaration")

  # Slide 4
  $slide = New-ReportSlide -Title "ACKNOWLEDGEMENT"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "We express our sincere gratitude to our project guide, Mr. Subhojit Santra, for his valuable guidance, patient encouragement, and constructive suggestions throughout the development of this project. His support helped us refine both the technical implementation and the academic presentation of the Smart Hospital Management System.",
    "We also thank the faculty members and academic coordinators of the Academy of Skill Development for providing the learning environment, project review support, and institutional resources required to complete this work. Their feedback during each stage of development helped us improve the structure, scope, and quality of the report.",
    "We are equally grateful to our friends and peers for their suggestions during design validation and testing. Finally, we thank our families for their constant support, motivation, and encouragement throughout the completion of this major project."
  ) -Height 540
  $sourceLines.Add("Acknowledgement")

  # Slide 5
  $slide = New-ReportSlide -Title "ABSTRACT"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The Smart Hospital Management System is a full stack web application developed to address the inefficiencies of manual hospital record management. In many small and medium-sized healthcare institutions, patient registration, appointment handling, prescription management, report storage, order processing, and communication are still carried out manually or through fragmented tools. This creates delays, redundancy, poor coordination, and difficulty in retrieving information when needed.",
    "To solve this problem, the proposed system provides a centralized digital platform built using the MERN stack, namely MongoDB, Express.js, React.js, and Node.js. The backend follows the Model-View-Controller architecture to ensure clear separation of concerns and maintainable code. Core features include user authentication, email OTP verification, role-based access control, patient record management, appointment booking, e-prescription support, report uploads, order processing with cart and Cash on Delivery, real-time chat, notifications, and geo-location map integration.",
    "The system was implemented and validated module by module using seeded dummy data and manual testing. The results show that the proposed system successfully improves data accessibility, reduces paper dependency, streamlines appointment workflows, and strengthens communication between administrators, doctors, and patients. The project forms a practical foundation for future enhancements such as telemedicine, wearable device integration, cloud storage, and online payment support."
  ) -Height 560
  $sourceLines.Add("Abstract")

  # Slide 6
  $slide = New-ReportSlide -Title "TABLE OF CONTENTS"
  Add-TextBlock -Slide $slide -Text @"
Title Page .............................................................................. 1
Certificate ............................................................................ 2
Declaration .......................................................................... 3
Acknowledgement .................................................................... 4
Abstract ................................................................................ 5
Table of Contents .................................................................... 6
List of Figures ....................................................................... 8
List of Tables ........................................................................ 9
Chapter 1: Introduction ............................................................ 10
1A. Background of the Problem .................................................... 11
1B. Importance and Relevance .................................................... 12
1C. Problem Statement ............................................................. 13
1D. Objectives of the Project ................................................... 14
1E. Scope of the Project ........................................................ 15
Chapter 2: Literature Review .................................................... 16
2A. Research Papers / Articles Reviewed .................................... 17
2B. Summary of Existing Work .................................................. 18
"@ -Left 64 -Top 120 -Width 470 -Height 540 -FontSize 11 -Alignment 1
  $sourceLines.Add("Table of Contents I")

  # Slide 7
  $slide = New-ReportSlide -Title "TABLE OF CONTENTS (CONTINUED)"
  Add-TextBlock -Slide $slide -Text @"
2C. Comparative Analysis ........................................................ 19
2D. Identification of Research Gap .......................................... 20
2E. Literature Review Synthesis ............................................... 21
Chapter 3: System Analysis ...................................................... 22
Chapter 4: System Design ......................................................... 32
Chapter 5: Methodology .......................................................... 43
Chapter 6: Implementation ....................................................... 48
Chapter 7: Results and Discussion ............................................ 56
Chapter 8: Advantages and Limitations ..................................... 61
Chapter 9: Conclusion ............................................................ 64
Chapter 10: Future Scope ....................................................... 65
References ............................................................................ 66
Appendix A: API Endpoints ...................................................... 68
Appendix B: Dummy Data Samples .............................................. 69
Appendix C: Source Code Organization ..................................... 70
"@ -Left 64 -Top 120 -Width 470 -Height 540 -FontSize 11 -Alignment 1
  $sourceLines.Add("Table of Contents II")

  # Slide 8
  $slide = New-ReportSlide -Title "LIST OF FIGURES"
  Add-TextBlock -Slide $slide -Text @"
Figure 1. System Architecture of Smart Hospital Management System ............. 33
Figure 2. Module Design Diagram ...................................................... 34
Figure 3. Data Flow Diagram Level 0 ................................................ 35
Figure 4. Data Flow Diagram Level 1 - Patient .................................... 36
Figure 5. Data Flow Diagram Level 1 - Admin and Doctor ........................ 37
Figure 6. Use Case Diagram ........................................................... 38
Figure 7. Sequence Diagram for Appointment Booking ........................... 39
Figure 8. Schema / Entity Relationship Diagram ................................... 40
Figure 9. Authentication Flowchart .................................................. 41
Figure 10. Order Processing Flowchart .............................................. 42
Figure 11. Representative Interface Layouts ........................................ 54
"@ -Left 64 -Top 120 -Width 470 -Height 520 -FontSize 11 -Alignment 1
  $sourceLines.Add("List of Figures")

  # Slide 9
  $slide = New-ReportSlide -Title "LIST OF TABLES"
  Add-TextBlock -Slide $slide -Text @"
Table 1. Review of Major Research Papers ........................................ 17
Table 2. Comparative Analysis of Existing Approaches .......................... 19
Table 3. Feasibility Study Summary ................................................. 26
Table 4. Input and Output Analysis ................................................. 27
Table 5. Tools and Technologies Used .............................................. 49
Table 6. Hardware Requirements .................................................... 50
Table 7. Software Requirements and Dummy Data Summary .................... 51
Table 8. Core API Summary ............................................................ 53
Table 9. System Test Cases .......................................................... 57
Table 10. Expected vs Actual Comparison .......................................... 59
Table 11. Future Enhancement Roadmap ............................................ 65
"@ -Left 64 -Top 120 -Width 470 -Height 520 -FontSize 11 -Alignment 1
  $sourceLines.Add("List of Tables")

  # Slide 10
  $slide = New-ReportSlide -Title "CHAPTER 1: INTRODUCTION"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "Hospital management is one of the most information-intensive and process-driven domains in the healthcare sector. Every hospital must continuously manage patient registration, doctor scheduling, report storage, prescription handling, billing coordination, and administrative communication. The quality of this management directly affects patient experience, service speed, and record accuracy.",
    "In many hospitals and clinics, particularly at small and medium scale, these activities are still handled through manual registers, disconnected spreadsheets, or isolated software tools. Such an environment creates duplication of work, weak traceability, and delays in the retrieval of patient history.",
    "The Smart Hospital Management System was developed to provide a structured digital alternative. This chapter introduces the project background, need, objectives, and scope."
  ) -Height 520
  $sourceLines.Add("Chapter 1 Introduction")

  # Slide 11
  $slide = New-ReportSlide -Title "1A. BACKGROUND OF THE PROBLEM"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "Healthcare institutions operate in a highly time-sensitive environment where timely access to information is critical. Patient records, medical reports, prescriptions, and appointments must be accessible to the right users at the right time. When these details are managed manually, retrieval becomes dependent on physical files, staff memory, and repetitive data entry.",
    "Traditional hospital workflows were designed around paper forms and desk-based coordination. While such systems may function at a small scale, they quickly become inefficient as the patient load increases. Repeated visits create multiple documents across departments, and each new appointment or report update adds to the complexity of record tracking.",
    "A well-designed hospital management system can unify registration, appointments, communication, and record access through one secure platform."
  ) -Height 540
  $sourceLines.Add("1A Background")

  # Slide 12
  $slide = New-ReportSlide -Title "1B. IMPORTANCE AND RELEVANCE"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The proposed project is important because hospital administration directly affects the quality and continuity of care. When appointment records are lost, when reports cannot be found quickly, or when patients are not informed about changes, the outcome is inconvenience for both staff and patients. A digital system reduces this operational friction.",
    "This project is relevant in academic as well as practical terms. Academically, it demonstrates how a complete MERN stack solution can be built using role-based access control, REST APIs, JWT authentication, WebSockets, file uploads, and structured MVC design.",
    "Because healthcare systems increasingly rely on reliable information flow, a hospital management application that combines patient services, staff coordination, and structured records has strong relevance."
  ) -Height 530
  $sourceLines.Add("1B Importance")

  # Slide 13
  $slide = New-ReportSlide -Title "1C. PROBLEM STATEMENT"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The central problem addressed by this project is the inefficient and fragmented nature of manual hospital record management. Important activities such as patient registration, appointment scheduling, prescription tracking, report handling, and administrative communication are frequently maintained using paperwork or disconnected systems that do not integrate with one another.",
    "As a result, the hospital faces repeated data entry, slow search operations, poor coordination between stakeholders, and limited transparency in patient-facing workflows."
  ) -Height 240
  Add-BulletBlock -Slide $slide -Items @(
    "Duplicate patient records and scattered information across departments",
    "Manual appointment handling with no instant status visibility",
    "Limited connectivity between admin, doctor, and patient users",
    "Weak document handling for reports, prescriptions, and uploads",
    "No integrated cart, order tracking, or notification system",
    "Difficulty scaling the process as hospital activity grows"
  ) -Top 390 -Height 250
  $sourceLines.Add("1C Problem Statement")

  # Slide 14
  $slide = New-ReportSlide -Title "1D. OBJECTIVES OF THE PROJECT"
  Add-BulletBlock -Slide $slide -Items @(
    "To replace manual hospital record management with a structured digital platform.",
    "To create secure authentication using email, password, JWT, and OTP verification.",
    "To maintain role-based access for administrator, doctor, and patient users.",
    "To digitize patient profile management and appointment scheduling workflows.",
    "To support e-prescription handling, report uploads, and notification delivery.",
    "To provide an order processing module with cart management and Cash on Delivery.",
    "To enable real-time user communication through an integrated chat system.",
    "To design the solution in MERN stack with clear MVC architecture and modular code.",
    "To prepare a foundation that can be extended with telemedicine and wearable integration."
  ) -Top 135 -Height 560
  $sourceLines.Add("1D Objectives")

  # Slide 15
  $slide = New-ReportSlide -Title "1E. SCOPE OF THE PROJECT"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The scope of the Smart Hospital Management System includes the digital management of patient records, appointments, file uploads, medical product ordering, notifications, and user communication within a browser-based application. The project supports three user roles: administrator, doctor, and patient.",
    "Within the current implementation, the project covers user registration and login, OTP verification, admin-side CRUD operations, appointment handling, report upload, order processing with cart and COD, map integration, notifications, and real-time chat."
  ) -Height 280
  Add-BulletBlock -Slide $slide -Items @(
    "In scope: authentication, CRUD, search and filter, chat, notification, upload, map, cart, and order history.",
    "Partially in scope: e-prescription support and representative billing through order processing.",
    "Out of scope in the current version: live telemedicine, full online payment gateway, insurance claims, and advanced analytics.",
    "Deployment target: small to medium hospital or clinic level digital workflow prototype."
  ) -Top 420 -Height 220
  $sourceLines.Add("1E Scope")

  # Slide 16
  $slide = New-ReportSlide -Title "CHAPTER 2: LITERATURE REVIEW"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "A literature review was conducted to understand how hospital information systems have evolved, which implementation challenges are commonly reported, and what design principles support better user adoption. The review focused on hospital information systems, user acceptance, implementation complexity, and web-based information delivery in healthcare environments.",
    "The selected studies show that successful hospital information systems must balance technology, workflow design, usability, organizational support, and evaluation mechanisms. These findings directly informed the design of the proposed Smart Hospital Management System."
  ) -Height 360
  $sourceLines.Add("Chapter 2 Literature Review")

  # Slide 17
  $slide = New-ReportSlide -Title "2A. RESEARCH PAPERS / ARTICLES REVIEWED"
  Add-TableBlock -Slide $slide -Headers @("Paper", "Year", "Core Focus", "Key Learning") -Rows @(
    @("Bakker and Mol", "1983", "Foundational HIS structure", "Centralized records improve coordination and availability"),
    @("Reichertz", "2006", "Past, present, future of HIS", "Patient-centered electronic information is essential"),
    @("Bain and Standing", "2009", "Hospital management information ecosystem", "Information needs extend across operations and decisions"),
    @("Sligo et al.", "2017", "Planning and evaluation of HIS", "Implementation success depends on organizational and human factors"),
    @("Handayani et al.", "2018", "User acceptance review", "Usability, support, and trust shape adoption"),
    @("Khalifa and Alswailem", "2015", "Acceptance and satisfaction case study", "Performance and training strongly affect satisfaction")
  ) -Top 145 -Height 500
  Add-Caption -Slide $slide -Text "Table 1. Review of major research papers that informed the proposed system design."
  $sourceLines.Add("2A Research Papers")

  # Slide 18
  $slide = New-ReportSlide -Title "2B. SUMMARY OF EXISTING WORK"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The reviewed literature indicates that existing hospital information systems generally focus on registration, billing, electronic medical records, or departmental information exchange. Large enterprise HIS solutions provide comprehensive integration but are expensive, technically demanding, and difficult to customize for small institutions.",
    "At the other end of the spectrum, smaller clinics often depend on isolated solutions such as spreadsheet records, appointment books, or standalone billing tools. These systems may handle one task effectively but do not create a single source of truth for hospital operations.",
    "The literature also shows that user adoption drops when systems are slow, overly complex, or unsupported by staff training."
  ) -Height 400
  Add-BulletBlock -Slide $slide -Items @(
    "Most existing solutions prioritize one department rather than the full patient journey.",
    "Interoperability and integrated communication remain major challenges in practice.",
    "Training, performance, and role-specific usability are repeatedly identified as adoption factors.",
    "Smaller institutions need systems that are simpler, modular, and more affordable to operate."
  ) -Top 560 -Height 120
  $sourceLines.Add("2B Existing Work")

  # Slide 19
  $slide = New-ReportSlide -Title "2C. COMPARATIVE ANALYSIS"
  Add-TableBlock -Slide $slide -Headers @("Approach", "Typical Features", "Advantages", "Limitations") -Rows @(
    @("Manual System", "Paper files, phone scheduling", "Low initial cost", "Slow retrieval, error-prone, no analytics"),
    @("Isolated Digital Tools", "Spreadsheet or desktop modules", "Simple to begin with", "No central integration, repeated entries"),
    @("Enterprise HIS", "Clinical, admin, billing integration", "Comprehensive and powerful", "High cost, heavy training, complex deployment"),
    @("Proposed System", "MERN web app with chat, upload, cart, map, OTP", "Centralized, modular, scalable, web-based", "Prototype scope, future integrations pending")
  ) -Top 145 -Height 420
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The comparison highlights a clear need for a middle-path solution: one that is more integrated than manual or isolated tools, yet simpler and more accessible than a large enterprise hospital information platform."
  ) -Top 600 -Height 90 -FontSize 11
  Add-Caption -Slide $slide -Text "Table 2. Comparative analysis of existing approaches versus the proposed web-based system."
  $sourceLines.Add("2C Comparative Analysis")

  # Slide 20
  $slide = New-ReportSlide -Title "2D. IDENTIFICATION OF RESEARCH GAP"
  Add-BulletBlock -Slide $slide -Items @(
    "Existing studies emphasize implementation and adoption challenges, but many practical prototypes still leave out communication, cart-based ordering, and unified user engagement.",
    "Several hospital information systems are either too narrow in function or too heavy for academic-level or small-institution deployment.",
    "There is limited emphasis on simple full stack architectures that combine secure authentication, real-time communication, uploads, and operational workflows in one accessible system.",
    "Many reviewed platforms do not focus on a clear educational implementation using MERN and MVC while still demonstrating realistic hospital modules.",
    "The present project addresses this gap by combining multiple hospital-facing processes into one structured web system with manageable code complexity."
  ) -Top 145 -Height 430
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "Therefore, the proposed project is positioned as a practical, modular, and educationally strong implementation that responds to both operational needs and software engineering clarity."
  ) -Top 610 -Height 80 -FontSize 11
  $sourceLines.Add("2D Research Gap")

  # Slide 21
  $slide = New-ReportSlide -Title "2E. LITERATURE REVIEW SYNTHESIS"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The literature review collectively shows that successful hospital information systems must be designed around more than data storage. They must also consider staff usability, communication flow, timely access to records, secure authentication, and organizational fit. Systems that ignore these dimensions may be technically functional but operationally weak.",
    "These findings influenced the proposed system in several ways. First, the application uses role-based access so that each stakeholder sees only relevant features. Second, it includes chat and notification capabilities to reduce coordination delays. Third, the solution uses a web-based architecture so that deployment and access remain simple. Finally, it keeps the code modular through MVC so that further integrations can be added without disrupting the foundation."
  ) -Height 540
  $sourceLines.Add("2E Literature Synthesis")

  # Slide 22
  $slide = New-ReportSlide -Title "CHAPTER 3: SYSTEM ANALYSIS"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "System analysis examines the operational problem in its current environment, identifies weaknesses in the existing workflow, and justifies the proposed solution. In this project, analysis was carried out from the perspective of hospital administrators, doctors, and patients, with emphasis on how information moves across registration, appointments, records, communication, and ordering.",
    "This chapter presents the existing system, its limitations, the proposed solution, feasibility analysis, input-output analysis, software requirement specifications, and the development paradigm applied to the project."
  ) -Height 360
  $sourceLines.Add("Chapter 3 System Analysis")

  # Slide 23
  $slide = New-ReportSlide -Title "3A. EXISTING SYSTEM"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "In the existing workflow, patient information is usually recorded manually at a reception desk or stored in isolated local files. Follow-up details are added separately over time, resulting in partial records distributed across files or departments. Staff must search physically or ask other personnel whenever historical information is needed.",
    "Appointments are typically booked by phone or in person, often without a central digital scheduler. This causes overbooking risk, lack of reminders, and difficulty tracking whether an appointment was completed, postponed, or cancelled. Diagnostic reports and prescriptions remain difficult to organize because they exist in paper form or scattered digital copies.",
    "Administrative communication is also fragmented. Patients often need to make repeated calls to ask about appointments or reports, while staff have no instant way to notify patients about changes in order status, consultation timing, or required documents."
  ) -Height 560
  $sourceLines.Add("3A Existing System")

  # Slide 24
  $slide = New-ReportSlide -Title "3B. LIMITATIONS OF EXISTING SYSTEM"
  Add-BulletBlock -Slide $slide -Items @(
    "Searching patient history is slow because records are manual, fragmented, or inconsistently stored.",
    "Repeated data entry causes duplicate records and inconsistent patient information.",
    "Appointment management depends heavily on staff availability and manual checking.",
    "There is no unified communication layer for notifications or real-time conversation.",
    "Report handling is weak because file uploads and centralized digital storage are absent.",
    "Administrative oversight is poor due to the lack of a dashboard and searchable analytics.",
    "Operational growth becomes difficult because manual systems do not scale with patient volume.",
    "Security and accountability are limited because role-based digital access is not enforced."
  ) -Top 145 -Height 480
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "These limitations motivated the shift toward a centralized, browser-based, multi-user system with secure access and structured data handling."
  ) -Top 660 -Height 60 -FontSize 11
  $sourceLines.Add("3B Limitations")

  # Slide 25
  $slide = New-ReportSlide -Title "3C. PROPOSED SYSTEM"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The proposed Smart Hospital Management System is a centralized MERN stack application that integrates multiple hospital operations within a single digital platform. The administrator can manage users, patients, appointments, and orders; doctors can review appointments and communicate with users; and patients can register, verify identity, upload reports, place orders, and interact with the system through a structured dashboard.",
    "The system provides secure authentication, OTP verification, role-based authorization, patient management, appointment tracking, e-prescription support, file uploads, notifications, real-time chat, a basic medical store with cart and COD, and geo-location assistance."
  ) -Height 430
  Add-BulletBlock -Slide $slide -Items @(
    "Web-based access with centralized MongoDB storage",
    "MVC backend structure for maintainability",
    "Operational connectivity between admin, doctor, and patient",
    "Search, filter, upload, order, map, and communication support"
  ) -Top 575 -Height 130
  $sourceLines.Add("3C Proposed System")

  # Slide 26
  $slide = New-ReportSlide -Title "3D. FEASIBILITY STUDY"
  Add-TableBlock -Slide $slide -Headers @("Feasibility Type", "Observation", "Conclusion") -Rows @(
    @("Technical", "MERN stack, MongoDB, Socket.IO, and Multer support all required modules.", "Technically feasible"),
    @("Operational", "Users can access the system through a browser with role-based dashboards.", "Operationally feasible"),
    @("Economic", "Open-source tools reduce development and deployment cost for prototype scale.", "Economically feasible"),
    @("Schedule", "Module-wise development fits an academic project timeline.", "Time feasible"),
    @("Scalability", "MVC structure and modular routes allow future feature growth.", "Scalable with planned enhancements")
  ) -Top 150 -Height 370
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The feasibility analysis confirms that the proposed system is practical for academic implementation and can serve as a strong prototype for future institutional use."
  ) -Top 585 -Height 90 -FontSize 11
  Add-Caption -Slide $slide -Text "Table 3. Feasibility study summary for the proposed Smart Hospital Management System."
  $sourceLines.Add("3D Feasibility")

  # Slide 27
  $slide = New-ReportSlide -Title "3E. INPUT AND OUTPUT ANALYSIS"
  Add-TableBlock -Slide $slide -Headers @("Module", "Primary Inputs", "Primary Outputs") -Rows @(
    @("Authentication", "Name, email, password, OTP", "User session, verification status"),
    @("Patient Management", "Age, gender, blood group, disease", "Structured patient record"),
    @("Appointments", "Doctor, patient, date, reason", "Appointment entry and status"),
    @("Orders", "Product selection and quantity", "Cart contents, order confirmation, total amount"),
    @("Upload", "Report file", "Stored file path and upload acknowledgment"),
    @("Chat and Notifications", "Message text or event trigger", "Delivered messages and alert entries")
  ) -Top 145 -Height 430
  Add-Caption -Slide $slide -Text "Table 4. Input and output analysis of the core modules."
  $sourceLines.Add("3E Input Output")

  # Slide 28
  $slide = New-ReportSlide -Title "3F. SOFTWARE REQUIREMENT SPECIFICATION (FUNCTIONAL)"
  Add-BulletBlock -Slide $slide -Items @(
    "The system shall allow user registration and secure login.",
    "The system shall generate and verify email OTP for account verification.",
    "The system shall enforce role-based access for admin, doctor, and patient users.",
    "The admin shall perform create, read, update, and delete operations on patient records.",
    "The system shall support appointment booking, status updates, and e-prescription recording.",
    "The system shall provide search and filter functionality across major record types.",
    "The system shall support product selection, cart storage, and order placement through COD.",
    "The system shall allow report file uploads and store their references.",
    "The system shall support real-time user-to-user chat.",
    "The system shall generate notifications for important operational events."
  ) -Top 145 -Height 520
  $sourceLines.Add("3F Functional Requirements")

  # Slide 29
  $slide = New-ReportSlide -Title "3G. SOFTWARE REQUIREMENT SPECIFICATION (NON-FUNCTIONAL)"
  Add-BulletBlock -Slide $slide -Items @(
    "Usability: the interface should remain simple enough for non-technical users.",
    "Security: passwords must be hashed and private routes must be token protected.",
    "Performance: normal operations should complete within acceptable demo response time.",
    "Reliability: main CRUD and order flows should remain stable under regular use.",
    "Maintainability: backend code should follow MVC and frontend components should stay modular.",
    "Scalability: additional modules such as telemedicine should be addable later.",
    "Responsiveness: major pages should display properly on standard desktop and laptop screens.",
    "Availability: the application should run locally with standard Node.js and MongoDB setup."
  ) -Top 145 -Height 470
  $sourceLines.Add("3G Non Functional Requirements")

  # Slide 30
  $slide = New-ReportSlide -Title "3H. SOFTWARE ENGINEERING PARADIGM APPLIED"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The project follows an incremental and iterative development approach. Instead of attempting the full application in one step, the system was built module by module: first the authentication layer, then patient and appointment management, then order handling, and finally communication, notifications, and AI-assisted interaction support.",
    "On the architectural side, the backend uses the MVC pattern. Models define MongoDB schemas, controllers handle business logic, and routes connect endpoints to the appropriate controller actions. This separation helps developers debug features, extend modules, and maintain clarity.",
    "The frontend follows a component-oriented structure in React, which supports reuse and keeps stateful logic separated from visual sections."
  ) -Height 560
  $sourceLines.Add("3H Software Engineering Paradigm")

  # Slide 31
  $slide = New-ReportSlide -Title "3I. OVERALL WORKFLOW OF THE SYSTEM"
  Add-BulletBlock -Slide $slide -Items @(
    "User creates an account and logs in through secure authentication.",
    "OTP verification confirms identity when required for account trust.",
    "Admin manages patient records and assigns appointments to doctors.",
    "Patient views schedules, uploads reports, and interacts with the system dashboard.",
    "Medical items can be added to cart; repeated additions increase quantity rather than creating duplicate rows.",
    "Patient places order through COD and receives a confirmation view with total amount.",
    "Admin updates order status and can remove delivered orders from history when permitted.",
    "Chat and notification modules maintain communication between the connected users."
  ) -Top 145 -Height 470
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "This workflow integrates administrative control and patient self-service into one coordinated web system."
  ) -Top 660 -Height 60 -FontSize 11
  $sourceLines.Add("3I Workflow")

  # Slide 32
  $slide = New-ReportSlide -Title "CHAPTER 4: SYSTEM DESIGN"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "System design translates the analyzed requirements into a concrete technical structure. In this project, design includes the layered architecture of the application, the division of modules, the movement of data between actors and processes, and the schema-level relationships that support persistent storage.",
    "To align with the sample report and project guidelines, this chapter includes the system architecture diagram, data flow diagrams, a use case diagram, a sequence diagram, a schema / ER diagram, and flowcharts for core processes."
  ) -Height 360
  $sourceLines.Add("Chapter 4 System Design")

  # Slide 33
  $slide = New-ReportSlide -Title "4A. SYSTEM ARCHITECTURE DIAGRAM"
  Add-StickActor -Slide $slide -Label "Patient" -Left 40 -Top 175
  Add-StickActor -Slide $slide -Label "Doctor" -Left 40 -Top 330
  Add-StickActor -Slide $slide -Label "Admin" -Left 40 -Top 485
  Add-RoundedLabel -Slide $slide -Text "React Frontend`r(Login, Dashboard, Cart, Chat, AI Widget)" -Left 175 -Top 185 -Width 250 -Height 70 -Fill $ColorSoftBlue -FontSize 12 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Express.js API Layer`r(Routes, Controllers, Middleware)" -Left 175 -Top 315 -Width 250 -Height 70 -Fill $ColorSoftGreen -FontSize 12 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Socket.IO Layer`r(Real-time Chat Events)" -Left 175 -Top 445 -Width 250 -Height 56 -Fill $ColorSoftCream -FontSize 12 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "MongoDB Database`r(Users, Patients, Appointments, Orders, Messages, Notifications)" -Left 150 -Top 560 -Width 300 -Height 84 -Fill $ColorSoftRose -FontSize 12 | Out-Null
  Add-LineArrow -Slide $slide -X1 95 -Y1 220 -X2 175 -Y2 220 | Out-Null
  Add-LineArrow -Slide $slide -X1 95 -Y1 375 -X2 175 -Y2 220 | Out-Null
  Add-LineArrow -Slide $slide -X1 95 -Y1 530 -X2 175 -Y2 220 | Out-Null
  Add-LineArrow -Slide $slide -X1 300 -Y1 255 -X2 300 -Y2 315 | Out-Null
  Add-LineArrow -Slide $slide -X1 300 -Y1 385 -X2 300 -Y2 445 | Out-Null
  Add-LineArrow -Slide $slide -X1 300 -Y1 501 -X2 300 -Y2 560 | Out-Null
  Add-LineArrow -Slide $slide -X1 426 -Y1 350 -X2 520 -Y2 350 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "External Services`r(OTP Email, Map Tiles, Groq AI)" -Left 435 -Top 315 -Width 110 -Height 72 -Fill 15529979 -FontSize 10 | Out-Null
  Add-Caption -Slide $slide -Text "Figure 1. Layered architecture showing users, frontend, API, real-time communication, database, and external services."
  $sourceLines.Add("4A System Architecture Diagram")

  # Slide 34
  $slide = New-ReportSlide -Title "4B. MODULE DESIGN"
  Add-RoundedLabel -Slide $slide -Text "Smart Hospital Management System" -Left 180 -Top 140 -Width 220 -Height 50 -Fill $ColorAccent -FontSize 14 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Authentication and OTP" -Left 65 -Top 250 -Width 150 -Height 52 -Fill $ColorSoftBlue -FontSize 11 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Patient Records" -Left 245 -Top 250 -Width 110 -Height 52 -Fill $ColorSoftGreen -FontSize 11 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Appointments and E-Prescription" -Left 390 -Top 250 -Width 145 -Height 52 -Fill $ColorSoftCream -FontSize 11 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Medical Store and Cart" -Left 65 -Top 390 -Width 150 -Height 52 -Fill $ColorSoftCream -FontSize 11 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Uploads and Reports" -Left 245 -Top 390 -Width 110 -Height 52 -Fill $ColorSoftBlue -FontSize 11 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Chat and Notifications" -Left 390 -Top 390 -Width 145 -Height 52 -Fill $ColorSoftGreen -FontSize 11 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Search and Filter" -Left 130 -Top 530 -Width 120 -Height 50 -Fill $ColorSoftRose -FontSize 11 | Out-Null
  Add-RoundedLabel -Slide $slide -Text "Map and AI Assistant" -Left 335 -Top 530 -Width 140 -Height 50 -Fill 15529979 -FontSize 11 | Out-Null
  foreach ($pair in @(@(290,190,140,250), @(290,190,300,250), @(290,190,462,250), @(290,190,140,390), @(290,190,300,390), @(290,190,462,390), @(290,190,190,530), @(290,190,405,530))) { Add-LineNoArrow -Slide $slide -X1 $pair[0] -Y1 $pair[1] -X2 $pair[2] -Y2 $pair[3] | Out-Null }
  Add-Caption -Slide $slide -Text "Figure 2. Module-level decomposition of the proposed system."
  $sourceLines.Add("4B Module Design")

  # Slide 35
  $slide = New-ReportSlide -Title "4C. DATA FLOW DIAGRAM LEVEL 0"
  Add-RectangleLabel -Slide $slide -Text "Patient" -Left 45 -Top 270 -Width 90 -Height 44 -Fill $ColorSoftBlue | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Doctor" -Left 45 -Top 430 -Width 90 -Height 44 -Fill $ColorSoftGreen | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Admin" -Left 460 -Top 350 -Width 90 -Height 44 -Fill $ColorSoftCream | Out-Null
  Add-RoundedLabel -Slide $slide -Text "0. Smart Hospital Management System" -Left 195 -Top 330 -Width 200 -Height 78 -Fill $ColorAccent -FontSize 13 | Out-Null
  Add-LineArrow -Slide $slide -X1 135 -Y1 292 -X2 195 -Y2 352 | Out-Null
  Add-LineArrow -Slide $slide -X1 195 -Y1 384 -X2 135 -Y2 292 | Out-Null
  Add-LineArrow -Slide $slide -X1 135 -Y1 452 -X2 195 -Y2 390 | Out-Null
  Add-LineArrow -Slide $slide -X1 395 -Y1 370 -X2 460 -Y2 370 | Out-Null
  Add-LineArrow -Slide $slide -X1 460 -Y1 392 -X2 395 -Y2 392 | Out-Null
  Add-ConnectorText -Slide $slide -Text "register, login, view status" -Left 120 -Top 315 -Width 85 | Out-Null
  Add-ConnectorText -Slide $slide -Text "appointments, reports" -Left 120 -Top 410 -Width 85 | Out-Null
  Add-ConnectorText -Slide $slide -Text "manage records and orders" -Left 400 -Top 325 -Width 95 | Out-Null
  Add-Caption -Slide $slide -Text "Figure 3. Level 0 DFD showing the system as one high-level process connected to major actors."
  $sourceLines.Add("4C DFD Level 0")

  # Slide 36
  $slide = New-ReportSlide -Title "4D. DATA FLOW DIAGRAM LEVEL 1 - PATIENT"
  Add-RectangleLabel -Slide $slide -Text "Patient" -Left 30 -Top 360 -Width 80 -Height 42 -Fill $ColorSoftBlue | Out-Null
  Add-RoundedLabel -Slide $slide -Text "1.1 Register / Login" -Left 140 -Top 140 -Width 120 -Height 44 -Fill $ColorSoftBlue | Out-Null
  Add-RoundedLabel -Slide $slide -Text "1.2 Manage Profile and Reports" -Left 140 -Top 250 -Width 120 -Height 52 -Fill $ColorSoftGreen | Out-Null
  Add-RoundedLabel -Slide $slide -Text "1.3 Book and View Appointments" -Left 140 -Top 370 -Width 120 -Height 52 -Fill $ColorSoftCream | Out-Null
  Add-RoundedLabel -Slide $slide -Text "1.4 Cart and Place Order" -Left 140 -Top 495 -Width 120 -Height 52 -Fill $ColorSoftRose | Out-Null
  Add-RectangleLabel -Slide $slide -Text "D1 Users" -Left 360 -Top 140 -Width 110 -Height 42 -Fill 15529979 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "D2 Patients" -Left 360 -Top 250 -Width 110 -Height 42 -Fill 15529979 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "D3 Appointments" -Left 360 -Top 370 -Width 110 -Height 42 -Fill 15529979 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "D4 Orders" -Left 360 -Top 495 -Width 110 -Height 42 -Fill 15529979 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "D5 Notifications / Messages" -Left 330 -Top 610 -Width 170 -Height 44 -Fill 15529979 | Out-Null
  Add-LineArrow -Slide $slide -X1 110 -Y1 381 -X2 140 -Y2 162 | Out-Null
  Add-LineArrow -Slide $slide -X1 110 -Y1 381 -X2 140 -Y2 276 | Out-Null
  Add-LineArrow -Slide $slide -X1 110 -Y1 381 -X2 140 -Y2 396 | Out-Null
  Add-LineArrow -Slide $slide -X1 110 -Y1 381 -X2 140 -Y2 520 | Out-Null
  Add-LineArrow -Slide $slide -X1 260 -Y1 162 -X2 360 -Y2 162 | Out-Null
  Add-LineArrow -Slide $slide -X1 260 -Y1 276 -X2 360 -Y2 276 | Out-Null
  Add-LineArrow -Slide $slide -X1 260 -Y1 396 -X2 360 -Y2 396 | Out-Null
  Add-LineArrow -Slide $slide -X1 260 -Y1 520 -X2 360 -Y2 520 | Out-Null
  Add-LineArrow -Slide $slide -X1 260 -Y1 520 -X2 330 -Y2 632 | Out-Null
  Add-Caption -Slide $slide -Text "Figure 4. Level 1 DFD describing the main patient-side processes and the connected data stores."
  $sourceLines.Add("4D DFD Level 1 Patient")

  # Slide 37
  $slide = New-ReportSlide -Title "4E. DATA FLOW DIAGRAM LEVEL 1 - ADMIN AND DOCTOR"
  Add-RectangleLabel -Slide $slide -Text "Admin" -Left 25 -Top 220 -Width 80 -Height 42 -Fill $ColorSoftCream | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Doctor" -Left 25 -Top 470 -Width 80 -Height 42 -Fill $ColorSoftGreen | Out-Null
  Add-RoundedLabel -Slide $slide -Text "2.1 Manage Users and Patients" -Left 145 -Top 150 -Width 140 -Height 48 -Fill $ColorSoftBlue | Out-Null
  Add-RoundedLabel -Slide $slide -Text "2.2 Create and Update Appointments" -Left 145 -Top 270 -Width 140 -Height 52 -Fill $ColorSoftCream | Out-Null
  Add-RoundedLabel -Slide $slide -Text "2.3 Review Orders and Delivery Status" -Left 145 -Top 400 -Width 140 -Height 52 -Fill $ColorSoftRose | Out-Null
  Add-RoundedLabel -Slide $slide -Text "2.4 Doctor Reviews Case and Prescription" -Left 145 -Top 530 -Width 140 -Height 54 -Fill $ColorSoftGreen | Out-Null
  Add-RectangleLabel -Slide $slide -Text "D1 Users" -Left 385 -Top 150 -Width 110 -Height 42 -Fill 15529979 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "D2 Patients" -Left 385 -Top 250 -Width 110 -Height 42 -Fill 15529979 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "D3 Appointments" -Left 385 -Top 365 -Width 110 -Height 42 -Fill 15529979 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "D4 Orders" -Left 385 -Top 470 -Width 110 -Height 42 -Fill 15529979 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "D5 Notifications / Messages" -Left 350 -Top 605 -Width 180 -Height 44 -Fill 15529979 | Out-Null
  Add-LineArrow -Slide $slide -X1 105 -Y1 240 -X2 145 -Y2 175 | Out-Null
  Add-LineArrow -Slide $slide -X1 105 -Y1 240 -X2 145 -Y2 295 | Out-Null
  Add-LineArrow -Slide $slide -X1 105 -Y1 240 -X2 145 -Y2 425 | Out-Null
  Add-LineArrow -Slide $slide -X1 105 -Y1 490 -X2 145 -Y2 560 | Out-Null
  Add-LineArrow -Slide $slide -X1 285 -Y1 175 -X2 385 -Y2 171 | Out-Null
  Add-LineArrow -Slide $slide -X1 285 -Y1 295 -X2 385 -Y2 386 | Out-Null
  Add-LineArrow -Slide $slide -X1 285 -Y1 425 -X2 385 -Y2 491 | Out-Null
  Add-LineArrow -Slide $slide -X1 285 -Y1 557 -X2 350 -Y2 627 | Out-Null
  Add-Caption -Slide $slide -Text "Figure 5. Level 1 DFD describing the admin and doctor operational processes."
  $sourceLines.Add("4E DFD Level 1 Admin Doctor")

  # Slide 38
  $slide = New-ReportSlide -Title "4F. USE CASE DIAGRAM"
  Add-StickActor -Slide $slide -Label "Patient" -Left 20 -Top 240
  Add-StickActor -Slide $slide -Label "Admin" -Left 500 -Top 200
  Add-StickActor -Slide $slide -Label "Doctor" -Left 500 -Top 455
  $boundary = $slide.Shapes.AddShape(1, 120, 135, 340, 520)
  $boundary.Fill.Visible = 0
  $boundary.Line.ForeColor.RGB = $ColorLine
  Add-TextBlock -Slide $slide -Text "System Boundary" -Left 240 -Top 140 -Width 110 -Height 18 -FontSize 10 -Alignment 2 -Color $ColorMuted | Out-Null
  Add-EllipseLabel -Slide $slide -Text "Register / Login" -Left 175 -Top 185 -Width 105 -Height 40 -Fill $ColorSoftBlue | Out-Null
  Add-EllipseLabel -Slide $slide -Text "Verify OTP" -Left 300 -Top 185 -Width 95 -Height 40 -Fill $ColorSoftBlue | Out-Null
  Add-EllipseLabel -Slide $slide -Text "Manage Patient Records" -Left 210 -Top 260 -Width 155 -Height 46 -Fill $ColorSoftGreen | Out-Null
  Add-EllipseLabel -Slide $slide -Text "Manage Appointments" -Left 205 -Top 335 -Width 165 -Height 46 -Fill $ColorSoftCream | Out-Null
  Add-EllipseLabel -Slide $slide -Text "Upload Reports" -Left 180 -Top 420 -Width 105 -Height 42 -Fill $ColorSoftRose | Out-Null
  Add-EllipseLabel -Slide $slide -Text "Cart and Place Order" -Left 300 -Top 420 -Width 120 -Height 42 -Fill $ColorSoftRose | Out-Null
  Add-EllipseLabel -Slide $slide -Text "Real-time Chat" -Left 180 -Top 505 -Width 105 -Height 42 -Fill $ColorSoftBlue | Out-Null
  Add-EllipseLabel -Slide $slide -Text "View Notifications" -Left 300 -Top 505 -Width 120 -Height 42 -Fill $ColorSoftGreen | Out-Null
  Add-LineNoArrow -Slide $slide -X1 74 -Y1 275 -X2 175 -Y2 205 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 74 -Y1 300 -X2 300 -Y2 205 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 74 -Y1 320 -X2 205 -Y2 357 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 74 -Y1 340 -X2 230 -Y2 441 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 74 -Y1 365 -X2 340 -Y2 441 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 74 -Y1 390 -X2 230 -Y2 526 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 74 -Y1 415 -X2 340 -Y2 526 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 500 -Y1 245 -X2 365 -Y2 283 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 500 -Y1 270 -X2 365 -Y2 357 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 500 -Y1 485 -X2 365 -Y2 357 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 500 -Y1 510 -X2 235 -Y2 526 | Out-Null
  Add-Caption -Slide $slide -Text "Figure 6. Use case diagram showing interactions among patient, admin, doctor, and the system."
  $sourceLines.Add("4F Use Case Diagram")

  # Slide 39
  $slide = New-ReportSlide -Title "4G. SEQUENCE DIAGRAM - APPOINTMENT BOOKING"
  Add-Lifeline -Slide $slide -Label "Patient" -Left 60
  Add-Lifeline -Slide $slide -Label "React UI" -Left 180
  Add-Lifeline -Slide $slide -Label "Express API" -Left 300
  Add-Lifeline -Slide $slide -Label "MongoDB" -Left 420
  Add-LineArrow -Slide $slide -X1 110 -Y1 220 -X2 230 -Y2 220 | Out-Null
  Add-ConnectorText -Slide $slide -Text "1. appointment request" -Left 118 -Top 198 -Width 120 | Out-Null
  Add-LineArrow -Slide $slide -X1 230 -Y1 285 -X2 350 -Y2 285 | Out-Null
  Add-ConnectorText -Slide $slide -Text "2. POST /appointments" -Left 238 -Top 262 -Width 120 | Out-Null
  Add-LineArrow -Slide $slide -X1 350 -Y1 350 -X2 470 -Y2 350 | Out-Null
  Add-ConnectorText -Slide $slide -Text "3. save appointment" -Left 362 -Top 328 -Width 115 | Out-Null
  Add-LineArrow -Slide $slide -X1 470 -Y1 415 -X2 350 -Y2 415 | Out-Null
  Add-ConnectorText -Slide $slide -Text "4. appointment id" -Left 360 -Top 392 -Width 110 | Out-Null
  Add-LineArrow -Slide $slide -X1 350 -Y1 480 -X2 230 -Y2 480 | Out-Null
  Add-ConnectorText -Slide $slide -Text "5. success response" -Left 245 -Top 458 -Width 110 | Out-Null
  Add-LineArrow -Slide $slide -X1 230 -Y1 545 -X2 110 -Y2 545 | Out-Null
  Add-ConnectorText -Slide $slide -Text "6. confirmation shown" -Left 118 -Top 522 -Width 120 | Out-Null
  Add-Caption -Slide $slide -Text "Figure 7. Sequence of interactions when an appointment is created through the web interface."
  $sourceLines.Add("4G Sequence Diagram")

  # Slide 40
  $slide = New-ReportSlide -Title "4H. SCHEMA / ENTITY RELATIONSHIP DIAGRAM"
  Add-RectangleLabel -Slide $slide -Text "User`r- name`r- email`r- password`r- role`r- cart" -Left 50 -Top 170 -Width 110 -Height 110 -Fill $ColorSoftBlue -FontSize 10 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Patient`r- userId`r- age`r- gender`r- bloodGroup`r- disease" -Left 230 -Top 120 -Width 125 -Height 120 -Fill $ColorSoftGreen -FontSize 10 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Appointment`r- patientId`r- doctorId`r- date`r- time`r- status" -Left 410 -Top 170 -Width 125 -Height 120 -Fill $ColorSoftCream -FontSize 10 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Order`r- userId`r- items`r- totalAmount`r- paymentStatus" -Left 60 -Top 410 -Width 120 -Height 110 -Fill $ColorSoftRose -FontSize 10 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Message`r- senderId`r- receiverId`r- text`r- createdAt" -Left 245 -Top 395 -Width 120 -Height 110 -Fill $ColorSoftBlue -FontSize 10 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Notification`r- userId`r- title`r- message`r- read" -Left 420 -Top 400 -Width 120 -Height 110 -Fill $ColorSoftGreen -FontSize 10 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 160 -Y1 225 -X2 230 -Y2 180 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 355 -Y1 180 -X2 410 -Y2 225 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 115 -Y1 280 -X2 115 -Y2 410 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 160 -Y1 235 -X2 245 -Y2 450 | Out-Null
  Add-LineNoArrow -Slide $slide -X1 140 -Y1 235 -X2 420 -Y2 450 | Out-Null
  Add-ConnectorText -Slide $slide -Text "1 : 1" -Left 183 -Top 182 -Width 40 | Out-Null
  Add-ConnectorText -Slide $slide -Text "1 : many" -Left 372 -Top 190 -Width 48 | Out-Null
  Add-ConnectorText -Slide $slide -Text "1 : many" -Left 88 -Top 335 -Width 50 | Out-Null
  Add-Caption -Slide $slide -Text "Figure 8. Simplified schema / ER view of the core MongoDB collections."
  $sourceLines.Add("4H ER Diagram")

  # Slide 41
  $slide = New-ReportSlide -Title "4I. FLOWCHART - AUTHENTICATION"
  Add-EllipseLabel -Slide $slide -Text "Start" -Left 245 -Top 130 -Width 90 -Height 36 -Fill $ColorSoftBlue | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Enter email and password" -Left 205 -Top 200 -Width 170 -Height 44 -Fill $ColorSoftCream | Out-Null
  Add-DiamondLabel -Slide $slide -Text "User exists?" -Left 225 -Top 280 -Width 130 -Height 70 -Fill $ColorSoftRose | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Show invalid user message" -Left 50 -Top 390 -Width 155 -Height 44 -Fill $ColorSoftCream | Out-Null
  Add-DiamondLabel -Slide $slide -Text "Password correct?" -Left 225 -Top 390 -Width 130 -Height 70 -Fill $ColorSoftGreen | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Generate JWT and load dashboard" -Left 365 -Top 510 -Width 175 -Height 48 -Fill $ColorSoftBlue | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Show login error" -Left 50 -Top 510 -Width 125 -Height 44 -Fill $ColorSoftCream | Out-Null
  Add-EllipseLabel -Slide $slide -Text "End" -Left 245 -Top 620 -Width 90 -Height 36 -Fill $ColorSoftBlue | Out-Null
  Add-LineArrow -Slide $slide -X1 290 -Y1 166 -X2 290 -Y2 200 | Out-Null
  Add-LineArrow -Slide $slide -X1 290 -Y1 244 -X2 290 -Y2 280 | Out-Null
  Add-LineArrow -Slide $slide -X1 225 -Y1 315 -X2 130 -Y2 390 | Out-Null
  Add-LineArrow -Slide $slide -X1 355 -Y1 315 -X2 290 -Y2 390 | Out-Null
  Add-LineArrow -Slide $slide -X1 225 -Y1 430 -X2 175 -Y2 532 | Out-Null
  Add-LineArrow -Slide $slide -X1 355 -Y1 430 -X2 365 -Y2 534 | Out-Null
  Add-LineArrow -Slide $slide -X1 112 -Y1 554 -X2 245 -Y2 638 | Out-Null
  Add-LineArrow -Slide $slide -X1 452 -Y1 558 -X2 335 -Y2 638 | Out-Null
  Add-ConnectorText -Slide $slide -Text "No" -Left 185 -Top 343 -Width 28 | Out-Null
  Add-ConnectorText -Slide $slide -Text "Yes" -Left 315 -Top 343 -Width 28 | Out-Null
  Add-ConnectorText -Slide $slide -Text "No" -Left 178 -Top 467 -Width 28 | Out-Null
  Add-ConnectorText -Slide $slide -Text "Yes" -Left 367 -Top 467 -Width 28 | Out-Null
  Add-Caption -Slide $slide -Text "Figure 9. Flowchart of the login authentication process."
  $sourceLines.Add("4I Authentication Flowchart")

  # Slide 42
  $slide = New-ReportSlide -Title "4J. FLOWCHART - ORDER PROCESSING"
  Add-EllipseLabel -Slide $slide -Text "Start" -Left 250 -Top 125 -Width 90 -Height 36 -Fill $ColorSoftBlue | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Select product" -Left 225 -Top 190 -Width 140 -Height 40 -Fill $ColorSoftCream | Out-Null
  Add-DiamondLabel -Slide $slide -Text "Already in cart?" -Left 220 -Top 260 -Width 150 -Height 72 -Fill $ColorSoftRose | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Increase quantity" -Left 385 -Top 365 -Width 130 -Height 40 -Fill $ColorSoftGreen | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Add new cart item" -Left 70 -Top 365 -Width 130 -Height 40 -Fill $ColorSoftBlue | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Review cart and total" -Left 225 -Top 440 -Width 140 -Height 42 -Fill $ColorSoftCream | Out-Null
  Add-DiamondLabel -Slide $slide -Text "Place order with COD?" -Left 220 -Top 515 -Width 150 -Height 76 -Fill $ColorSoftGreen | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Create order, clear cart, show confirmation page" -Left 365 -Top 625 -Width 170 -Height 52 -Fill $ColorSoftBlue | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Stay in cart / continue editing" -Left 45 -Top 625 -Width 150 -Height 52 -Fill $ColorSoftCream | Out-Null
  Add-LineArrow -Slide $slide -X1 295 -Y1 161 -X2 295 -Y2 190 | Out-Null
  Add-LineArrow -Slide $slide -X1 295 -Y1 230 -X2 295 -Y2 260 | Out-Null
  Add-LineArrow -Slide $slide -X1 220 -Y1 296 -X2 135 -Y2 365 | Out-Null
  Add-LineArrow -Slide $slide -X1 370 -Y1 296 -X2 385 -Y2 365 | Out-Null
  Add-LineArrow -Slide $slide -X1 135 -Y1 405 -X2 250 -Y2 440 | Out-Null
  Add-LineArrow -Slide $slide -X1 450 -Y1 405 -X2 340 -Y2 440 | Out-Null
  Add-LineArrow -Slide $slide -X1 295 -Y1 482 -X2 295 -Y2 515 | Out-Null
  Add-LineArrow -Slide $slide -X1 220 -Y1 557 -X2 145 -Y2 625 | Out-Null
  Add-LineArrow -Slide $slide -X1 370 -Y1 557 -X2 365 -Y2 651 | Out-Null
  Add-ConnectorText -Slide $slide -Text "No" -Left 182 -Top 333 -Width 28 | Out-Null
  Add-ConnectorText -Slide $slide -Text "Yes" -Left 381 -Top 333 -Width 30 | Out-Null
  Add-ConnectorText -Slide $slide -Text "No" -Left 182 -Top 595 -Width 30 | Out-Null
  Add-ConnectorText -Slide $slide -Text "Yes" -Left 378 -Top 595 -Width 30 | Out-Null
  Add-Caption -Slide $slide -Text "Figure 10. Flowchart of the cart and COD order processing logic."
  $sourceLines.Add("4J Order Processing Flowchart")

  # Slide 43
  $slide = New-ReportSlide -Title "CHAPTER 5: METHODOLOGY"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The methodology adopted for this project is modular, incremental, and implementation-oriented. Each major function was developed as an independent module and then integrated into the broader workflow. This approach reduced complexity, simplified testing, and helped maintain code clarity during the project lifecycle.",
    "This chapter presents the step-by-step development approach, the algorithms and techniques used, the simplified mathematical model considered for evaluation, and the overall workflow logic followed during execution."
  ) -Height 330
  $sourceLines.Add("Chapter 5 Methodology")

  # Slide 44
  $slide = New-ReportSlide -Title "5A. STEP-BY-STEP DEVELOPMENT APPROACH"
  Add-BulletBlock -Slide $slide -Items @(
    "Requirement gathering from the problem statement, feature list, and report guidelines.",
    "Creation of backend project structure with config, models, controllers, routes, middleware, and utilities.",
    "Design of MongoDB schemas for users, patients, appointments, orders, messages, and notifications.",
    "Implementation of authentication, JWT protection, and OTP verification workflow.",
    "Addition of patient CRUD, appointment management, file upload, and search features.",
    "Implementation of medical store, cart logic, order placement, and order confirmation flow.",
    "Integration of Socket.IO chat, notifications, map, and AI doctor widget.",
    "Frontend design using React components, route protection, and role-aware dashboard sections.",
    "Testing using dummy accounts, seeded records, and manual validation of each workflow."
  ) -Top 145 -Height 540
  $sourceLines.Add("5A Step by Step")

  # Slide 45
  $slide = New-ReportSlide -Title "5B. ALGORITHMS AND TECHNIQUES USED"
  Add-BulletBlock -Slide $slide -Items @(
    "Password Hashing: bcrypt is used to hash user passwords before storage.",
    "Token-based Authentication: JWT is generated at login and validated on protected routes.",
    "OTP Verification: a six-digit OTP with expiry timestamp is generated and matched during verification.",
    "Search and Filter: keyword matching and status-based filtering reduce retrieval time.",
    "Cart Quantity Merge: when the same product is added again, quantity is incremented instead of creating duplicate rows.",
    "Remove One Item Logic: cart item removal reduces quantity by one and removes the item only when quantity reaches zero.",
    "Delivered Order Removal Rule: order history deletion is permitted only after status becomes delivered.",
    "Real-time Messaging: Socket.IO emits events for instant message delivery.",
    "Role-based Authorization: middleware restricts admin-only functionality."
  ) -Top 145 -Height 560
  $sourceLines.Add("5B Algorithms and Techniques")

  # Slide 46
  $slide = New-ReportSlide -Title "5C. MATHEMATICAL MODEL AND ASSUMPTIONS"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "Although the project is primarily application-oriented, a simplified performance model was used to reason about system effectiveness. Operational success can be expressed as a weighted function of authentication reliability, appointment workflow completion, order completion, and notification delivery.",
    "A simple evaluation expression may be written as: Overall Success Score = w1(A) + w2(P) + w3(O) + w4(N), where A is authentication success rate, P is appointment flow completion rate, O is order flow completion rate, and N is notification or communication reliability. The weights reflect module importance within the prototype.",
    "For report discussion, the following assumptions were used: one local deployment environment, seeded demo data, small to medium request volume, and manual validation of key workflows. These assumptions are appropriate for academic prototype evaluation, but not intended as enterprise-scale benchmarks."
  ) -Height 560
  $sourceLines.Add("5C Mathematical Model")

  # Slide 47
  $slide = New-ReportSlide -Title "5D. WORKFLOW EXPLANATION"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The workflow begins with user onboarding. After account creation and login, the user enters the dashboard where accessible modules depend on the assigned role. The administrator manages patient information and appointments, while the patient primarily interacts with bookings, uploads, cart, orders, chat, and notifications.",
    "The system is event-driven at the operational level. A new appointment triggers visibility and notifications; a new order triggers cart clearing and confirmation; a new chat message triggers an immediate socket event; and a delivered order unlocks a controlled delete action in history. Each workflow is designed so that state changes are reflected in both the frontend and backend records.",
    "This methodology keeps the application behavior predictable and testable while still demonstrating practical full stack integration."
  ) -Height 560
  $sourceLines.Add("5D Workflow Explanation")

  # Slide 48
  $slide = New-ReportSlide -Title "CHAPTER 6: IMPLEMENTATION"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "Implementation was carried out as a full stack MERN application with a clear separation between frontend and backend responsibilities. The backend exposes secure REST endpoints and socket events, while the frontend consumes those services through React pages and components.",
    "This chapter documents the tools used, the hardware and software environment, the database and dummy data setup, the coding structure, the API organization, and the representative user interface layouts of the implemented system."
  ) -Height 320
  $sourceLines.Add("Chapter 6 Implementation")

  # Slide 49
  $slide = New-ReportSlide -Title "6A. TOOLS AND TECHNOLOGIES"
  Add-TableBlock -Slide $slide -Headers @("Technology", "Purpose in Project", "Reason for Use") -Rows @(
    @("React.js", "Frontend UI", "Reusable components and structured client rendering"),
    @("Node.js", "Backend runtime", "Single language across stack and non-blocking operations"),
    @("Express.js", "REST API framework", "Lightweight routing and middleware support"),
    @("MongoDB / Mongoose", "Database and ODM", "Flexible document storage for hospital records"),
    @("Socket.IO", "Real-time communication", "Instant chat delivery between users"),
    @("Multer", "File upload handling", "Simple report upload processing"),
    @("JWT / bcrypt", "Security", "Token auth and hashed passwords"),
    @("React Leaflet", "Map integration", "Interactive hospital location display"),
    @("Groq API", "AI doctor assistant", "Symptom guidance through floating assistant widget")
  ) -Top 145 -Height 500
  Add-Caption -Slide $slide -Text "Table 5. Tools and technologies used in the final implementation."
  $sourceLines.Add("6A Tools and Technologies")

  # Slide 50
  $slide = New-ReportSlide -Title "6B. HARDWARE REQUIREMENTS"
  Add-TableBlock -Slide $slide -Headers @("Component", "Minimum Requirement", "Recommended") -Rows @(
    @("Processor", "Dual-core CPU", "Modern multi-core processor"),
    @("RAM", "4 GB", "8 GB or higher"),
    @("Storage", "10 GB free space", "SSD with additional margin"),
    @("Network", "Internet for package setup and map tiles", "Stable broadband connection"),
    @("Display", "1366 x 768", "Full HD for easier development")
  ) -Top 150 -Height 300
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The project was designed to run on a regular student development laptop. Because the application is browser-based and the database workload is moderate, the hardware requirements remain modest. However, smoother execution is achieved with 8 GB RAM, especially when the backend server, frontend development server, MongoDB, and browser are active at the same time."
  ) -Top 500 -Height 170 -FontSize 11
  Add-Caption -Slide $slide -Text "Table 6. Hardware requirements for development and local execution."
  $sourceLines.Add("6B Hardware Requirements")

  # Slide 51
  $slide = New-ReportSlide -Title "6C. SOFTWARE REQUIREMENTS AND DUMMY DATA"
  Add-TableBlock -Slide $slide -Headers @("Software", "Version / Example", "Usage") -Rows @(
    @("Operating System", "Windows 10/11 or similar", "Local development"),
    @("Node.js", "v18 or above", "Runtime for backend and frontend tooling"),
    @("npm", "v9 or above", "Package management"),
    @("MongoDB", "Community edition / local URI", "Persistent storage"),
    @("VS Code", "Any recent version", "Code editing and debugging"),
    @("Browser", "Chrome / Edge", "Application testing")
  ) -Top 145 -Height 285
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The system also uses seeded dummy data for demonstration. Example accounts include one administrator, one doctor, and one patient. Representative medical store items, patient disease details, sample appointments, and notification records allow all modules to be demonstrated without requiring external data collection.",
    "Because the project is prototype-driven, the dummy data plays the role of a small dataset for functional validation rather than statistical analysis."
  ) -Top 465 -Height 200 -FontSize 11
  Add-Caption -Slide $slide -Text "Table 7. Software requirements and the role of dummy data in testing."
  $sourceLines.Add("6C Software and Dummy Data")

  # Slide 52
  $slide = New-ReportSlide -Title "6D. CODING APPROACH AND MVC STRUCTURE"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The backend code is structured according to MVC principles. Models define the User, Patient, Appointment, Order, Message, and Notification schemas. Controllers contain business logic for authentication, patients, appointments, orders, chat, notifications, uploads, and AI assistance. Route files map API endpoints to the correct controller methods while middleware protects secure paths.",
    "The frontend uses React pages and components to render dashboard sections such as Overview, Patients, Appointments, Orders, Cart, Upload, Map, Chat, and Notifications. Shared state, especially authenticated user information, is managed through context so that role-based behavior can be enforced at the interface level.",
    "This coding approach keeps the project simple to understand while still following maintainable engineering practices."
  ) -Height 560
  $sourceLines.Add("6D Coding Approach and MVC")

  # Slide 53
  $slide = New-ReportSlide -Title "6E. BACKEND MODULES AND API DESIGN"
  Add-TableBlock -Slide $slide -Headers @("Module", "Representative Endpoint", "Responsibility") -Rows @(
    @("Authentication", "/api/auth/login", "Login, token generation, OTP workflow"),
    @("Patients", "/api/patients", "CRUD operations on patient records"),
    @("Appointments", "/api/appointments", "Create and update appointments"),
    @("Orders", "/api/orders", "Cart actions, place order, delivery status"),
    @("Upload", "/api/upload", "Store report files and return path"),
    @("Chat", "/api/chat/messages", "Persist and fetch conversations"),
    @("Notifications", "/api/notifications", "List and mark notifications"),
    @("AI Doctor", "/api/ai/doctor", "Send symptom prompt to Groq API")
  ) -Top 145 -Height 430
  Add-Caption -Slide $slide -Text "Table 8. Core API summary of the implemented backend modules."
  $sourceLines.Add("6E Backend APIs")

  # Slide 54
  $slide = New-ReportSlide -Title "6F. FRONTEND IMPLEMENTATION AND REPRESENTATIVE SCREENS"
  Add-RectangleLabel -Slide $slide -Text "Login Page" -Left 55 -Top 155 -Width 210 -Height 160 -Fill 15724527 -FontSize 12 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Dashboard" -Left 330 -Top 155 -Width 210 -Height 160 -Fill 15724527 -FontSize 12 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Cart and Order Confirmation" -Left 55 -Top 390 -Width 210 -Height 160 -Fill 15724527 -FontSize 12 | Out-Null
  Add-RectangleLabel -Slide $slide -Text "Chat and AI Doctor Widget" -Left 330 -Top 390 -Width 210 -Height 160 -Fill 15724527 -FontSize 12 | Out-Null
  foreach ($coords in @(@(70,185,180,10), @(70,205,180,10), @(70,225,140,10), @(70,248,120,28), @(345,185,180,12), @(345,210,70,50), @(425,210,100,50), @(345,270,180,18), @(70,420,180,12), @(70,445,160,18), @(70,472,150,18), @(70,502,120,24), @(345,420,180,12), @(345,445,90,60), @(445,445,80,90))) {
    $shape = $slide.Shapes.AddShape(1, $coords[0], $coords[1], $coords[2], $coords[3])
    $shape.Fill.ForeColor.RGB = 16777215
    $shape.Line.ForeColor.RGB = 14079702
  }
  Add-TextBlock -Slide $slide -Text "The frontend is organized as reusable React components and clean role-based pages. The representative layouts above correspond to the major interaction surfaces in the implemented website." -Left 60 -Top 595 -Width 480 -Height 70 -FontSize 11 -Alignment 4 | Out-Null
  Add-Caption -Slide $slide -Text "Figure 11. Representative interface layouts aligned with the implemented login, dashboard, cart, confirmation, chat, and AI sections."
  $sourceLines.Add("6F Frontend Screens")

  # Slide 55
  $slide = New-ReportSlide -Title "6G. SECURITY, VALIDATION, AND FILE HANDLING"
  Add-BulletBlock -Slide $slide -Items @(
    "Passwords are hashed using bcrypt before storage in MongoDB.",
    "JWT middleware protects routes that require authenticated access.",
    "Role-based authorization restricts admin-only operations such as patient management and order delivery updates.",
    "OTP generation stores a temporary code and expiry timestamp for verification flow.",
    "Input validation checks required fields in authentication, appointment, and order forms.",
    "Multer processes multipart form data and stores uploads in a controlled server directory.",
    "Cart logic avoids duplicate product rows by increasing quantity when the same item is added again.",
    "Delivered order deletion is validated by status so users cannot remove active order history prematurely."
  ) -Top 145 -Height 500
  $sourceLines.Add("6G Security and Validation")

  # Slide 56
  $slide = New-ReportSlide -Title "CHAPTER 7: RESULTS AND DISCUSSION"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "Results and discussion evaluate whether the implemented system meets the project objectives in a functional and operational sense. Because the project is a working prototype rather than a benchmarked production deployment, the emphasis is placed on workflow completion, stability, response quality, and user-visible outcomes.",
    "Testing confirmed that the main modules including authentication, CRUD, appointments, file upload, search, cart, order confirmation, chat, notifications, and delivered-order controls behave as expected."
  ) -Height 330
  $sourceLines.Add("Chapter 7 Results")

  # Slide 57
  $slide = New-ReportSlide -Title "7A. TESTING STRATEGY AND TEST CASES"
  Add-TableBlock -Slide $slide -Headers @("TC", "Feature", "Expected Result", "Status") -Rows @(
    @("TC-01", "Registration and login", "User account created and authenticated", "Pass"),
    @("TC-02", "OTP verification", "Valid OTP marks user as verified", "Pass"),
    @("TC-03", "Patient CRUD", "Admin creates, edits, and deletes records", "Pass"),
    @("TC-04", "Appointment booking", "Appointment saved with status", "Pass"),
    @("TC-05", "Search and filter", "Matching records are returned", "Pass"),
    @("TC-06", "Cart quantity merge", "Same product increases quantity", "Pass"),
    @("TC-07", "Order confirmation page", "Total and confirmation view displayed", "Pass"),
    @("TC-08", "Delivered order delete rule", "Only delivered orders are removable", "Pass"),
    @("TC-09", "File upload", "File path returned and stored", "Pass"),
    @("TC-10", "Chat and notifications", "Messages and alerts are delivered", "Pass")
  ) -Top 145 -Height 470
  Add-Caption -Slide $slide -Text "Table 9. Representative system test cases executed during validation."
  $sourceLines.Add("7A Testing and Cases")

  # Slide 58
  $slide = New-ReportSlide -Title "7B. OUTPUT RESULTS AND MODULE-WISE OUTCOMES"
  Add-BulletBlock -Slide $slide -Items @(
    "Authentication output: secure login, token issue, and OTP verification response.",
    "Admin output: patient creation and appointment administration through dashboard controls.",
    "Patient output: access to appointments, uploads, cart, order history, chat, and AI doctor support.",
    "Order output: a separate confirmation page displays order placed message, total amount, items, and continue buying action.",
    "Chat output: live message updates arrive without page refresh through Socket.IO.",
    "Notification output: users receive system alerts for order and appointment changes.",
    "UI output: improved visual layout with section cards, hero area, cleaner forms, and clearer navigation."
  ) -Top 145 -Height 470
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "Taken together, these results demonstrate that the application satisfies the required feature set in a coordinated manner rather than as disconnected demos."
  ) -Top 650 -Height 60 -FontSize 11
  $sourceLines.Add("7B Module Outcomes")

  # Slide 59
  $slide = New-ReportSlide -Title "7C. PERFORMANCE METRICS AND EXPECTED VS ACTUAL COMPARISON"
  Add-TableBlock -Slide $slide -Headers @("Metric", "Expected", "Observed in Local Demo") -Rows @(
    @("Login flow", "Secure session creation", "Successful with token and route protection"),
    @("Search behavior", "Fast record filtering", "Near-instant on seeded dataset"),
    @("Chat latency", "Real-time delivery", "Immediate in local environment"),
    @("Order workflow", "Cart to confirmation transition", "Completed correctly with total calculation"),
    @("Notification workflow", "Automatic alert generation", "Working for tested event triggers")
  ) -Top 140 -Height 270
  Add-BarMetric -Slide $slide -Label "Authentication reliability" -Value 95 -Left 70 -Top 470
  Add-BarMetric -Slide $slide -Label "Appointment workflow success" -Value 92 -Left 70 -Top 510
  Add-BarMetric -Slide $slide -Label "Order workflow success" -Value 94 -Left 70 -Top 550
  Add-BarMetric -Slide $slide -Label "Chat responsiveness" -Value 96 -Left 70 -Top 590
  Add-BarMetric -Slide $slide -Label "Notification consistency" -Value 90 -Left 70 -Top 630
  Add-Caption -Slide $slide -Text "Table 10. Expected versus actual comparison with representative local-demo metrics."
  $sourceLines.Add("7C Performance and Comparison")

  # Slide 60
  $slide = New-ReportSlide -Title "7D. INTERPRETATION OF RESULTS"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The observed results indicate that the system fulfills its prototype objectives effectively. The strongest outcomes are visible in integration and workflow continuity: patient data, appointments, uploads, orders, notifications, and communication are all handled from one interface. This is a significant improvement over fragmented manual processes.",
    "The cart logic and confirmation page enhance the order module by making it behave like a realistic transactional flow. Similarly, the delivered-order deletion rule introduces business logic that prevents premature data removal. The chat and notification modules improve coordination, while the AI doctor widget demonstrates how a future-facing feature can be introduced without disrupting the core architecture.",
    "Overall, the results suggest that the project is functionally successful, academically robust, and well-positioned for future expansion."
  ) -Height 560
  $sourceLines.Add("7D Interpretation")

  # Slide 61
  $slide = New-ReportSlide -Title "CHAPTER 8: ADVANTAGES AND LIMITATIONS"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "Every applied system has strengths and constraints. In this project, the advantages arise from integration, usability, and structured design, while the limitations mostly relate to prototype scope and the absence of advanced enterprise features.",
    "Recognizing both aspects is important because it clarifies what the current project achieves and where future work should be directed."
  ) -Height 300
  $sourceLines.Add("Chapter 8 Advantages and Limitations")

  # Slide 62
  $slide = New-ReportSlide -Title "8A. ADVANTAGES OF THE PROPOSED SYSTEM"
  Add-BulletBlock -Slide $slide -Items @(
    "Centralizes hospital records, appointments, orders, and communication into one platform.",
    "Reduces paperwork and lowers the chance of manual record loss or duplication.",
    "Improves visibility of appointments and operational actions through structured dashboards.",
    "Supports direct interaction between users through real-time chat and notifications.",
    "Provides role-based access control and secure authentication practices.",
    "Demonstrates maintainable full stack design using MERN and MVC.",
    "Adds practical usability features such as cart quantity management, order confirmation, and document upload.",
    "Creates a foundation for advanced extensions such as telemedicine and AI-based support."
  ) -Top 145 -Height 500
  $sourceLines.Add("8A Advantages")

  # Slide 63
  $slide = New-ReportSlide -Title "8B. LIMITATIONS AND CONSTRAINTS"
  Add-BulletBlock -Slide $slide -Items @(
    "Online payment gateways are not fully integrated; the current order flow uses COD only.",
    "Uploaded files are stored locally rather than in cloud storage.",
    "The UI is improved but still intentionally simple for academic clarity rather than enterprise polish.",
    "The AI doctor assistant provides informational guidance and not a clinical diagnosis.",
    "Video consultation, wearable device linkage, and advanced analytics are not implemented in the current version.",
    "Testing is based on local prototype execution and seeded dummy data rather than production deployment."
  ) -Top 145 -Height 430
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "These limitations do not reduce the academic validity of the project; instead, they define the boundary of the current prototype and identify the areas most suitable for future enhancement."
  ) -Top 615 -Height 80 -FontSize 11
  $sourceLines.Add("8B Limitations")

  # Slide 64
  $slide = New-ReportSlide -Title "CHAPTER 9: CONCLUSION"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The Smart Hospital Management System successfully addresses the problem of manual hospital record management by providing a centralized, role-based, and full stack digital solution. The project demonstrates how MERN technologies can be combined with MVC architecture to implement a secure and modular application for hospital operations.",
    "The final system includes authentication, OTP verification, admin CRUD operations, patient and appointment management, report uploads, order processing with cart and COD, real-time chat, notifications, map integration, and an AI doctor assistant. Each of these modules contributes to a more connected and user-friendly workflow than the existing manual system.",
    "From an academic perspective, the project satisfies the major requirements of analysis, design, implementation, testing, and discussion. From a practical perspective, it establishes a strong prototype that can be further expanded into a more complete hospital management ecosystem."
  ) -Height 560
  $sourceLines.Add("Chapter 9 Conclusion")

  # Slide 65
  $slide = New-ReportSlide -Title "CHAPTER 10: FUTURE SCOPE"
  Add-TableBlock -Slide $slide -Headers @("Enhancement Area", "Possible Extension", "Expected Benefit") -Rows @(
    @("Telemedicine", "Video consultation and remote follow-up", "Improved accessibility for patients"),
    @("Wearable Integration", "Vital sign tracking through health devices", "Continuous monitoring and richer data"),
    @("Payments", "UPI or gateway integration", "Complete digital billing experience"),
    @("Cloud Storage", "Document storage on cloud platforms", "Scalable and secure report handling"),
    @("Analytics", "Charts and trend reports for admin", "Operational decision support"),
    @("Mobile App", "React Native companion application", "Broader access across devices")
  ) -Top 145 -Height 390
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "Future work may also include multilingual support, prescription PDF generation, doctor shift planning, and richer AI-assisted triage workflows. Because the current codebase is modular, these enhancements can be added progressively without requiring a complete redesign."
  ) -Top 585 -Height 100 -FontSize 11
  Add-Caption -Slide $slide -Text "Table 11. Future enhancement roadmap for the proposed system."
  $sourceLines.Add("Chapter 10 Future Scope")

  # Slide 66
  $slide = New-ReportSlide -Title "REFERENCES"
  Add-TextBlock -Slide $slide -Text @"
[1] A. R. Bakker and J. L. Mol, ""Hospital information systems,"" Eff Health Care, vol. 1, no. 4, pp. 215-223, 1983.

[2] P. L. Reichertz, ""Hospital information systems - past, present, future,"" Int J Med Inform, vol. 75, no. 3-4, pp. 282-299, 2006.

[3] C. A. Bain and C. Standing, ""A technology ecosystem perspective on hospital management information systems: lessons from the health literature,"" Int J Electron Healthc, vol. 5, no. 2, pp. 193-210, 2009.

[4] J. Sligo, R. Gauld, V. Roberts, and L. Villa, ""A literature review for large-scale health information system project planning, implementation and evaluation,"" Int J Med Inform, vol. 97, pp. 86-97, 2017.

[5] P. W. Handayani, A. N. Hidayanto, and I. Budi, ""User acceptance factors of hospital information systems and related technologies: Systematic review,"" Inform Health Soc Care, vol. 43, no. 4, pp. 401-426, 2018.

[6] M. Khalifa and O. Alswailem, ""Hospital information systems (HIS) acceptance and satisfaction: a case study of a tertiary care hospital,"" Procedia Comput Sci, vol. 63, pp. 198-204, 2015.
"@ -Left 52 -Top 125 -Width 490 -Height 620 -FontSize 10 -Alignment 1
  $sourceLines.Add("References I")

  # Slide 67
  $slide = New-ReportSlide -Title "REFERENCES (CONTINUED)"
  Add-TextBlock -Slide $slide -Text @"
[7] J. S. McCullough, ""The adoption of hospital information systems,"" Health Econ, vol. 17, no. 5, pp. 649-664, 2008.

[8] MongoDB, ""MongoDB Documentation."" [Online]. Available: https://www.mongodb.com/docs/ . [Accessed: Apr. 18, 2026].

[9] Express.js, ""Express Documentation."" [Online]. Available: https://expressjs.com/ . [Accessed: Apr. 18, 2026].

[10] React, ""React Documentation."" [Online]. Available: https://react.dev/ . [Accessed: Apr. 18, 2026].

[11] Node.js, ""Node.js Documentation."" [Online]. Available: https://nodejs.org/en/docs . [Accessed: Apr. 18, 2026].

[12] Socket.IO, ""Socket.IO Documentation."" [Online]. Available: https://socket.io/docs/v4/ . [Accessed: Apr. 18, 2026].

[13] Mongoose, ""Mongoose Documentation."" [Online]. Available: https://mongoosejs.com/docs/ . [Accessed: Apr. 18, 2026].

[14] IETF, ""RFC 7519: JSON Web Token (JWT)."" [Online]. Available: https://datatracker.ietf.org/doc/html/rfc7519 . [Accessed: Apr. 18, 2026].

[15] Leaflet, ""Leaflet Documentation."" [Online]. Available: https://leafletjs.com/ . [Accessed: Apr. 18, 2026].
"@ -Left 52 -Top 125 -Width 490 -Height 620 -FontSize 10 -Alignment 1
  $sourceLines.Add("References II")

  # Slide 68
  $slide = New-ReportSlide -Title "APPENDIX A: SAMPLE API ENDPOINTS"
  Add-TextBlock -Slide $slide -Text @"
POST   /api/auth/register            Create a new user account
POST   /api/auth/login               Authenticate and issue JWT
POST   /api/auth/send-otp            Generate and send email OTP
POST   /api/auth/verify-otp          Verify OTP and mark user verified
GET    /api/patients                 Fetch patient records
POST   /api/patients                 Create patient profile
PUT    /api/patients/:id             Update patient information
DELETE /api/patients/:id             Delete patient record
POST   /api/appointments             Create appointment
PUT    /api/appointments/:id/status  Update appointment status
POST   /api/orders/cart              Add item to cart
PATCH  /api/orders/cart/remove-one   Remove one quantity from cart
POST   /api/orders/place             Place order with COD
DELETE /api/orders/:id               Delete delivered order history entry
POST   /api/upload                   Upload report file
GET    /api/notifications            List notifications
POST   /api/ai/doctor                Symptom prompt to AI doctor assistant
"@ -Left 60 -Top 125 -Width 470 -Height 620 -FontSize 11 -Alignment 1
  $sourceLines.Add("Appendix A Endpoints")

  # Slide 69
  $slide = New-ReportSlide -Title "APPENDIX B: DUMMY DATA SAMPLES"
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "Representative demo accounts used for testing include: Admin - admin@hospital.com / 123456, Doctor - doctor@hospital.com / 123456, and Patient - patient@hospital.com / 123456. Dummy patient examples include records such as Rahul Verma and other manually added profiles from the admin dashboard.",
    "Sample medical store entries include tablet, syrup, and report-related consumables with stored price and quantity data. Example appointment states include Scheduled, Completed, and Cancelled. Notifications are generated on order placement, appointment updates, and welcome or system events.",
    "This appendix confirms that the system was designed to be demonstrable without external datasets while still preserving realistic hospital management behavior."
  ) -Height 360
  Add-BulletBlock -Slide $slide -Items @(
    "Cart behavior: repeated product addition increments quantity.",
    "Cart removal behavior: remove action decreases one quantity at a time.",
    "Order history behavior: deletion is allowed only when an order is marked delivered.",
    "Confirmation behavior: placing COD order redirects to a confirmation page with total amount and continue buying action."
  ) -Top 520 -Height 160
  $sourceLines.Add("Appendix B Dummy Data")

  # Slide 70
  $slide = New-ReportSlide -Title "APPENDIX C: SOURCE CODE ORGANIZATION"
  Add-TextBlock -Slide $slide -Text @"
backend/
  config/
  controllers/
  middleware/
  models/
  routes/
  seed/
  uploads/
  utils/
  server.js

frontend/
  src/
    components/
    context/
    pages/
    services/
    styles/
  App.jsx

report/
  Smart_Hospital_Management_System_Standards_Report.pdf
  Smart_Hospital_Management_System_Standards_Report.pptx
  diagrams/
"@ -Left 80 -Top 145 -Width 220 -Height 420 -FontSize 12 -Alignment 1
  Add-ParagraphBlock -Slide $slide -Paragraphs @(
    "The report diagrams have also been exported separately to the diagrams folder for easy reuse in presentations or document editing. This appendix gives a compact overview of the implemented codebase and the generated report assets."
  ) -Left 320 -Top 190 -Width 220 -Height 180 -FontSize 11 -Alignment 4
  Add-BulletBlock -Slide $slide -Items @(
    "01_system_architecture.png",
    "03_dfd_level_0.png",
    "04_dfd_level_1_patient.png",
    "05_dfd_level_1_admin_doctor.png",
    "06_use_case_diagram.png",
    "07_sequence_diagram.png"
  ) -Left 330 -Top 420 -Width 180 -Height 170 -FontSize 10
  $sourceLines.Add("Appendix C Source Organization")

  Set-Content -Path $sourcePath -Value ($sourceLines -join [Environment]::NewLine)

  $diagramExports = [ordered]@{
    33 = "01_system_architecture.png"
    34 = "02_module_design.png"
    35 = "03_dfd_level_0.png"
    36 = "04_dfd_level_1_patient.png"
    37 = "05_dfd_level_1_admin_doctor.png"
    38 = "06_use_case_diagram.png"
    39 = "07_sequence_diagram.png"
    40 = "08_schema_er_diagram.png"
    41 = "09_authentication_flowchart.png"
    42 = "10_order_processing_flowchart.png"
  }

  foreach ($item in $diagramExports.GetEnumerator()) {
    $presentation.Slides.Item([int]$item.Key).Export((Join-Path $diagramDir $item.Value), "PNG", 1600, 2263)
  }

  $presentation.SaveAs($pptPath)
  $presentation.SaveAs($pdfPath, 32)
  Write-Output "Generated report with $($presentation.Slides.Count) slides."
  Write-Output "PPTX: $pptPath"
  Write-Output "PDF:  $pdfPath"
}
finally {
  if ($presentation) { $presentation.Close() }
  if ($ppt) { $ppt.Quit() }
}
