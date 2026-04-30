$ErrorActionPreference = "Stop"

$projectRoot = "C:\Users\shrey\OneDrive\Documents\New project"
$pptPath = Join-Path $projectRoot "Smart_Hospital_Presentation_Fixed.pptx"

$slides = @(
  @{
    Title = "Smart Hospital Management System"
    Lines = @(
      "A MERN Stack Based Full Stack Web Application",
      "Presented By: Your Name",
      "Roll Number: Your Roll Number",
      "Department: Your Department",
      "Guide: Guide Name",
      "Academic Year: 2025-2026"
    )
  },
  @{
    Title = "Introduction"
    Lines = @(
      "Hospitals handle large amounts of patient and administrative data every day.",
      "Manual record management is slow, repetitive, and difficult to maintain.",
      "Digital systems improve speed, accuracy, communication, and accessibility.",
      "This project provides a centralized platform for admin, doctors, and patients."
    )
  },
  @{
    Title = "Problem Statement"
    Lines = @(
      "Manual hospital record management causes delays and operational errors.",
      "Patient records are difficult to maintain, search, and update quickly.",
      "Appointment booking and tracking are inefficient in manual workflows.",
      "Reports, billing, and communication often remain disconnected.",
      "A smart digital solution is needed to streamline hospital operations."
    )
  },
  @{
    Title = "Proposed Solution"
    Lines = @(
      "Develop a full stack Smart Hospital Management System using MERN.",
      "Provide centralized patient and appointment management.",
      "Support secure login with email OTP verification.",
      "Enable admin CRUD operations and hospital workflow monitoring.",
      "Add real-time communication, notifications, file upload, and map integration."
    )
  },
  @{
    Title = "Objectives"
    Lines = @(
      "Digitize hospital activities and reduce paperwork.",
      "Improve security through authentication and OTP verification.",
      "Manage patients, appointments, and orders efficiently.",
      "Improve communication through chat and notifications.",
      "Build a maintainable full stack application using MVC."
    )
  },
  @{
    Title = "Technology Stack"
    Lines = @(
      "Frontend: React.js, JSX, CSS, React Router",
      "Backend: Node.js and Express.js",
      "Database: MongoDB with Mongoose",
      "Authentication: JWT, bcrypt, Email OTP",
      "Real-Time Chat: Socket.IO",
      "Uploads and Map: Multer, React Leaflet, OpenStreetMap"
    )
  },
  @{
    Title = "Why MERN Stack"
    Lines = @(
      "Uses JavaScript across the full application.",
      "Supports fast development and clean integration between layers.",
      "React provides reusable components and easy UI updates.",
      "MongoDB stores flexible healthcare-related data structures.",
      "Express and Node create lightweight and scalable APIs."
    )
  },
  @{
    Title = "System Architecture"
    Lines = @(
      "Presentation Layer: React frontend for user interaction.",
      "Application Layer: Express routes, controllers, and middleware.",
      "Data Layer: MongoDB collections and Mongoose models.",
      "Real-Time Layer: Socket.IO for instant user-to-user chat.",
      "Overall flow: User to Frontend to Backend API to Database."
    )
  },
  @{
    Title = "MVC Pattern"
    Lines = @(
      "Model: Defines the structure of application data.",
      "View: React pages and components that display data.",
      "Controller: Business logic and request handling.",
      "MVC improves maintainability, structure, and scalability.",
      "It keeps data, logic, and UI concerns clearly separated."
    )
  },
  @{
    Title = "Main Features"
    Lines = @(
      "User Authentication and Email OTP Verification",
      "Admin CRUD Operations",
      "Search and Filter System",
      "Appointment Management and E-Prescription Support",
      "Order Processing with Add to Cart and COD",
      "File Upload, Map Integration, Chat, and Notifications"
    )
  },
  @{
    Title = "User Roles"
    Lines = @(
      "Admin: Manage users, patients, appointments, and orders.",
      "Doctor: View assigned appointments and communicate with users.",
      "Patient: Register, log in, upload reports, book appointments, and place orders.",
      "Role-based access ensures secure and focused workflows."
    )
  },
  @{
    Title = "Modules of the System"
    Lines = @(
      "Authentication Module",
      "Admin Management Module",
      "Patient Management Module",
      "Appointment Module",
      "Order Processing Module",
      "File Upload, Chat, Notification, and Map Modules"
    )
  },
  @{
    Title = "Database Design"
    Lines = @(
      "Main collections: Users, Patients, Appointments, Orders, Messages, Notifications.",
      "Users store identity, role, OTP, and cart details.",
      "Patients store health profile information.",
      "Appointments link doctors and patients with schedule details.",
      "Orders, messages, and notifications support workflow and communication."
    )
  },
  @{
    Title = "Order Processing System"
    Lines = @(
      "Users can browse medical store items.",
      "Selected items are added to cart and stored in the user record.",
      "Orders are placed using Cash on Delivery.",
      "Admins can track and update order status.",
      "Notifications are generated when orders are placed or updated."
    )
  },
  @{
    Title = "Real-Time Chat and Notifications"
    Lines = @(
      "Socket.IO enables live user-to-user communication.",
      "Messages are stored in the database for conversation history.",
      "Notifications inform users about important actions.",
      "Examples include welcome messages, appointment updates, and order updates."
    )
  },
  @{
    Title = "Additional Features"
    Lines = @(
      "Search and filter improve record accessibility.",
      "File upload allows digital report submission.",
      "Geo-location map shows the hospital location.",
      "Role-protected routes improve security.",
      "The interface is simple, structured, and easy to use."
    )
  },
  @{
    Title = "Testing and Results"
    Lines = @(
      "Tested registration, login, OTP verification, and protected routes.",
      "Verified patient CRUD, appointment booking, and search workflows.",
      "Checked cart, order placement, file upload, and notification updates.",
      "Real-time chat worked with message storage and live delivery.",
      "The system successfully demonstrated the required project features."
    )
  },
  @{
    Title = "Advantages"
    Lines = @(
      "Reduces paperwork and manual record handling.",
      "Improves accessibility of patient and appointment data.",
      "Centralizes multiple hospital workflows in one system.",
      "Enhances communication between users.",
      "Provides a strong base for future healthcare features."
    )
  },
  @{
    Title = "Limitations and Future Scope"
    Lines = @(
      "Current payment support is limited to COD.",
      "File storage is local and not cloud-based yet.",
      "Telemedicine and wearable integration are future features.",
      "Possible enhancements include online payments, PDF generation, and mobile support."
    )
  },
  @{
    Title = "Conclusion"
    Lines = @(
      "The Smart Hospital Management System replaces manual processes with a digital platform.",
      "It demonstrates secure, modular, and scalable MERN-based development.",
      "The project successfully combines hospital management, communication, and record handling.",
      "It forms a strong foundation for future healthcare system expansion."
    )
  },
  @{
    Title = "Thank You"
    Lines = @(
      "Thank You",
      "Any Questions?"
    )
  }
)

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$presentation = $ppt.Presentations.Add()

try {
  foreach ($index in 0..($slides.Count - 1)) {
    $slideData = $slides[$index]
    $layout = if ($index -eq 0 -or $slideData.Title -eq "Thank You") { 2 } else { 2 }
    $slide = $presentation.Slides.Add($index + 1, $layout)

    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.ForeColor.RGB = 16777215

    $titleShape = $slide.Shapes.Title
    $titleShape.TextFrame.TextRange.Text = $slideData.Title
    $titleShape.TextFrame.TextRange.Font.Name = "Aptos Display"
    $titleShape.TextFrame.TextRange.Font.Size = 26
    $titleShape.TextFrame.TextRange.Font.Bold = -1
    $titleShape.TextFrame.TextRange.Font.Color.RGB = 4869652

    $contentShape = $slide.Shapes.Placeholders(2)
    $contentShape.Left = 60
    $contentShape.Top = 120
    $contentShape.Width = 600
    $contentShape.Height = 330
    $contentShape.Fill.ForeColor.RGB = 16316669
    $contentShape.Line.ForeColor.RGB = 15131881

    $textRange = $contentShape.TextFrame.TextRange
    $textRange.Text = ""

    if ($index -eq 0 -or $slideData.Title -eq "Thank You") {
      $textRange.Text = ($slideData.Lines -join "`r")
      $textRange.Font.Name = "Aptos"
      $textRange.Font.Size = 18
      $textRange.Font.Color.RGB = 8215373
      $textRange.ParagraphFormat.Bullet.Visible = 0
      $textRange.ParagraphFormat.Alignment = 2
    } else {
      $textRange.Text = ($slideData.Lines -join "`r")
      $textRange.Font.Name = "Aptos"
      $textRange.Font.Size = 20
      $textRange.Font.Color.RGB = 6044187
      $textRange.ParagraphFormat.Bullet.Visible = -1
      $textRange.ParagraphFormat.Bullet.Character = 8226
      $textRange.ParagraphFormat.SpaceAfter = 8
    }

    $accent = $slide.Shapes.AddShape(1, 40, 90, 100, 4)
    $accent.Fill.ForeColor.RGB = 9926442
    $accent.Line.Visible = 0
  }

  if (Test-Path $pptPath) {
    Remove-Item -LiteralPath $pptPath -Force
  }

  $presentation.SaveAs($pptPath)
}
finally {
  $presentation.Close()
  $ppt.Quit()
}
