$ErrorActionPreference = "Stop"

$projectRoot = "C:\Users\shrey\OneDrive\Documents\New project"
$docxPath = Join-Path $projectRoot "Smart_Hospital_Report.docx"
$pdfPath = Join-Path $projectRoot "Smart_Hospital_Report.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$document = $word.Documents.Add()
$selection = $word.Selection

function Add-Paragraph {
  param(
    [string]$Text,
    [int]$Style = 0,
    [switch]$Center,
    [switch]$Bold,
    [int]$SpaceAfter = 8
  )

  $selection.Style = $Style
  $selection.ParagraphFormat.Alignment = if ($Center) { 1 } else { 0 }
  $selection.Font.Bold = if ($Bold) { 1 } else { 0 }
  $selection.TypeText($Text)
  $selection.TypeParagraph()
  $selection.ParagraphFormat.SpaceAfter = $SpaceAfter
  $selection.Font.Bold = 0
}

function Add-Bullets {
  param([string[]]$Items)

  $selection.Range.ListFormat.ApplyBulletDefault() | Out-Null
  foreach ($item in $Items) {
    $selection.TypeText($item)
    $selection.TypeParagraph()
  }
  $selection.Range.ListFormat.RemoveNumbers()
}

function Add-PageBreak {
  $selection.InsertBreak(7)
}

function Add-Table {
  param(
    [string[]]$Headers,
    [object[][]]$Rows
  )

  $range = $selection.Range
  $table = $document.Tables.Add($range, $Rows.Count + 1, $Headers.Count)
  for ($i = 0; $i -lt $Headers.Count; $i++) {
    $table.Cell(1, $i + 1).Range.Text = $Headers[$i]
    $table.Cell(1, $i + 1).Range.Bold = 1
  }
  for ($r = 0; $r -lt $Rows.Count; $r++) {
    for ($c = 0; $c -lt $Headers.Count; $c++) {
      $table.Cell($r + 2, $c + 1).Range.Text = [string]$Rows[$r][$c]
    }
  }
  $selection.MoveDown() | Out-Null
  $selection.TypeParagraph()
}

New-Item -ItemType Directory -Force -Path (Join-Path $projectRoot "scripts") | Out-Null

Add-Paragraph -Text "SMART HOSPITAL MANAGEMENT SYSTEM" -Style -1 -Center -Bold -SpaceAfter 16
Add-Paragraph -Text "Detailed Project Report" -Center -Bold
Add-Paragraph -Text "Submitted in partial fulfillment of the requirements for the award of the degree/course" -Center
Add-Paragraph -Text "Submitted By: Your Name" -Center
Add-Paragraph -Text "Roll Number: Your Roll Number" -Center
Add-Paragraph -Text "Department: Your Department" -Center
Add-Paragraph -Text "College: Your College Name" -Center
Add-Paragraph -Text "Project Guide: Guide Name" -Center
Add-Paragraph -Text "Academic Year: 2025-2026" -Center
Add-PageBreak

Add-Paragraph -Text "Certificate" -Style -2 -Center -Bold
Add-Paragraph -Text "This is to certify that the project report entitled Smart Hospital Management System submitted by Your Name is a bona fide work carried out under my guidance and supervision in partial fulfillment of the requirements for the award of the degree/course during the academic year 2025-2026."
Add-Paragraph -Text "To the best of my knowledge, the work presented in this report is original and has not been submitted elsewhere for the award of any degree, diploma, or certificate."
Add-Paragraph -Text "Project Guide Signature ____________________"
Add-Paragraph -Text "Head of Department Signature ____________________"
Add-PageBreak

Add-Paragraph -Text "Declaration" -Style -2 -Center -Bold
Add-Paragraph -Text "I hereby declare that the project report entitled Smart Hospital Management System is an original piece of work carried out by me under the guidance of my project mentor. The report has been prepared for academic purposes and has not been submitted earlier, either in full or in part, to any institution or university for the award of any degree or diploma."
Add-Paragraph -Text "I further declare that all sources of information used in this report have been duly acknowledged."
Add-Paragraph -Text "Student Signature ____________________"
Add-PageBreak

Add-Paragraph -Text "Acknowledgement" -Style -2 -Center -Bold
Add-Paragraph -Text "I would like to express my sincere gratitude to my project guide, faculty members, and department for their continuous support, guidance, and encouragement throughout the development of this project. Their insights helped me understand both the technical and practical aspects of designing a real-world hospital management solution."
Add-Paragraph -Text "I am also thankful to my classmates and friends for their suggestions during the planning, development, and testing phases. Their feedback allowed me to improve the structure and usability of the project. Finally, I would like to thank my family for their constant support and motivation during the completion of this work."
Add-PageBreak

Add-Paragraph -Text "Abstract" -Style -2 -Center -Bold
Add-Paragraph -Text "The Smart Hospital Management System is a web-based full stack application designed to digitize and simplify hospital record management. Many small and medium hospitals still depend on manual files, registers, and disconnected systems to manage patients, appointments, prescriptions, reports, billing, and communication. Such manual processes often result in delays, record duplication, difficulty in retrieval, and reduced service quality."
Add-Paragraph -Text "The proposed system addresses these problems through a centralized platform developed using the MERN stack, namely MongoDB, Express.js, React.js, and Node.js. The backend follows the MVC pattern to maintain clear separation between data models, business logic, and routing layers. The system supports three primary roles: administrator, doctor, and patient. It includes modules for secure authentication, email OTP verification, patient and appointment management, admin CRUD operations, file upload, order processing with cart and Cash on Delivery support, search and filtering, geo-location map integration, real-time chat, and notifications."
Add-Paragraph -Text "The objective of the project is to reduce paperwork, improve data accessibility, enhance communication, and create a structured application that can be understood, maintained, and extended easily. The current implementation forms a strong foundation for future modules such as telemedicine, wearable integration, cloud document management, and digital payment systems."
Add-PageBreak

Add-Paragraph -Text "Table of Contents" -Style -2 -Center -Bold
Add-Bullets -Items @(
  "1. Introduction",
  "2. Problem Statement",
  "3. Need for the System",
  "4. Existing System",
  "5. Proposed System",
  "6. Objectives",
  "7. Scope of the Project",
  "8. Literature and Background Study",
  "9. Requirement Analysis",
  "10. Technology Stack",
  "11. System Architecture",
  "12. MVC Design Pattern",
  "13. Modules of the System",
  "14. Database Design",
  "15. Implementation Details",
  "16. Algorithms and Process Flow",
  "17. Testing and Validation",
  "18. Results and Discussion",
  "19. Advantages",
  "20. Limitations",
  "21. Future Scope",
  "22. Conclusion",
  "23. References"
)
Add-PageBreak

Add-Paragraph -Text "Chapter 1: Introduction" -Style -2 -Bold
Add-Paragraph -Text "Healthcare institutions manage sensitive and high-volume information every day. From the moment a patient registers at the reception desk to the moment treatment is completed, the hospital generates and consumes multiple forms of data including demographic information, appointment schedules, prescriptions, laboratory reports, billing details, doctor notes, and communication records. In a manual environment, these pieces of information are often scattered across paper files or isolated systems, which leads to inefficiency."
Add-Paragraph -Text "Hospital administration is not limited to medical treatment alone. It includes workflow coordination, scheduling, record safety, quick retrieval of data, communication between stakeholders, and operational visibility for the management team. When these activities are not digitized, staff spend extra time searching records, manually verifying appointments, maintaining files, and coordinating with patients."
Add-Paragraph -Text "The Smart Hospital Management System has been designed as a practical web-based solution to these operational challenges. It provides a centralized application where an administrator can manage records, a doctor can view assigned appointments, and a patient can access personal services such as appointments, reports, notifications, and orders. The project emphasizes clarity, simplicity, and maintainability, making it suitable both as an academic submission and as a prototype for future development."

Add-Paragraph -Text "Chapter 2: Problem Statement" -Style -2 -Bold
Add-Paragraph -Text "The problem considered in this project is the continued dependence on manual hospital record management. Manual systems are slow, repetitive, and highly dependent on staff availability. Searching for patient history, managing doctor schedules, handling reports, or tracking appointment status becomes difficult when records are not maintained in a centralized digital environment."
Add-Paragraph -Text "In addition, communication between hospital users is often fragmented. Patients may need to visit physically or rely on calls for updates. Notification systems may be absent. Reports and prescriptions can be misplaced. Billing and order-related activities may operate separately from the rest of the workflow. These issues directly affect efficiency, accuracy, and the overall patient experience."

Add-Paragraph -Text "Chapter 3: Need for the System" -Style -2 -Bold
Add-Paragraph -Text "A hospital management system is needed to introduce consistency, speed, and traceability into healthcare operations. Digitization reduces dependency on paper-based records and improves the ability of hospital staff to access the right information at the right time. By organizing patient data, appointments, orders, and notifications inside one platform, the hospital can operate with better control and less duplication."
Add-Paragraph -Text "The proposed project specifically addresses the need for a beginner-friendly but functional digital solution. It is not designed as an overly complex enterprise platform; instead, it is structured to demonstrate the most important components of a smart hospital workflow in a way that is easy to implement, explain, and improve further."
Add-PageBreak

Add-Paragraph -Text "Chapter 4: Existing System" -Style -2 -Bold
Add-Paragraph -Text "In the existing manual setup, most activities are performed through physical files, registers, or basic spreadsheets. Patient registration is done manually at the reception. Doctor appointments are noted in diaries or separate sheets. Test reports may be stored physically or in unstructured folders. There is often no direct connection between appointment data, prescription details, and follow-up communication."
Add-Paragraph -Text "This causes several operational issues. Record duplication becomes common because old data cannot be found quickly. Appointments may overlap or be delayed. The medical store or billing counter may not have access to complete treatment context. Communication becomes reactive instead of organized, and the hospital lacks a reliable dashboard for monitoring daily operations."
Add-Paragraph -Text "Limitations of the Existing System:" -Bold
Add-Bullets -Items @(
  "Heavy dependence on paper-based records",
  "Slow retrieval of patient and appointment history",
  "Greater risk of human error and data inconsistency",
  "No real-time communication between users",
  "Poor integration between appointments, reports, and orders",
  "Difficulty in scaling processes when patient volume increases"
)

Add-Paragraph -Text "Chapter 5: Proposed System" -Style -2 -Bold
Add-Paragraph -Text "The proposed Smart Hospital Management System centralizes the major workflows of a hospital through a browser-based application. The system allows users to log in securely, access features based on their role, and interact with hospital services digitally. It includes an administrator panel for CRUD operations, patient and appointment management, medical order handling, file upload, notifications, and a chat module."
Add-Paragraph -Text "The proposed system emphasizes structured development. The backend is organized using the MVC pattern with clearly separated models, controllers, routes, and middleware. The frontend is built with reusable React components and pages using JSX. The result is a codebase that is simpler to understand and maintain."
Add-Paragraph -Text "Key Characteristics of the Proposed System:" -Bold
Add-Bullets -Items @(
  "Centralized data access",
  "Role-based user interaction",
  "Secure authentication and authorization",
  "Digital workflow for appointments and patient records",
  "Integrated notifications and communication",
  "Extensibility for future healthcare features"
)
Add-PageBreak

Add-Paragraph -Text "Chapter 6: Objectives" -Style -2 -Bold
Add-Bullets -Items @(
  "To replace manual hospital administration with a digital system",
  "To reduce paperwork and repetitive tasks",
  "To create a secure authentication process using JWT and email OTP",
  "To manage patient records through admin-controlled CRUD operations",
  "To simplify appointment creation, tracking, and status updates",
  "To provide a medical order module with add-to-cart and COD checkout",
  "To support document uploads for patient reports",
  "To improve communication using real-time chat and notifications",
  "To create a maintainable MERN application following MVC principles"
)

Add-Paragraph -Text "Chapter 7: Scope of the Project" -Style -2 -Bold
Add-Paragraph -Text "The scope of the project includes the design and implementation of a hospital management prototype for digital record handling. It focuses on operations commonly needed in a clinic or hospital environment: authentication, patient data, appointments, chat, file upload, notifications, map display, and order management. The system is suitable for demonstration, academic evaluation, and as a base architecture for larger healthcare applications."
Add-Paragraph -Text "The current scope intentionally avoids highly specialized medical modules such as insurance claim processing, advanced clinical decision support, laboratory device integration, telemedicine video streaming, and complex third-party billing infrastructure. These are reserved for future enhancement."

Add-Paragraph -Text "Chapter 8: Literature and Background Study" -Style -2 -Bold
Add-Paragraph -Text "Modern healthcare systems increasingly depend on digital platforms for operational efficiency. Hospital information systems combine patient registration, diagnosis records, appointment scheduling, billing, and communication into one integrated environment. The move toward digitalization is driven by the need for better service quality, reduced paperwork, stronger security, and easier reporting."
Add-Paragraph -Text "Web-based systems offer multiple advantages in this context. They can be used through standard browsers, reduce installation overhead, and can support multiple departments or user roles. Real-time communication technologies, document upload facilities, and centralized databases make such platforms more practical than isolated desktop tools."
Add-Paragraph -Text "The MERN stack is well suited to this use case because it uses JavaScript across the entire application. React helps build responsive user interfaces, Express and Node create a lightweight backend service layer, and MongoDB offers flexible document storage for diverse entities such as appointments, messages, and notifications. The MVC pattern supports long-term maintainability by separating responsibilities."
Add-PageBreak

Add-Paragraph -Text "Chapter 9: Requirement Analysis" -Style -2 -Bold
Add-Paragraph -Text "Functional Requirements:" -Bold
Add-Bullets -Items @(
  "User registration, login, logout, and profile access",
  "Email OTP generation and verification",
  "Role-based access for admin, doctor, and patient",
  "CRUD operations for patient management in admin panel",
  "Appointment creation, view, update, and delete",
  "Search and filter functionality for quick record lookup",
  "Medical store order processing with cart and COD support",
  "File upload for medical reports or related documents",
  "Geo-location map display for hospital location",
  "Real-time user-to-user chat",
  "Notification generation and read status management"
)
Add-Paragraph -Text "Non-Functional Requirements:" -Bold
Add-Bullets -Items @(
  "The system should be easy to understand and use",
  "The system should provide secure route protection",
  "The codebase should be modular and maintainable",
  "The frontend should work on common desktop and laptop screens",
  "The application should support future scalability"
)
Add-Paragraph -Text "Hardware and Software Requirements:" -Bold
Add-Bullets -Items @(
  "Computer or laptop with at least 4 GB RAM",
  "Modern web browser",
  "Node.js and npm",
  "MongoDB",
  "Visual Studio Code or a similar editor"
)

Add-Paragraph -Text "Chapter 10: Technology Stack" -Style -2 -Bold
Add-Table -Headers @("Layer", "Technology", "Purpose") -Rows @(
  @("Frontend", "React.js with JSX, CSS, React Router", "Builds the user interface and handles navigation"),
  @("Backend", "Node.js, Express.js", "Creates the REST API and business logic"),
  @("Database", "MongoDB, Mongoose", "Stores application data in collections and schemas"),
  @("Authentication", "JWT, bcrypt, email OTP", "Provides secure login and user verification"),
  @("Real-Time", "Socket.IO", "Enables live chat between users"),
  @("Uploads", "Multer", "Processes multipart form data and file uploads"),
  @("Map", "React Leaflet, OpenStreetMap", "Displays the hospital location")
)
Add-Paragraph -Text "The choice of the MERN stack was based on consistency and development speed. Since JavaScript is used on both the frontend and backend, the learning curve remains manageable and the project structure becomes more uniform. This is especially beneficial for educational projects that aim to demonstrate end-to-end full stack capabilities without introducing too many different languages or frameworks."
Add-PageBreak

Add-Paragraph -Text "Chapter 11: System Architecture" -Style -2 -Bold
Add-Paragraph -Text "The architecture of the Smart Hospital Management System can be viewed as a layered structure. The presentation layer consists of the React frontend that collects user input and displays data. The application layer consists of Express routes, controllers, middleware, and helper functions. The data layer consists of MongoDB collections and Mongoose models. Real-time communication is handled through Socket.IO, which works alongside the normal HTTP request-response cycle."
Add-Paragraph -Text "This layered separation improves maintainability and keeps the system organized. The frontend remains focused on interaction and display. The backend enforces rules, validation, and access control. The database stores persistent records. Real-time features are handled separately without mixing socket logic into every request."
Add-Paragraph -Text "[Architecture Diagram Placeholder: User -> React Frontend -> Express API -> MongoDB, plus Socket.IO for chat]" -Center

Add-Paragraph -Text "Chapter 12: MVC Design Pattern" -Style -2 -Bold
Add-Paragraph -Text "Model:" -Bold
Add-Paragraph -Text "Models define the structure of the application data. Each important entity in the system has its own schema: User, Patient, Appointment, Order, Message, and Notification. Models ensure that related fields are grouped together and stored consistently."
Add-Paragraph -Text "View:" -Bold
Add-Paragraph -Text "In this project, the view layer is represented by the React frontend. It contains pages such as login, registration, and dashboard, along with reusable components such as Navbar, StatCard, ChatBox, and MapView."
Add-Paragraph -Text "Controller:" -Bold
Add-Paragraph -Text "Controllers implement business logic. They receive input from routes, query or update the models, and send appropriate responses back to the frontend. For example, the authentication controller manages login, register, OTP sending, and verification. Other controllers manage appointments, patients, orders, notifications, and chat."
Add-Paragraph -Text "Using MVC reduces code duplication and makes the application easier to debug. It also helps in explaining the project because each file group has a clear responsibility."

Add-Paragraph -Text "Chapter 13: Modules of the System" -Style -2 -Bold
Add-Paragraph -Text "Authentication Module:" -Bold
Add-Paragraph -Text "The authentication module provides secure account creation and login functionality. Users can register with role information and then log in using email and password. Passwords are hashed using bcrypt before storage. JWT tokens are generated after successful authentication so the client can access protected API routes."
Add-Paragraph -Text "Email OTP is included as an additional verification mechanism. When a user requests an OTP, the backend generates a six-digit code, stores it temporarily in the database along with an expiry time, and attempts to send it through email. In demo mode, the OTP can still be viewed and tested even without SMTP credentials."
Add-Paragraph -Text "Admin Module:" -Bold
Add-Paragraph -Text "The admin module acts as the control center of the application. The administrator can view the dashboard summary, inspect users, create patient records, book appointments, update status values, delete selected records, and track medical orders. This module demonstrates how CRUD operations can be implemented in a practical and role-protected manner."
Add-Paragraph -Text "Patient Management Module:" -Bold
Add-Paragraph -Text "The patient module stores profile information such as age, gender, blood group, disease details, reports, and prescription references. This module links medical profile data to the base user account. It allows the hospital to manage patient-specific information separately from authentication fields."
Add-Paragraph -Text "Appointment Module:" -Bold
Add-Paragraph -Text "The appointment module is used to create, update, and track visits between patients and doctors. Each appointment contains the selected patient, doctor, date, time, reason, status, and e-prescription field. Appointment states such as scheduled, completed, or cancelled help represent the treatment workflow clearly."
Add-Paragraph -Text "Search and Filter Module:" -Bold
Add-Paragraph -Text "As the number of records grows, searching becomes essential. The search and filter module helps users quickly locate data by names, diseases, order items, and appointment status. This improves usability and reduces the time required to retrieve important information."
Add-Paragraph -Text "Order Processing Module:" -Bold
Add-Paragraph -Text "This module simulates a hospital medical store experience. Users can add items to a cart, review selected products, and place an order using Cash on Delivery. The order record stores the purchased items, amount, payment status, and order status. The administrator can later update delivery-related states."
Add-Paragraph -Text "File Upload Module:" -Bold
Add-Paragraph -Text "File upload is useful for medical reports and supporting documents. The system uses Multer to handle multipart form uploads and stores files locally on the server. The returned file path can be associated with patient records or displayed to the user for confirmation."
Add-Paragraph -Text "Geo-location Map Module:" -Bold
Add-Paragraph -Text "The map module uses React Leaflet and OpenStreetMap to display the hospital location. Although simple, this feature enhances the user experience and demonstrates the integration of external mapping capabilities in the application."
Add-Paragraph -Text "Real-Time Chat Module:" -Bold
Add-Paragraph -Text "The real-time chat module allows users to send and receive messages instantly. Socket.IO manages the real-time connection, while the message collection stores chat history. This feature improves coordination between users and supports more dynamic digital interaction than a traditional static dashboard."
Add-Paragraph -Text "Notification Module:" -Bold
Add-Paragraph -Text "Notifications are generated for important actions such as welcome events, appointment updates, and order updates. Users can retrieve their notification list and mark items as read. This keeps the system interactive and ensures that users stay aware of important changes."
Add-PageBreak

Add-Paragraph -Text "Chapter 14: Database Design" -Style -2 -Bold
Add-Paragraph -Text "Database design is a critical part of the Smart Hospital Management System because every feature depends on well-structured data. MongoDB is used as the storage layer, and Mongoose schemas are defined for each main entity. The use of a document-based database allows flexibility while still preserving structure through schema definitions."
Add-Paragraph -Text "User Collection:" -Bold
Add-Paragraph -Text "Stores identity and authentication-related fields such as name, email, password, role, phone, address, OTP fields, verification status, and cart items. It acts as the base collection for all system roles."
Add-Paragraph -Text "Patient Collection:" -Bold
Add-Paragraph -Text "Stores medical profile information associated with a user account. This includes age, gender, disease, blood group, reports, and prescriptions."
Add-Paragraph -Text "Appointment Collection:" -Bold
Add-Paragraph -Text "Stores the relationship between patients and doctors with appointment date, time, reason, status, and e-prescription data."
Add-Paragraph -Text "Order Collection:" -Bold
Add-Paragraph -Text "Stores cart checkout results such as order items, payment method, payment status, and current order state."
Add-Paragraph -Text "Message Collection:" -Bold
Add-Paragraph -Text "Stores sender, receiver, and message text for real-time and historical communication."
Add-Paragraph -Text "Notification Collection:" -Bold
Add-Paragraph -Text "Stores user-specific alerts with title, message, and read status."
Add-Paragraph -Text "[ER Diagram Placeholder: Users linked to Patients, Appointments, Orders, Messages, and Notifications]" -Center
Add-Table -Headers @("Collection", "Main Fields", "Purpose") -Rows @(
  @("Users", "name, email, password, role, otp, cart", "Authentication and identity management"),
  @("Patients", "user, age, gender, bloodGroup, disease", "Medical profile storage"),
  @("Appointments", "patient, doctor, date, time, status, reason", "Visit scheduling and tracking"),
  @("Orders", "user, items, totalAmount, paymentStatus", "Medical order processing"),
  @("Messages", "sender, receiver, text", "Chat history"),
  @("Notifications", "user, title, message, read", "User alerts and updates")
)

Add-Paragraph -Text "Chapter 15: Implementation Details" -Style -2 -Bold
Add-Paragraph -Text "Backend Implementation:" -Bold
Add-Paragraph -Text "The backend starts with a structured folder organization containing configuration, models, controllers, routes, middleware, utilities, and seed data. Database connectivity is established through a dedicated configuration file. Middleware handles route protection and role-based authorization. Controllers are written separately for each feature domain, which keeps the application logic clean and modular."
Add-Paragraph -Text "Authentication is handled with JWT and bcrypt. OTP generation is implemented through controller logic and a utility that can send email via SMTP if configured. The backend also includes routes for patient CRUD, appointment creation and updates, order processing, file uploads using Multer, notification retrieval, and message persistence for chat history."
Add-Paragraph -Text "Frontend Implementation:" -Bold
Add-Paragraph -Text "The frontend is built using React with JSX and Vite. Routing is managed with React Router, allowing separate pages for login, registration, and dashboard access. Shared UI elements are divided into reusable components such as Navbar, SectionCard, NotificationBell, MapView, and ChatBox."
Add-Paragraph -Text "State management is intentionally simple. An authentication context stores the current logged-in user and provides helper functions for login, registration, OTP verification, and logout. Axios is configured with an interceptor to attach the JWT token to API requests automatically. The dashboard page loads records from the backend and displays them in different tab sections depending on the user role."
Add-Paragraph -Text "Real-Time Feature Implementation:" -Bold
Add-Paragraph -Text "Socket.IO is integrated both on the backend server and in the chat component on the frontend. When a user opens the chat section, a socket connection is established and associated with the current user. Messages are stored in MongoDB through API calls and also emitted in real time to online recipients."
Add-Paragraph -Text "Seed Data Implementation:" -Bold
Add-Paragraph -Text "To simplify demonstration and testing, a seed script inserts dummy records for admin, doctor, and patient accounts along with sample patient profile data, one appointment, one order, notifications, and a chat message. This ensures that the project can be shown immediately without requiring full manual data entry."
Add-PageBreak

Add-Paragraph -Text "Chapter 16: Algorithms and Process Flow" -Style -2 -Bold
Add-Paragraph -Text "Registration and Login Flow:" -Bold
Add-Bullets -Items @(
  "User fills registration form and submits details",
  "Backend checks whether the email already exists",
  "Password is hashed and user account is stored",
  "User logs in using email and password",
  "Backend validates credentials and generates JWT",
  "Frontend stores token and grants access to protected pages"
)
Add-Paragraph -Text "OTP Verification Flow:" -Bold
Add-Bullets -Items @(
  "User requests OTP using registered email",
  "System generates a six-digit code",
  "OTP and expiry time are stored temporarily in the database",
  "Email utility attempts delivery or falls back to demo display",
  "User enters OTP for verification",
  "System checks accuracy and validity period",
  "On success, user receives authenticated session data"
)
Add-Paragraph -Text "Appointment Booking Flow:" -Bold
Add-Bullets -Items @(
  "Admin or authorized user selects patient and doctor",
  "Date, time, and reason are entered",
  "Backend creates an appointment record",
  "Notification is generated for the patient",
  "Appointment appears in dashboard lists and can be updated later"
)
Add-Paragraph -Text "Order Processing Flow:" -Bold
Add-Bullets -Items @(
  "User selects medical store items",
  "Items are added to the cart within the user record",
  "User places order through COD checkout",
  "System calculates total amount and creates an order record",
  "Cart is cleared and notification is created",
  "Admin may later update order status such as delivered"
)
Add-Paragraph -Text "Chat Flow:" -Bold
Add-Bullets -Items @(
  "User opens chat module and selects a contact",
  "Previous message history is fetched from the database",
  "Socket connection joins the active user session",
  "New message is sent through API and socket emit",
  "Receiver gets the message instantly if online",
  "Message also remains stored for later retrieval"
)
Add-Paragraph -Text "[Flowchart Placeholder: Login, OTP, Appointment, Order, and Chat process diagrams]" -Center
Add-PageBreak

Add-Paragraph -Text "Chapter 17: Testing and Validation" -Style -2 -Bold
Add-Paragraph -Text "The project was tested manually using dummy data and role-based accounts. Since the application consists of multiple integrated modules, feature testing was carried out module by module to ensure that actions performed in one part of the system correctly affected related parts such as notifications and status updates."
Add-Table -Headers @("Test Case", "Input", "Expected Result", "Status") -Rows @(
  @("User Registration", "Valid new user details", "Account is created successfully", "Pass"),
  @("User Login", "Correct email and password", "JWT-based session starts", "Pass"),
  @("OTP Send", "Registered email", "OTP is generated and returned or sent", "Pass"),
  @("OTP Verify", "Correct OTP", "User verification succeeds", "Pass"),
  @("Admin Add Patient", "Patient details form", "Patient record is added", "Pass"),
  @("Appointment Booking", "Patient, doctor, date, reason", "Appointment is stored and listed", "Pass"),
  @("Search and Filter", "Name or status query", "Matching data is displayed", "Pass"),
  @("Add to Cart", "Medical store item", "Item appears in cart", "Pass"),
  @("Place Order", "Cart with items", "Order is created with COD", "Pass"),
  @("File Upload", "Supported file", "File path is returned successfully", "Pass"),
  @("Real-Time Chat", "Send message", "Message is stored and received live", "Pass"),
  @("Notification Read", "Mark read action", "Notification updates to read state", "Pass")
)
Add-Paragraph -Text "Validation Strategy:" -Bold
Add-Bullets -Items @(
  "Required input fields were checked before record creation",
  "Protected routes were validated using JWT middleware",
  "Role-based access restrictions were enforced for admin-only operations",
  "OTP expiry was checked before verification",
  "Empty-cart order placement was prevented"
)

Add-Paragraph -Text "Chapter 18: Results and Discussion" -Style -2 -Bold
Add-Paragraph -Text "The Smart Hospital Management System successfully demonstrates the digital handling of multiple hospital workflows in a single application. The integration of authentication, appointment management, patient CRUD, order processing, uploads, map display, notifications, and chat shows that a wide variety of operational tasks can be coordinated effectively through a full stack web solution."
Add-Paragraph -Text "The application is especially useful from an educational perspective because it shows the interaction between frontend state management, API communication, database models, middleware protection, and real-time messaging. It also highlights the benefit of organizing backend code with the MVC pattern. Each module remains focused, and updates can be made without heavily disturbing unrelated parts of the system."
Add-Paragraph -Text "From the usability perspective, the system provides a clear flow: users authenticate, access role-based dashboard sections, perform the actions available to them, and receive notifications when important events happen. The interface is intentionally simple rather than heavily stylized, which makes it suitable for demonstrations and further customization."
Add-Paragraph -Text "[Screenshot Placeholder: Login Page, Register Page, Dashboard, Appointments, Orders, Uploads, Chat, Notifications, Map]" -Center

Add-Paragraph -Text "Chapter 19: Advantages" -Style -2 -Bold
Add-Bullets -Items @(
  "Reduces dependency on manual files and registers",
  "Centralizes multiple hospital workflows in one system",
  "Improves accessibility of patient and appointment information",
  "Supports secure login and role-based access",
  "Enhances communication through chat and notifications",
  "Provides maintainable project structure using MVC",
  "Allows future expansion to more advanced healthcare features"
)

Add-Paragraph -Text "Chapter 20: Limitations" -Style -2 -Bold
Add-Bullets -Items @(
  "Current payment handling is limited to Cash on Delivery",
  "Uploaded files are stored locally rather than on cloud storage",
  "The user interface is functional but intentionally basic",
  "Telemedicine and wearable device support are not yet implemented",
  "Advanced analytics and reporting dashboards are not included"
)

Add-Paragraph -Text "Chapter 21: Future Scope" -Style -2 -Bold
Add-Paragraph -Text "The current project provides a strong base for future improvements. The most natural extension is telemedicine, where patients could book and attend video consultations remotely. Another major enhancement would be wearable device integration, allowing selected patient health parameters such as heart rate, oxygen saturation, or step count to be collected and displayed through the application."
Add-Paragraph -Text "Cloud storage integration would allow uploaded reports and prescriptions to be stored more securely and accessed from multiple devices. A payment gateway could be added for online transactions. PDF generation for bills, prescriptions, and reports would make the system more realistic for hospital operations. SMS and email reminder automation could further improve patient engagement and appointment attendance."
Add-Paragraph -Text "In the long term, the project can also be extended with analytics dashboards, mobile app versions, laboratory integration, doctor availability scheduling, insurance support, and multilingual interfaces. Such additions would help transform the project from a functional academic prototype into a broader healthcare platform."

Add-Paragraph -Text "Chapter 22: Conclusion" -Style -2 -Bold
Add-Paragraph -Text "The Smart Hospital Management System is a practical and structured full stack project that addresses the issue of manual hospital record management. By using the MERN stack and organizing backend code through the MVC pattern, the project achieves a balance between simplicity, functionality, and maintainability."
Add-Paragraph -Text "The system includes essential features such as user authentication, OTP verification, admin CRUD operations, patient and appointment management, report upload, order handling with COD, notifications, geo-location mapping, and real-time chat. These features together demonstrate how a digital platform can improve hospital workflow, data accessibility, and communication."
Add-Paragraph -Text "Although the current version is intentionally simple and academic in nature, it establishes a solid foundation for future growth. With additional modules and integrations, it can evolve into a more comprehensive healthcare management solution."

Add-Paragraph -Text "References" -Style -2 -Bold
Add-Bullets -Items @(
  "MongoDB Official Documentation",
  "Mongoose Official Documentation",
  "Express.js Official Documentation",
  "React Official Documentation",
  "Node.js Official Documentation",
  "Socket.IO Official Documentation",
  "Leaflet Official Documentation",
  "OpenStreetMap Documentation",
  "JWT Documentation",
  "Multer Documentation"
)

$document.SaveAs([ref]$docxPath)
$document.ExportAsFixedFormat($pdfPath, 17)
$document.Close()
$word.Quit()
