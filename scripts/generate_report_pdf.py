from pathlib import Path
import textwrap


PROJECT_ROOT = Path(r"C:\Users\shrey\OneDrive\Documents\New project")
PDF_PATH = PROJECT_ROOT / "Smart_Hospital_Report.pdf"
TXT_PATH = PROJECT_ROOT / "Smart_Hospital_Report_Source.txt"

PAGE_WIDTH = 595
PAGE_HEIGHT = 842
MARGIN = 56


def pdf_escape(text):
    return (
        text.replace("\\", "\\\\")
        .replace("(", "\\(")
        .replace(")", "\\)")
    )


class PDFWriter:
    def __init__(self):
        self.pages = []
        self.current_page = []
        self.y = PAGE_HEIGHT - MARGIN
        self.fonts = {
            "regular": "F1",
            "bold": "F2",
            "italic": "F3",
        }

    def new_page(self):
        if self.current_page:
            self.pages.append(self.current_page)
        self.current_page = []
        self.y = PAGE_HEIGHT - MARGIN

    def ensure_space(self, lines=1, leading=16):
        if self.y - (lines * leading) < MARGIN:
            self.new_page()

    def add_line(self, text, size=12, font="regular", x=MARGIN, leading=16):
        self.ensure_space(1, leading)
        cmd = f"BT /{self.fonts[font]} {size} Tf 1 0 0 1 {x} {self.y} Tm ({pdf_escape(text)}) Tj ET"
        self.current_page.append(cmd)
        self.y -= leading

    def add_centered(self, text, size=12, font="regular", leading=16):
        approx_width = len(text) * size * 0.43
        x = max(MARGIN, (PAGE_WIDTH - approx_width) / 2)
        self.add_line(text, size=size, font=font, x=x, leading=leading)

    def add_paragraph(self, text, size=12, font="regular", indent=0, leading=18, gap=6):
        width_chars = max(40, int((PAGE_WIDTH - 2 * MARGIN - indent) / (size * 0.52)))
        lines = textwrap.wrap(text, width=width_chars)
        self.ensure_space(max(1, len(lines)), leading)
        for line in lines:
            self.add_line(line, size=size, font=font, x=MARGIN + indent, leading=leading)
        self.y -= gap

    def add_bullets(self, items, size=12):
        for item in items:
            wrapped = textwrap.wrap(item, width=78)
            if not wrapped:
                continue
            self.add_line(f"- {wrapped[0]}", size=size, x=MARGIN + 10, leading=18)
            for line in wrapped[1:]:
                self.add_line(line, size=size, x=MARGIN + 26, leading=18)
            self.y -= 3
        self.y -= 6

    def add_spacer(self, amount=12):
        self.y -= amount
        if self.y < MARGIN:
            self.new_page()

    def finalize(self):
        if self.current_page:
            self.pages.append(self.current_page)

        objects = []

        def add_object(data):
            objects.append(data)
            return len(objects)

        font_regular = add_object("<< /Type /Font /Subtype /Type1 /BaseFont /Times-Roman >>")
        font_bold = add_object("<< /Type /Font /Subtype /Type1 /BaseFont /Times-Bold >>")
        font_italic = add_object("<< /Type /Font /Subtype /Type1 /BaseFont /Times-Italic >>")

        page_refs = []
        content_refs = []

        for page in self.pages:
            stream = "\n".join(page)
            content = f"<< /Length {len(stream.encode('latin-1', errors='replace'))} >>\nstream\n{stream}\nendstream"
            content_ref = add_object(content)
            content_refs.append(content_ref)
            page_refs.append(None)

        pages_ref_placeholder = add_object("PAGES_PLACEHOLDER")

        for index, content_ref in enumerate(content_refs):
            page_obj = (
                f"<< /Type /Page /Parent {pages_ref_placeholder} 0 R "
                f"/MediaBox [0 0 {PAGE_WIDTH} {PAGE_HEIGHT}] "
                f"/Resources << /Font << /F1 {font_regular} 0 R /F2 {font_bold} 0 R /F3 {font_italic} 0 R >> >> "
                f"/Contents {content_ref} 0 R >>"
            )
            page_refs[index] = add_object(page_obj)

        kids = " ".join(f"{ref} 0 R" for ref in page_refs)
        objects[pages_ref_placeholder - 1] = f"<< /Type /Pages /Kids [{kids}] /Count {len(page_refs)} >>"
        catalog_ref = add_object(f"<< /Type /Catalog /Pages {pages_ref_placeholder} 0 R >>")

        output = [b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"]
        offsets = [0]
        current = len(output[0])

        for i, obj in enumerate(objects, start=1):
          encoded = f"{i} 0 obj\n{obj}\nendobj\n".encode("latin-1", errors="replace")
          offsets.append(current)
          output.append(encoded)
          current += len(encoded)

        xref_start = current
        xref = [f"xref\n0 {len(objects) + 1}\n".encode("latin-1")]
        xref.append(b"0000000000 65535 f \n")
        for off in offsets[1:]:
            xref.append(f"{off:010} 00000 n \n".encode("latin-1"))
        trailer = (
            f"trailer\n<< /Size {len(objects) + 1} /Root {catalog_ref} 0 R >>\n"
            f"startxref\n{xref_start}\n%%EOF"
        ).encode("latin-1")
        output.extend(xref)
        output.append(trailer)
        PDF_PATH.write_bytes(b"".join(output))


writer = PDFWriter()


def heading(text):
    writer.add_paragraph(text, size=16, font="bold", leading=22, gap=4)


def subheading(text):
    writer.add_paragraph(text, size=13, font="bold", leading=18, gap=2)


source_lines = []


def add_source(text=""):
    source_lines.append(text)


writer.add_spacer(80)
writer.add_centered("SMART HOSPITAL MANAGEMENT SYSTEM", size=22, font="bold", leading=28)
writer.add_centered("Detailed Project Report", size=16, font="bold", leading=22)
writer.add_spacer(30)
for line in [
    "Submitted in partial fulfillment of the requirements for the award of the degree/course",
    "Submitted By: Your Name",
    "Roll Number: Your Roll Number",
    "Department: Your Department",
    "College: Your College Name",
    "Project Guide: Guide Name",
    "Academic Year: 2025-2026",
]:
    writer.add_centered(line, size=13, leading=20)
    add_source(line)
writer.new_page()

heading("Certificate")
for para in [
    "This is to certify that the project report entitled Smart Hospital Management System submitted by Your Name is a bona fide work carried out under my guidance and supervision in partial fulfillment of the requirements for the award of the degree or course during the academic year 2025-2026.",
    "To the best of my knowledge, the work presented in this report is original and has not been submitted elsewhere for the award of any degree, diploma, or certificate.",
]:
    writer.add_paragraph(para)
    add_source(para)
writer.add_spacer(40)
writer.add_paragraph("Project Guide Signature: ____________________", font="bold")
writer.add_paragraph("Head of Department Signature: ____________________", font="bold")
writer.new_page()

heading("Declaration")
for para in [
    "I hereby declare that the project report entitled Smart Hospital Management System is an original piece of work carried out by me under the guidance of my project mentor. The report has been prepared for academic purposes and has not been submitted earlier, either in full or in part, to any institution or university for the award of any degree or diploma.",
    "I further declare that all sources of information used in this report have been duly acknowledged.",
]:
    writer.add_paragraph(para)
    add_source(para)
writer.add_spacer(40)
writer.add_paragraph("Student Signature: ____________________", font="bold")
writer.new_page()

heading("Acknowledgement")
for para in [
    "I would like to express my sincere gratitude to my project guide, faculty members, and department for their continuous support, guidance, and encouragement throughout the development of this project. Their insights helped me understand both the technical and practical aspects of designing a real-world hospital management solution.",
    "I am also thankful to my classmates and friends for their suggestions during the planning, development, and testing phases. Their feedback allowed me to improve the structure and usability of the project. Finally, I would like to thank my family for their constant support and motivation during the completion of this work.",
]:
    writer.add_paragraph(para)
    add_source(para)
writer.new_page()

heading("Abstract")
abstract_paras = [
    "The Smart Hospital Management System is a web-based full stack application designed to digitize and simplify hospital record management. Many small and medium hospitals still depend on manual files, registers, and disconnected systems to manage patients, appointments, prescriptions, reports, billing, and communication. Such manual processes often result in delays, record duplication, difficulty in retrieval, and reduced service quality.",
    "The proposed system addresses these problems through a centralized platform developed using the MERN stack, namely MongoDB, Express.js, React.js, and Node.js. The backend follows the MVC pattern to maintain clear separation between data models, business logic, and routing layers. The system supports three primary roles: administrator, doctor, and patient. It includes modules for secure authentication, email OTP verification, patient and appointment management, admin CRUD operations, file upload, order processing with cart and Cash on Delivery support, search and filtering, geo-location map integration, real-time chat, and notifications.",
    "The objective of the project is to reduce paperwork, improve data accessibility, enhance communication, and create a structured application that can be understood, maintained, and extended easily. The current implementation forms a strong foundation for future modules such as telemedicine, wearable integration, cloud document management, and digital payment systems.",
]
for para in abstract_paras:
    writer.add_paragraph(para)
    add_source(para)
writer.new_page()

heading("Table of Contents")
toc = [
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
    "23. References",
    "24. Appendix A: Module Summary",
    "25. Appendix B: Sample API Endpoints",
    "26. Appendix C: Suggested Screenshots and Diagrams",
]
writer.add_bullets(toc)
writer.new_page()

sections = [
    (
        "Chapter 1: Introduction",
        [
            "Healthcare institutions manage sensitive and high-volume information every day. From the moment a patient registers at the reception desk to the moment treatment is completed, the hospital generates and consumes multiple forms of data including demographic information, appointment schedules, prescriptions, laboratory reports, billing details, doctor notes, and communication records. In a manual environment, these pieces of information are often scattered across paper files or isolated systems, which leads to inefficiency.",
            "Hospital administration is not limited to medical treatment alone. It includes workflow coordination, scheduling, record safety, quick retrieval of data, communication between stakeholders, and operational visibility for the management team. When these activities are not digitized, staff spend extra time searching records, manually verifying appointments, maintaining files, and coordinating with patients.",
            "The Smart Hospital Management System has been designed as a practical web-based solution to these operational challenges. It provides a centralized application where an administrator can manage records, a doctor can view assigned appointments, and a patient can access personal services such as appointments, reports, notifications, and orders. The project emphasizes clarity, simplicity, and maintainability, making it suitable both as an academic submission and as a prototype for future development.",
        ],
    ),
    (
        "Chapter 2: Problem Statement",
        [
            "The problem considered in this project is the continued dependence on manual hospital record management. Manual systems are slow, repetitive, and highly dependent on staff availability. Searching for patient history, managing doctor schedules, handling reports, or tracking appointment status becomes difficult when records are not maintained in a centralized digital environment.",
            "In addition, communication between hospital users is often fragmented. Patients may need to visit physically or rely on calls for updates. Notification systems may be absent. Reports and prescriptions can be misplaced. Billing and order-related activities may operate separately from the rest of the workflow. These issues directly affect efficiency, accuracy, and the overall patient experience.",
        ],
    ),
    (
        "Chapter 3: Need for the System",
        [
            "A hospital management system is needed to introduce consistency, speed, and traceability into healthcare operations. Digitization reduces dependency on paper-based records and improves the ability of hospital staff to access the right information at the right time. By organizing patient data, appointments, orders, and notifications inside one platform, the hospital can operate with better control and less duplication.",
            "The proposed project specifically addresses the need for a beginner-friendly but functional digital solution. It is not designed as an overly complex enterprise platform; instead, it is structured to demonstrate the most important components of a smart hospital workflow in a way that is easy to implement, explain, and improve further.",
        ],
    ),
    (
        "Chapter 4: Existing System",
        [
            "In the existing manual setup, most activities are performed through physical files, registers, or basic spreadsheets. Patient registration is done manually at the reception. Doctor appointments are noted in diaries or separate sheets. Test reports may be stored physically or in unstructured folders. There is often no direct connection between appointment data, prescription details, and follow-up communication.",
            "This causes several operational issues. Record duplication becomes common because old data cannot be found quickly. Appointments may overlap or be delayed. The medical store or billing counter may not have access to complete treatment context. Communication becomes reactive instead of organized, and the hospital lacks a reliable dashboard for monitoring daily operations.",
        ],
    ),
]

for title, paras in sections:
    heading(title)
    add_source(title)
    for para in paras:
        writer.add_paragraph(para)
        add_source(para)

subheading("Limitations of the Existing System")
writer.add_bullets([
    "Heavy dependence on paper-based records",
    "Slow retrieval of patient and appointment history",
    "Greater risk of human error and data inconsistency",
    "No real-time communication between users",
    "Poor integration between appointments, reports, and orders",
    "Difficulty in scaling processes when patient volume increases",
])
writer.new_page()

heading("Chapter 5: Proposed System")
for para in [
    "The proposed Smart Hospital Management System centralizes the major workflows of a hospital through a browser-based application. The system allows users to log in securely, access features based on their role, and interact with hospital services digitally. It includes an administrator panel for CRUD operations, patient and appointment management, medical order handling, file upload, notifications, and a chat module.",
    "The proposed system emphasizes structured development. The backend is organized using the MVC pattern with clearly separated models, controllers, routes, and middleware. The frontend is built with reusable React components and pages using JSX. The result is a codebase that is simpler to understand and maintain.",
]:
    writer.add_paragraph(para)
    add_source(para)
subheading("Key Characteristics of the Proposed System")
writer.add_bullets([
    "Centralized data access",
    "Role-based user interaction",
    "Secure authentication and authorization",
    "Digital workflow for appointments and patient records",
    "Integrated notifications and communication",
    "Extensibility for future healthcare features",
])

heading("Chapter 6: Objectives")
writer.add_bullets([
    "To replace manual hospital administration with a digital system",
    "To reduce paperwork and repetitive tasks",
    "To create a secure authentication process using JWT and email OTP",
    "To manage patient records through admin-controlled CRUD operations",
    "To simplify appointment creation, tracking, and status updates",
    "To provide a medical order module with add-to-cart and COD checkout",
    "To support document uploads for patient reports",
    "To improve communication using real-time chat and notifications",
    "To create a maintainable MERN application following MVC principles",
])

heading("Chapter 7: Scope of the Project")
for para in [
    "The scope of the project includes the design and implementation of a hospital management prototype for digital record handling. It focuses on operations commonly needed in a clinic or hospital environment: authentication, patient data, appointments, chat, file upload, notifications, map display, and order management. The system is suitable for demonstration, academic evaluation, and as a base architecture for larger healthcare applications.",
    "The current scope intentionally avoids highly specialized medical modules such as insurance claim processing, advanced clinical decision support, laboratory device integration, telemedicine video streaming, and complex third-party billing infrastructure. These are reserved for future enhancement.",
]:
    writer.add_paragraph(para)
    add_source(para)

heading("Chapter 8: Literature and Background Study")
for para in [
    "Modern healthcare systems increasingly depend on digital platforms for operational efficiency. Hospital information systems combine patient registration, diagnosis records, appointment scheduling, billing, and communication into one integrated environment. The move toward digitalization is driven by the need for better service quality, reduced paperwork, stronger security, and easier reporting.",
    "Web-based systems offer multiple advantages in this context. They can be used through standard browsers, reduce installation overhead, and can support multiple departments or user roles. Real-time communication technologies, document upload facilities, and centralized databases make such platforms more practical than isolated desktop tools.",
    "The MERN stack is well suited to this use case because it uses JavaScript across the entire application. React helps build responsive user interfaces, Express and Node create a lightweight backend service layer, and MongoDB offers flexible document storage for diverse entities such as appointments, messages, and notifications. The MVC pattern supports long-term maintainability by separating responsibilities.",
]:
    writer.add_paragraph(para)
    add_source(para)
writer.new_page()

heading("Chapter 9: Requirement Analysis")
subheading("Functional Requirements")
writer.add_bullets([
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
    "Notification generation and read status management",
])
subheading("Non-Functional Requirements")
writer.add_bullets([
    "The system should be easy to understand and use",
    "The system should provide secure route protection",
    "The codebase should be modular and maintainable",
    "The frontend should work on common desktop and laptop screens",
    "The application should support future scalability",
])
subheading("Hardware and Software Requirements")
writer.add_bullets([
    "Computer or laptop with at least 4 GB RAM",
    "Modern web browser",
    "Node.js and npm",
    "MongoDB",
    "Visual Studio Code or a similar editor",
])

heading("Chapter 10: Technology Stack")
for para in [
    "The project uses the MERN stack, which includes MongoDB, Express.js, React.js, and Node.js. This stack was selected because it supports full stack development using JavaScript and offers an efficient workflow for building web applications with a modern frontend and scalable backend.",
    "React.js is used to create the user interface with reusable JSX components. Node.js and Express.js are used to create REST API endpoints, middleware, and real-time server functionality. MongoDB stores application data in a flexible document format, which is ideal for evolving data structures such as notifications and chat messages. Socket.IO is used for real-time chat, while Multer handles file uploads. React Leaflet and OpenStreetMap provide map integration.",
]:
    writer.add_paragraph(para)
    add_source(para)
subheading("Technology Summary")
writer.add_bullets([
    "Frontend: React.js, JSX, CSS, React Router",
    "Backend: Node.js and Express.js",
    "Database: MongoDB with Mongoose",
    "Authentication: JWT, bcrypt, and OTP",
    "Real-Time Communication: Socket.IO",
    "File Upload: Multer",
    "Map Integration: React Leaflet and OpenStreetMap",
])
writer.new_page()

heading("Chapter 11: System Architecture")
for para in [
    "The architecture of the Smart Hospital Management System can be viewed as a layered structure. The presentation layer consists of the React frontend that collects user input and displays data. The application layer consists of Express routes, controllers, middleware, and helper functions. The data layer consists of MongoDB collections and Mongoose models. Real-time communication is handled through Socket.IO, which works alongside the normal HTTP request-response cycle.",
    "This layered separation improves maintainability and keeps the system organized. The frontend remains focused on interaction and display. The backend enforces rules, validation, and access control. The database stores persistent records. Real-time features are handled separately without mixing socket logic into every request.",
]:
    writer.add_paragraph(para)
    add_source(para)
writer.add_paragraph("[Architecture Diagram Placeholder: User -> React Frontend -> Express API -> MongoDB, plus Socket.IO for chat]", font="italic")

heading("Chapter 12: MVC Design Pattern")
for label, para in [
    ("Model", "Models define the structure of the application data. Each important entity in the system has its own schema: User, Patient, Appointment, Order, Message, and Notification. Models ensure that related fields are grouped together and stored consistently."),
    ("View", "In this project, the view layer is represented by the React frontend. It contains pages such as login, registration, and dashboard, along with reusable components such as Navbar, StatCard, ChatBox, and MapView."),
    ("Controller", "Controllers implement business logic. They receive input from routes, query or update the models, and send appropriate responses back to the frontend. For example, the authentication controller manages login, register, OTP sending, and verification. Other controllers manage appointments, patients, orders, notifications, and chat."),
]:
    subheading(label)
    writer.add_paragraph(para)
    add_source(para)
writer.add_paragraph("Using MVC reduces code duplication and makes the application easier to debug. It also helps in explaining the project because each file group has a clear responsibility.")

heading("Chapter 13: Modules of the System")
module_items = [
    ("Authentication Module", "The authentication module provides secure account creation and login functionality. Users can register with role information and then log in using email and password. Passwords are hashed using bcrypt before storage. JWT tokens are generated after successful authentication so the client can access protected API routes. Email OTP is included as an additional verification mechanism."),
    ("Admin Module", "The admin module acts as the control center of the application. The administrator can view the dashboard summary, inspect users, create patient records, book appointments, update status values, delete selected records, and track medical orders. This module demonstrates how CRUD operations can be implemented in a practical and role-protected manner."),
    ("Patient Management Module", "The patient module stores profile information such as age, gender, blood group, disease details, reports, and prescription references. This module links medical profile data to the base user account and supports structured digital record management."),
    ("Appointment Module", "The appointment module is used to create, update, and track visits between patients and doctors. Each appointment contains the selected patient, doctor, date, time, reason, status, and e-prescription field. Appointment states such as scheduled, completed, or cancelled help represent the treatment workflow clearly."),
    ("Search and Filter Module", "The search and filter module helps users quickly locate data by names, diseases, order items, and appointment status. This improves usability and reduces the time required to retrieve important information."),
    ("Order Processing Module", "This module simulates a hospital medical store experience. Users can add items to a cart, review selected products, and place an order using Cash on Delivery. The order record stores the purchased items, amount, payment status, and order status. The administrator can later update delivery-related states."),
    ("File Upload Module", "File upload is useful for medical reports and supporting documents. The system uses Multer to handle multipart form uploads and stores files locally on the server. The returned file path can be associated with patient records or displayed to the user for confirmation."),
    ("Geo-location Map Module", "The map module uses React Leaflet and OpenStreetMap to display the hospital location. Although simple, this feature enhances the user experience and demonstrates the integration of external mapping capabilities in the application."),
    ("Real-Time Chat Module", "The real-time chat module allows users to send and receive messages instantly. Socket.IO manages the real-time connection, while the message collection stores chat history. This feature improves coordination between users and supports more dynamic digital interaction."),
    ("Notification Module", "Notifications are generated for important actions such as welcome events, appointment updates, and order updates. Users can retrieve their notification list and mark items as read. This keeps the system interactive and ensures that users stay aware of important changes."),
]
for label, para in module_items:
    subheading(label)
    writer.add_paragraph(para)
    add_source(para)
writer.new_page()

heading("Chapter 14: Database Design")
for para in [
    "Database design is a critical part of the Smart Hospital Management System because every feature depends on well-structured data. MongoDB is used as the storage layer, and Mongoose schemas are defined for each main entity. The use of a document-based database allows flexibility while still preserving structure through schema definitions.",
    "The User collection stores identity and authentication-related fields such as name, email, password, role, phone, address, OTP fields, verification status, and cart items. The Patient collection stores profile information such as age, gender, disease, blood group, reports, and prescriptions. The Appointment collection stores the relationship between patients and doctors with appointment date, time, reason, status, and e-prescription data. The Order collection stores cart checkout results such as order items, payment method, payment status, and current order state. The Message collection stores sender, receiver, and message text for communication. The Notification collection stores user-specific alerts with title, message, and read status.",
]:
    writer.add_paragraph(para)
    add_source(para)
writer.add_paragraph("[ER Diagram Placeholder: Users linked to Patients, Appointments, Orders, Messages, and Notifications]", font="italic")

heading("Chapter 15: Implementation Details")
impl_paras = [
    "The backend starts with a structured folder organization containing configuration, models, controllers, routes, middleware, utilities, and seed data. Database connectivity is established through a dedicated configuration file. Middleware handles route protection and role-based authorization. Controllers are written separately for each feature domain, which keeps the application logic clean and modular.",
    "Authentication is handled with JWT and bcrypt. OTP generation is implemented through controller logic and a utility that can send email via SMTP if configured. The backend also includes routes for patient CRUD, appointment creation and updates, order processing, file uploads using Multer, notification retrieval, and message persistence for chat history.",
    "The frontend is built using React with JSX and Vite. Routing is managed with React Router, allowing separate pages for login, registration, and dashboard access. Shared UI elements are divided into reusable components such as Navbar, SectionCard, NotificationBell, MapView, and ChatBox.",
    "State management is intentionally simple. An authentication context stores the current logged-in user and provides helper functions for login, registration, OTP verification, and logout. Axios is configured with an interceptor to attach the JWT token to API requests automatically. The dashboard page loads records from the backend and displays them in different tab sections depending on the user role.",
    "Socket.IO is integrated both on the backend server and in the chat component on the frontend. When a user opens the chat section, a socket connection is established and associated with the current user. Messages are stored in MongoDB through API calls and also emitted in real time to online recipients.",
    "To simplify demonstration and testing, a seed script inserts dummy records for admin, doctor, and patient accounts along with sample patient profile data, one appointment, one order, notifications, and a chat message. This ensures that the project can be shown immediately without requiring full manual data entry.",
]
for para in impl_paras:
    writer.add_paragraph(para)
    add_source(para)
writer.new_page()

heading("Chapter 16: Algorithms and Process Flow")
flow_sections = {
    "Registration and Login Flow": [
        "User fills registration form and submits details",
        "Backend checks whether the email already exists",
        "Password is hashed and user account is stored",
        "User logs in using email and password",
        "Backend validates credentials and generates JWT",
        "Frontend stores token and grants access to protected pages",
    ],
    "OTP Verification Flow": [
        "User requests OTP using registered email",
        "System generates a six-digit code",
        "OTP and expiry time are stored temporarily in the database",
        "Email utility attempts delivery or falls back to demo display",
        "User enters OTP for verification",
        "System checks accuracy and validity period",
        "On success, user receives authenticated session data",
    ],
    "Appointment Booking Flow": [
        "Admin or authorized user selects patient and doctor",
        "Date, time, and reason are entered",
        "Backend creates an appointment record",
        "Notification is generated for the patient",
        "Appointment appears in dashboard lists and can be updated later",
    ],
    "Order Processing Flow": [
        "User selects medical store items",
        "Items are added to the cart within the user record",
        "User places order through COD checkout",
        "System calculates total amount and creates an order record",
        "Cart is cleared and notification is created",
        "Admin may later update order status such as delivered",
    ],
    "Chat Flow": [
        "User opens chat module and selects a contact",
        "Previous message history is fetched from the database",
        "Socket connection joins the active user session",
        "New message is sent through API and socket emit",
        "Receiver gets the message instantly if online",
        "Message also remains stored for later retrieval",
    ],
}
for label, items in flow_sections.items():
    subheading(label)
    writer.add_bullets(items)
writer.add_paragraph("[Flowchart Placeholder: Login, OTP, Appointment, Order, and Chat process diagrams]", font="italic")

heading("Chapter 17: Testing and Validation")
for para in [
    "The project was tested manually using dummy data and role-based accounts. Since the application consists of multiple integrated modules, feature testing was carried out module by module to ensure that actions performed in one part of the system correctly affected related parts such as notifications and status updates.",
    "Validation checks were applied at multiple levels. Required fields were checked before record creation, protected routes were guarded with JWT middleware, role-based access restrictions were enforced for admin-only operations, OTP expiry was validated before verification, and empty-cart order placement was prevented.",
]:
    writer.add_paragraph(para)
    add_source(para)
subheading("Representative Test Cases")
writer.add_bullets([
    "User Registration: valid new user details should create an account successfully",
    "User Login: correct credentials should start a JWT-based session",
    "OTP Send and Verify: registered email and correct OTP should verify successfully",
    "Admin Add Patient: valid patient data should create a patient record",
    "Appointment Booking: valid patient, doctor, and schedule should create an appointment",
    "Search and Filter: query text and status filters should return matching records",
    "Add to Cart and Place Order: selected items should be saved and converted into an order",
    "File Upload: supported file should upload and return a valid path",
    "Real-Time Chat: sent message should be stored and received live",
    "Notification Read: mark read action should update notification state",
])
writer.new_page()

heading("Chapter 18: Results and Discussion")
results_paras = [
    "The Smart Hospital Management System successfully demonstrates the digital handling of multiple hospital workflows in a single application. The integration of authentication, appointment management, patient CRUD, order processing, uploads, map display, notifications, and chat shows that a wide variety of operational tasks can be coordinated effectively through a full stack web solution.",
    "The application is especially useful from an educational perspective because it shows the interaction between frontend state management, API communication, database models, middleware protection, and real-time messaging. It also highlights the benefit of organizing backend code with the MVC pattern. Each module remains focused, and updates can be made without heavily disturbing unrelated parts of the system.",
    "From the usability perspective, the system provides a clear flow: users authenticate, access role-based dashboard sections, perform the actions available to them, and receive notifications when important events happen. The interface is intentionally simple rather than heavily stylized, which makes it suitable for demonstrations and further customization.",
]
for para in results_paras:
    writer.add_paragraph(para)
    add_source(para)
writer.add_paragraph("[Screenshot Placeholder: Login Page, Register Page, Dashboard, Appointments, Orders, Uploads, Chat, Notifications, Map]", font="italic")

heading("Chapter 19: Advantages")
writer.add_bullets([
    "Reduces dependency on manual files and registers",
    "Centralizes multiple hospital workflows in one system",
    "Improves accessibility of patient and appointment information",
    "Supports secure login and role-based access",
    "Enhances communication through chat and notifications",
    "Provides maintainable project structure using MVC",
    "Allows future expansion to more advanced healthcare features",
])

heading("Chapter 20: Limitations")
writer.add_bullets([
    "Current payment handling is limited to Cash on Delivery",
    "Uploaded files are stored locally rather than on cloud storage",
    "The user interface is functional but intentionally basic",
    "Telemedicine and wearable device support are not yet implemented",
    "Advanced analytics and reporting dashboards are not included",
])

heading("Chapter 21: Future Scope")
future_paras = [
    "The current project provides a strong base for future improvements. The most natural extension is telemedicine, where patients could book and attend video consultations remotely. Another major enhancement would be wearable device integration, allowing selected patient health parameters such as heart rate, oxygen saturation, or step count to be collected and displayed through the application.",
    "Cloud storage integration would allow uploaded reports and prescriptions to be stored more securely and accessed from multiple devices. A payment gateway could be added for online transactions. PDF generation for bills, prescriptions, and reports would make the system more realistic for hospital operations. SMS and email reminder automation could further improve patient engagement and appointment attendance.",
    "In the long term, the project can also be extended with analytics dashboards, mobile app versions, laboratory integration, doctor availability scheduling, insurance support, and multilingual interfaces. Such additions would help transform the project from a functional academic prototype into a broader healthcare platform.",
]
for para in future_paras:
    writer.add_paragraph(para)
    add_source(para)

heading("Chapter 22: Conclusion")
conclusion_paras = [
    "The Smart Hospital Management System is a practical and structured full stack project that addresses the issue of manual hospital record management. By using the MERN stack and organizing backend code through the MVC pattern, the project achieves a balance between simplicity, functionality, and maintainability.",
    "The system includes essential features such as user authentication, OTP verification, admin CRUD operations, patient and appointment management, report upload, order handling with COD, notifications, geo-location mapping, and real-time chat. These features together demonstrate how a digital platform can improve hospital workflow, data accessibility, and communication.",
    "Although the current version is intentionally simple and academic in nature, it establishes a solid foundation for future growth. With additional modules and integrations, it can evolve into a more comprehensive healthcare management solution.",
]
for para in conclusion_paras:
    writer.add_paragraph(para)
    add_source(para)

heading("Chapter 23: References")
writer.add_bullets([
    "MongoDB Official Documentation",
    "Mongoose Official Documentation",
    "Express.js Official Documentation",
    "React Official Documentation",
    "Node.js Official Documentation",
    "Socket.IO Official Documentation",
    "Leaflet Official Documentation",
    "OpenStreetMap Documentation",
    "JWT Documentation",
    "Multer Documentation",
])
writer.new_page()

heading("Appendix A: Module Summary")
appendix_a = [
    "Authentication: Handles registration, login, token creation, OTP generation, and OTP verification.",
    "Users: Stores role-based account details such as admin, doctor, and patient identity data.",
    "Patients: Maintains medical profile fields including age, gender, disease, blood group, reports, and prescriptions.",
    "Appointments: Schedules and tracks interactions between doctors and patients.",
    "Orders: Stores medical item purchases, COD payment state, and order progress.",
    "Messages: Keeps real-time chat history between users.",
    "Notifications: Stores alerts for important events such as appointment updates and order placement.",
    "Uploads: Accepts file submissions and stores report paths for later access.",
    "Map Module: Displays the hospital location through a web map interface.",
]
writer.add_bullets(appendix_a)
writer.add_paragraph("This appendix provides a concise snapshot of the major modules so that evaluators can quickly understand the structure of the project even before reviewing the codebase in detail.")

heading("Appendix B: Sample API Endpoints")
api_lines = [
    "POST /api/auth/register - Create a new account",
    "POST /api/auth/login - Log in with email and password",
    "POST /api/auth/send-otp - Generate and send OTP",
    "POST /api/auth/verify-otp - Verify OTP and sign in",
    "GET /api/users - List users for admin",
    "GET /api/users/contacts - Load chat contacts",
    "GET /api/patients - Fetch patient records",
    "POST /api/patients - Create a patient record",
    "GET /api/appointments - Fetch appointment records",
    "POST /api/appointments - Create an appointment",
    "GET /api/orders/cart - View cart",
    "POST /api/orders/cart - Add item to cart",
    "POST /api/orders/place - Place order with COD",
    "GET /api/notifications - Fetch notifications",
    "POST /api/chat - Save a chat message",
]
writer.add_bullets(api_lines)
writer.new_page()

heading("Appendix C: Suggested Screenshots and Diagrams")
for para in [
    "To improve the final report before submission, the following screenshots can be inserted at suitable points in the document: login page, register page, admin dashboard, create patient form, appointment section, medical store and cart section, upload section, map section, chat section, and notifications section.",
    "The following diagrams are also recommended: system architecture diagram, use case diagram, ER diagram, flowchart for login and OTP verification, and flowchart for order placement. Adding these visuals will improve readability and can help the final report comfortably span 25 to 30 PDF pages depending on formatting.",
    "This generated version has been written in a structured way so you can directly use it as the base report, then improve the final submission by inserting actual screenshots from your running application and replacing the placeholders with diagrams.",
]:
    writer.add_paragraph(para)
    add_source(para)

heading("Appendix D: Suggested Viva Discussion Points")
writer.add_bullets([
    "Why the MERN stack was chosen for this project",
    "How JWT protects private routes",
    "How OTP verification improves security",
    "Why MVC helps organize backend logic",
    "How Socket.IO differs from normal HTTP request-response flow",
    "How the order module can later support online payments",
    "How local uploads can be replaced with cloud storage",
    "Which future features would make the system closer to a real hospital platform",
])

writer.finalize()
TXT_PATH.write_text("\n".join(source_lines), encoding="utf-8")
