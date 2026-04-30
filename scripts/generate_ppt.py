from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape


PROJECT_ROOT = Path(r"C:\Users\shrey\OneDrive\Documents\New project")
PPT_PATH = PROJECT_ROOT / "Smart_Hospital_Presentation.pptx"


slides = [
    {
        "title": "Smart Hospital Management System",
        "subtitle": [
            "A MERN Stack Based Full Stack Web Application",
            "Presented By: Your Name",
            "Roll Number: Your Roll Number",
            "Department: Your Department",
            "Guide: Guide Name",
            "Academic Year: 2025-2026",
        ],
    },
    {
        "title": "Introduction",
        "bullets": [
            "Hospitals handle large amounts of patient and administrative data every day.",
            "Manual record management is slow, repetitive, and difficult to maintain.",
            "Digital systems improve speed, accuracy, communication, and accessibility.",
            "This project provides a centralized platform for admin, doctors, and patients.",
        ],
    },
    {
        "title": "Problem Statement",
        "bullets": [
            "Manual hospital record management causes delays and operational errors.",
            "Patient records are difficult to maintain, search, and update quickly.",
            "Appointment booking and tracking are inefficient in manual workflows.",
            "Reports, billing, and communication often remain disconnected.",
            "A smart digital solution is needed to streamline hospital operations.",
        ],
    },
    {
        "title": "Proposed Solution",
        "bullets": [
            "Develop a full stack Smart Hospital Management System using MERN.",
            "Provide centralized patient and appointment management.",
            "Support secure login with email OTP verification.",
            "Enable admin CRUD operations and hospital workflow monitoring.",
            "Add real-time communication, notifications, file upload, and map integration.",
        ],
    },
    {
        "title": "Objectives",
        "bullets": [
            "Digitize hospital activities and reduce paperwork.",
            "Improve security through authentication and OTP verification.",
            "Manage patients, appointments, and orders efficiently.",
            "Improve communication through chat and notifications.",
            "Build a maintainable full stack application using MVC.",
        ],
    },
    {
        "title": "Technology Stack",
        "bullets": [
            "Frontend: React.js, JSX, CSS, React Router",
            "Backend: Node.js and Express.js",
            "Database: MongoDB with Mongoose",
            "Authentication: JWT, bcrypt, Email OTP",
            "Real-Time Chat: Socket.IO",
            "Uploads and Map: Multer, React Leaflet, OpenStreetMap",
        ],
    },
    {
        "title": "Why MERN Stack",
        "bullets": [
            "Uses JavaScript across the full application.",
            "Supports fast development and clean integration between layers.",
            "React provides reusable components and easy UI updates.",
            "MongoDB stores flexible healthcare-related data structures.",
            "Express and Node create lightweight and scalable APIs.",
        ],
    },
    {
        "title": "System Architecture",
        "bullets": [
            "Presentation Layer: React frontend for user interaction.",
            "Application Layer: Express routes, controllers, and middleware.",
            "Data Layer: MongoDB collections and Mongoose models.",
            "Real-Time Layer: Socket.IO for instant user-to-user chat.",
            "Overall flow: User -> Frontend -> Backend API -> Database",
        ],
    },
    {
        "title": "MVC Pattern",
        "bullets": [
            "Model: Defines the structure of application data.",
            "View: React pages and components that display data.",
            "Controller: Business logic and request handling.",
            "MVC improves maintainability, structure, and scalability.",
            "It keeps data, logic, and UI concerns clearly separated.",
        ],
    },
    {
        "title": "Main Features",
        "bullets": [
            "User Authentication and Email OTP Verification",
            "Admin CRUD Operations",
            "Search and Filter System",
            "Appointment Management and E-Prescription Support",
            "Order Processing with Add to Cart and COD",
            "File Upload, Map Integration, Chat, and Notifications",
        ],
    },
    {
        "title": "User Roles",
        "bullets": [
            "Admin: Manage users, patients, appointments, and orders.",
            "Doctor: View assigned appointments and communicate with users.",
            "Patient: Register, log in, upload reports, book appointments, and place orders.",
            "Role-based access ensures secure and focused workflows.",
        ],
    },
    {
        "title": "Modules of the System",
        "bullets": [
            "Authentication Module",
            "Admin Management Module",
            "Patient Management Module",
            "Appointment Module",
            "Order Processing Module",
            "File Upload, Chat, Notification, and Map Modules",
        ],
    },
    {
        "title": "Database Design",
        "bullets": [
            "Main collections: Users, Patients, Appointments, Orders, Messages, Notifications.",
            "Users store identity, role, OTP, and cart details.",
            "Patients store health profile information.",
            "Appointments link doctors and patients with schedule details.",
            "Orders, messages, and notifications support workflow and communication.",
        ],
    },
    {
        "title": "Order Processing System",
        "bullets": [
            "Users can browse medical store items.",
            "Selected items are added to cart and stored in the user record.",
            "Orders are placed using Cash on Delivery.",
            "Admins can track and update order status.",
            "Notifications are generated when orders are placed or updated.",
        ],
    },
    {
        "title": "Real-Time Chat and Notifications",
        "bullets": [
            "Socket.IO enables live user-to-user communication.",
            "Messages are stored in the database for conversation history.",
            "Notifications inform users about important actions.",
            "Examples include welcome messages, appointment updates, and order updates.",
        ],
    },
    {
        "title": "Additional Features",
        "bullets": [
            "Search and filter improve record accessibility.",
            "File upload allows digital report submission.",
            "Geo-location map shows the hospital location.",
            "Role-protected routes improve security.",
            "The interface is simple, structured, and easy to use.",
        ],
    },
    {
        "title": "Testing and Results",
        "bullets": [
            "Tested registration, login, OTP verification, and protected routes.",
            "Verified patient CRUD, appointment booking, and search workflows.",
            "Checked cart, order placement, file upload, and notification updates.",
            "Real-time chat worked with message storage and live delivery.",
            "The system successfully demonstrated the required project features.",
        ],
    },
    {
        "title": "Advantages",
        "bullets": [
            "Reduces paperwork and manual record handling.",
            "Improves accessibility of patient and appointment data.",
            "Centralizes multiple hospital workflows in one system.",
            "Enhances communication between users.",
            "Provides a strong base for future healthcare features.",
        ],
    },
    {
        "title": "Limitations and Future Scope",
        "bullets": [
            "Current payment support is limited to COD.",
            "File storage is local and not cloud-based yet.",
            "Telemedicine and wearable integration are future features.",
            "Possible enhancements include online payments, PDF generation, and mobile support.",
        ],
    },
    {
        "title": "Conclusion",
        "bullets": [
            "The Smart Hospital Management System replaces manual processes with a digital platform.",
            "It demonstrates secure, modular, and scalable MERN-based development.",
            "The project successfully combines hospital management, communication, and record handling.",
            "It forms a strong foundation for future healthcare system expansion.",
        ],
    },
    {
        "title": "Thank You",
        "subtitle": [
            "Thank You",
            "Any Questions?",
        ],
    },
]


def xml_text_runs(lines, font_size=1800, color="1F3B5B", bold=False):
    runs = []
    for line in lines:
        line = escape(line)
        b_attr = ' b="1"' if bold else ""
        runs.append(
            f'<a:p><a:r><a:rPr lang="en-US" sz="{font_size}"{b_attr} dirty="0" smtClean="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:rPr>'
            f'<a:t>{line}</a:t></a:r><a:endParaRPr lang="en-US" sz="{font_size}"/></a:p>'
        )
    return "".join(runs)


def title_slide_xml(title, subtitle_lines, idx):
    title_runs = xml_text_runs([title], font_size=2600, color="14324A", bold=True)
    sub_runs = xml_text_runs(subtitle_lines, font_size=1500, color="355C7D")
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
 <p:cSld>
  <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F5F9FC"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
  <p:spTree>
   <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
   <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
   <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="685800" y="914400"/><a:ext cx="7772400" cy="1371600"/></a:xfrm></p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle/>{title_runs}</p:txBody>
   </p:sp>
   <p:sp>
    <p:nvSpPr><p:cNvPr id="3" name="Subtitle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="914400" y="2514600"/><a:ext cx="7315200" cy="2971800"/></a:xfrm></p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle/>{sub_runs}</p:txBody>
   </p:sp>
   <p:sp>
    <p:nvSpPr><p:cNvPr id="4" name="Accent"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="685800" y="1828800"/><a:ext cx="1828800" cy="68580"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="2A6F97"/></a:solidFill>
    </p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
   </p:sp>
  </p:spTree>
 </p:cSld>
 <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''


def bullet_slide_xml(title, bullets, idx):
    title_runs = xml_text_runs([title], font_size=2400, color="14324A", bold=True)
    bullet_xml = []
    for bullet in bullets:
        bullet = escape(bullet)
        bullet_xml.append(
            '<a:p>'
            '<a:pPr marL="342900" indent="-171450"><a:buChar char="•"/></a:pPr>'
            '<a:r><a:rPr lang="en-US" sz="1700" dirty="0"><a:solidFill><a:srgbClr val="23415A"/></a:solidFill></a:rPr>'
            f'<a:t>{bullet}</a:t></a:r>'
            '<a:endParaRPr lang="en-US" sz="1700"/></a:p>'
        )
    bullet_text = "".join(bullet_xml)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
 <p:cSld>
  <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
  <p:spTree>
   <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
   <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
   <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="685800" y="457200"/><a:ext cx="7772400" cy="914400"/></a:xfrm></p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle/>{title_runs}</p:txBody>
   </p:sp>
   <p:sp>
    <p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="868680" y="1463040"/><a:ext cx="7223760" cy="4206240"/></a:xfrm>
      <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="F5F9FC"/></a:solidFill>
      <a:ln><a:solidFill><a:srgbClr val="D7E6F0"/></a:solidFill></a:ln>
    </p:spPr>
    <p:txBody><a:bodyPr wrap="square" lIns="228600" tIns="171450" rIns="228600" bIns="171450"/><a:lstStyle/>{bullet_text}</p:txBody>
   </p:sp>
   <p:sp>
    <p:nvSpPr><p:cNvPr id="4" name="Accent"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="685800" y="1257300"/><a:ext cx="1143000" cy="57150"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="2A6F97"/></a:solidFill>
    </p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
   </p:sp>
  </p:spTree>
 </p:cSld>
 <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''


def slide_rels_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'''


def content_types_xml(slide_count):
    slide_overrides = "\n".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  {slide_overrides}
</Types>'''


def root_rels_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''


def presentation_xml(slide_count):
    slide_ids = "\n".join(
        f'<p:sldId id="{256 + i}" r:id="rId{i}"/>' for i in range(1, slide_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 saveSubsetFonts="1" autoCompressPictures="0">
  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:sldIdLst>
    {slide_ids}
  </p:sldIdLst>
</p:presentation>'''


def presentation_rels_xml(slide_count):
    rels = "\n".join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {rels}
</Relationships>'''


def theme_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Simple Theme">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="14324A"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="23415A"/></a:dk2>
      <a:lt2><a:srgbClr val="F5F9FC"/></a:lt2>
      <a:accent1><a:srgbClr val="2A6F97"/></a:accent1>
      <a:accent2><a:srgbClr val="4D8FB3"/></a:accent2>
      <a:accent3><a:srgbClr val="89C2D9"/></a:accent3>
      <a:accent4><a:srgbClr val="BDE0FE"/></a:accent4>
      <a:accent5><a:srgbClr val="7FB3D5"/></a:accent5>
      <a:accent6><a:srgbClr val="355C7D"/></a:accent6>
      <a:hlink><a:srgbClr val="2A6F97"/></a:hlink>
      <a:folHlink><a:srgbClr val="355C7D"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>'''


def core_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Smart Hospital Management System Presentation</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
</cp:coreProperties>'''


def app_xml(slide_count):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office PowerPoint</Application>
  <Slides>{slide_count}</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <PresentationFormat>On-screen Show (4:3)</PresentationFormat>
</Properties>'''


with ZipFile(PPT_PATH, "w", ZIP_DEFLATED) as zf:
    zf.writestr("[Content_Types].xml", content_types_xml(len(slides)))
    zf.writestr("_rels/.rels", root_rels_xml())
    zf.writestr("ppt/presentation.xml", presentation_xml(len(slides)))
    zf.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml(len(slides)))
    zf.writestr("ppt/theme/theme1.xml", theme_xml())
    zf.writestr("docProps/core.xml", core_xml())
    zf.writestr("docProps/app.xml", app_xml(len(slides)))

    for idx, slide in enumerate(slides, start=1):
        if "bullets" in slide:
            xml = bullet_slide_xml(slide["title"], slide["bullets"], idx)
        else:
            xml = title_slide_xml(slide["title"], slide["subtitle"], idx)
        zf.writestr(f"ppt/slides/slide{idx}.xml", xml)
        zf.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", slide_rels_xml())
