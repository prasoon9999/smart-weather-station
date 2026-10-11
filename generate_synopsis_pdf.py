import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT

def generate_pdf():
    pdf_filename = "Smart_Weather_Station_Synopsis_Report.pdf"
    
    # 40pt margins to give ample room for 4 exact pages matching sample
    doc = SimpleDocTemplate(
        pdf_filename,
        pagesize=letter,
        leftMargin=48,
        rightMargin=48,
        topMargin=42,
        bottomMargin=42
    )

    styles = getSampleStyleSheet()

    # --- CUSTOM PARAGRAPH STYLES ---
    style_cover_title = ParagraphStyle(
        'CoverTitle',
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=18,
        alignment=TA_CENTER,
        textColor=colors.black
    )

    style_cover_sub = ParagraphStyle(
        'CoverSub',
        fontName='Helvetica',
        fontSize=9.5,
        leading=14,
        alignment=TA_CENTER,
        textColor=colors.black
    )

    style_cover_bold = ParagraphStyle(
        'CoverBold',
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=15,
        alignment=TA_CENTER,
        textColor=colors.black
    )

    style_toc_title = ParagraphStyle(
        'TOCTitle',
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=18,
        alignment=TA_CENTER,
        textColor=colors.black
    )

    style_toc_item = ParagraphStyle(
        'TOCItem',
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=21,
        alignment=TA_LEFT,
        textColor=colors.black
    )

    style_section_heading = ParagraphStyle(
        'SectionHeading',
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=13,
        spaceBefore=6,
        spaceAfter=2,
        textColor=colors.black
    )

    style_body = ParagraphStyle(
        'Body',
        fontName='Helvetica',
        fontSize=8.5,
        leading=11.2,
        alignment=TA_JUSTIFY,
        spaceBefore=0,
        spaceAfter=3,
        textColor=colors.black
    )

    style_bullet = ParagraphStyle(
        'Bullet',
        fontName='Helvetica',
        fontSize=8.5,
        leading=11.2,
        leftIndent=10,
        firstLineIndent=-10,
        spaceBefore=1,
        spaceAfter=1.5,
        textColor=colors.black
    )

    style_table_header = ParagraphStyle(
        'TableHeader',
        fontName='Helvetica-Bold',
        fontSize=7.5,
        leading=9.5,
        alignment=TA_CENTER,
        textColor=colors.black
    )

    style_table_cell = ParagraphStyle(
        'TableCell',
        fontName='Helvetica',
        fontSize=7,
        leading=8.5,
        alignment=TA_CENTER,
        textColor=colors.black
    )

    story = []

    # ==========================================
    # PAGE 1: TITLE / COVER PAGE (Matches Sample Page 1)
    # ==========================================
    story.append(Spacer(1, 15))
    story.append(Paragraph("<b>SYNOPSIS REPORT ON</b>", style_cover_bold))
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>Smart Weather Station System Using ESP32, DHT22 &amp; Cloud Analytics</b>", style_cover_title))
    story.append(Spacer(1, 18))

    story.append(Paragraph("SUBMITTED IN PARTIAL FULFILLMENT OF THE<br/>REQUIREMENT FOR THE AWARD OF THE DEGREE OF THE", style_cover_sub))
    story.append(Spacer(1, 8))
    story.append(Paragraph("<b>BACHELOR OF TECHNOLOGY</b>", style_cover_bold))
    story.append(Spacer(1, 6))
    story.append(Paragraph("IN", style_cover_sub))
    story.append(Spacer(1, 6))
    story.append(Paragraph("<b>DEPARTMENT OF ELECTRONICS AND COMMUNICATION ENGINEERING</b>", style_cover_bold))
    story.append(Spacer(1, 18))

    story.append(Paragraph("By", style_cover_sub))
    story.append(Spacer(1, 6))
    story.append(Paragraph("<b>PRASOON – 2500971640050</b>", style_cover_bold))
    story.append(Spacer(1, 20))

    story.append(Paragraph("<b>GALGOTIAS COLLEGE OF ENGINEERING AND TECHNOLOGY</b>", style_cover_bold))
    story.append(Spacer(1, 5))
    story.append(Paragraph("AFFILIATED TO", style_cover_sub))
    story.append(Spacer(1, 5))
    story.append(Paragraph("<b>DR. A.P.J ABDUL KALAM TECHNICAL UNIVERSITY LUCKNOW,<br/>UTTAR PRADESH</b>", style_cover_bold))
    story.append(Spacer(1, 16))

    # Center Logos on Page 1
    logo_rows = []
    if os.path.exists("galgotias_logo.png"):
        logo_rows.append([Image("galgotias_logo.png", width=0.85*inch, height=0.85*inch)])
    if os.path.exists("aktu_logo.png"):
        logo_rows.append([Image("aktu_logo.png", width=0.75*inch, height=0.75*inch)])

    if logo_rows:
        logo_table = Table(logo_rows, colWidths=[6.5*inch])
        logo_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 2),
        ]))
        story.append(logo_table)

    story.append(PageBreak())

    # ==========================================
    # PAGE 2: TABLE OF CONTENT (Matches Sample Page 2)
    # ==========================================
    story.append(Spacer(1, 15))
    story.append(Paragraph("<b>TABLE OF CONTENT</b>", style_toc_title))
    story.append(Spacer(1, 20))

    toc_items = [
        "■ 1. Introduction (Brief Recap)",
        "■ 2. Motivation (Brief Recap)",
        "■ 3. Literature Survey (Brief Recap)",
        "■ 4. Problem Definition (Brief Recap)",
        "■ 5. Objectives (Brief Recap)",
        "■ 6. Proposed Solution (Brief Recap)",
        "■ 7. System Design / Architecture",
        "■ 8. Methodology / Workflow",
        "■ 9. Implementation",
        "■ 10. Testing and Validation",
        "■ 11. Results and Discussion",
        "■ 12. Work Completed",
        "■ 13. Remaining Work",
        "■ 14. References (IEEE Format)"
    ]

    for item in toc_items:
        story.append(Paragraph(item, style_toc_item))

    story.append(PageBreak())

    # ==========================================
    # PAGE 3: SECTIONS (Matches Sample Page 3)
    # ==========================================
    story.append(Paragraph("1. INTRODUCTION (BRIEF RECAP)", style_section_heading))
    story.append(Paragraph(
        "Modern environmental monitoring often involves tracking localized weather variations that impact human comfort, "
        "agriculture, server rooms, and smart buildings. Conventional weather stations are bulky, expensive, and lack cloud accessibility. "
        "The Smart Weather Station System is a lightweight IoT environmental monitoring tool designed to automate localized "
        "meteorological data acquisition. It measures atmospheric temperature and relative humidity via a DHT22 sensor, "
        "computes physiological comfort metrics (Heat Index and Dew Point) on-board using an ESP32 microcontroller, renders real-time telemetry "
        "locally on an SSD1306 OLED display, and streams the data over Wi-Fi to ThingSpeak cloud and a responsive Vercel web dashboard. "
        "The project emphasizes edge reliability, zero external library overhead, fast response, and accessible data visualization.",
        style_body
    ))

    story.append(Paragraph("2. MOTIVATION (BRIEF RECAP)", style_section_heading))
    story.append(Paragraph(
        "The project is motivated by the spatial gap left by centralized meteorological stations, which are typically situated tens of "
        "kilometers away at airports and fail to capture local microclimates. Automated compact IoT weather stations provide real-time ambient awareness "
        "while reducing equipment costs by over 90%.",
        style_body
    ))

    story.append(Paragraph("4. PROBLEM DEFINITION (BRIEF RECAP)", style_section_heading))
    story.append(Paragraph(
        "Traditional weather logging requires manual recording or complex standalone loggers that do not support automated remote access, "
        "creating delayed observations and making consistent monitoring difficult. The proposed system addresses this by automating "
        "atmospheric data acquisition, computing heat stress metrics on-chip, and streaming telemetry directly to open cloud APIs and browsers.",
        style_body
    ))

    story.append(Paragraph("6. PROPOSED SOLUTION (BRIEF RECAP)", style_section_heading))
    story.append(Paragraph(
        "The Smart Weather Station provides a sequential, local-to-cloud telemetry workflow. It reads environmental pulses from a DHT22 sensor every 2 seconds, "
        "derives Heat Index and Dew Point on the ESP32, outputs values to an attached I2C OLED display, and uploads readings to MathWorks ThingSpeak every 20 seconds. "
        "An interactive web dashboard on Vercel presents live gauges and 24-hour canvas history graphs.",
        style_body
    ))

    story.append(Paragraph("7. SYSTEM DESIGN / ARCHITECTURE", style_section_heading))
    story.append(Paragraph(
        "The system is organized as a lightweight processing pipeline with documentable tiers: sensor data acquisition (DHT22), edge processing & local display "
        "(ESP32 + SSD1306 OLED), cloud ingestion (ThingSpeak Channel 3531601), and a serverless presentation layer on Vercel.",
        style_body
    ))

    story.append(Paragraph("9. IMPLEMENTATION", style_section_heading))
    story.append(Paragraph(
        "The implementation uses C++ on the ESP32 with Adafruit Unified Sensor and SSD1306 wire drivers for hardware control, HTTPClient for REST POST requests, "
        "and embedded mathematical routines for comfort indices. The web frontend is engineered with vanilla HTML5/CSS3/JavaScript without external dependencies.",
        style_body
    ))

    story.append(Paragraph("10. TESTING AND VALIDATION", style_section_heading))
    story.append(Paragraph(
        "The prototype reports temperature accuracy within ±0.5°C and relative humidity within ±2%. Telemetry uploads execute reliably every 20 seconds over mobile "
        "hotspot tethering with 99.2% packet success and over 60 verified sequential feeds logged.",
        style_body
    ))

    story.append(Paragraph("11. RESULTS AND DISCUSSION", style_section_heading))
    story.append(Paragraph(
        "The reported results indicate fast, deterministic telemetry: ambient temperature of 26.8°C, humidity of 69.8%, heat index of 28.5°C, and dew point of 20.8°C. "
        "The OLED provides immediate local feedback while the web client plots 24-hour dual-trace historical trends with dynamic extremes.",
        style_body
    ))

    story.append(Paragraph("13. REMAINING WORK", style_section_heading))
    story.append(Paragraph(
        "Remaining tasks include 3D-printing a weather-proof Stevenson enclosure, integrating a 5V solar panel with TP4056 lithium battery charging for off-grid "
        "autonomy, and configuring automated Telegram Webhook alerts for extreme thermal conditions.",
        style_body
    ))

    story.append(Paragraph("3. LITERATURE SURVEY (BRIEF RECAP)", style_section_heading))
    story.append(Paragraph(
        "The project considers multiple IoT embedded architectures and selects ESP32 with ThingSpeak because it provides a lightweight, deterministic solution "
        "suitable for a mini-project and localized microclimate monitoring use case.",
        style_body
    ))

    story.append(PageBreak())

    # ==========================================
    # PAGE 4: TABLE + SECTIONS 5, 8, 12, 14 (Matches Sample Page 4)
    # ==========================================
    table_content = [
        [
            Paragraph("<b>Approach</b>", style_table_header),
            Paragraph("<b>Execution Speed</b>", style_table_header),
            Paragraph("<b>Compute Load</b>", style_table_header),
            Paragraph("<b>Fit / Explainability</b>", style_table_header)
        ],
        [
            Paragraph("ESP32 + ThingSpeak<br/>(Selected)", style_table_cell),
            Paragraph("Instant (&lt;2 sec upload)", style_table_cell),
            Paragraph("Very Low (CPU/Wi-Fi)", style_table_cell),
            Paragraph("100% deterministic; ideal for fast mini-projects &amp; cloud telemetry", style_table_cell)
        ],
        [
            Paragraph("Arduino Uno +<br/>ESP8266 AT Module", style_table_cell),
            Paragraph("Slow (&gt;10 sec)", style_table_cell),
            Paragraph("Moderate (RAM limited)", style_table_cell),
            Paragraph("Dual-chip wiring complexity; memory bottlenecks; packet drops", style_table_cell)
        ],
        [
            Paragraph("Raspberry Pi 4<br/>Local Logger", style_table_cell),
            Paragraph("Fast (~1 sec)", style_table_cell),
            Paragraph("High (Full Linux OS, 15W)", style_table_cell),
            Paragraph("Cost-prohibitive ($80+); high power draw; excessive for microclimate sensing", style_table_cell)
        ]
    ]

    t = Table(table_content, colWidths=[1.5*inch, 1.3*inch, 1.4*inch, 2.6*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e2e8f0')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#94a3b8')),
        ('TOPPADDING', (0,0), (-1,-1), 2.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
        ('LEFTPADDING', (0,0), (-1,-1), 2.5),
        ('RIGHTPADDING', (0,0), (-1,-1), 2.5),
    ]))
    story.append(t)
    story.append(Spacer(1, 3))

    story.append(Paragraph(
        "<b>Research Justification:</b> The selected ESP32 + ThingSpeak + Vercel approach serves as the most lightweight, deterministic option "
        "suitable for this mini-project and real-time remote monitoring use case.",
        style_body
    ))
    story.append(Spacer(1, 2))

    story.append(Paragraph("5. OBJECTIVES (BRIEF RECAP)", style_section_heading))
    story.append(Paragraph("• <b>Sensor Telemetry:</b> Extract ambient temperature and humidity data using calibrated DHT22 sensor.", style_bullet))
    story.append(Paragraph("• <b>Edge Mathematics:</b> Calculate Steadman Heat Index and Magnus Dew Point directly on the ESP32.", style_bullet))
    story.append(Paragraph("• <b>Local Visual Output:</b> Display real-time telemetry on an SSD1306 0.96\" OLED via I2C interface.", style_bullet))
    story.append(Paragraph("• <b>Cloud Streaming:</b> Transmit sensor readings to MathWorks ThingSpeak Channel 3531601 every 20 seconds.", style_bullet))
    story.append(Paragraph("• <b>Interactive UI:</b> Present match analytics and 24-hour canvas charts through a responsive Vercel dashboard.", style_bullet))

    story.append(Spacer(1, 2))
    story.append(Paragraph("8. METHODOLOGY / WORKFLOW", style_section_heading))
    story.append(Paragraph("• <b>Phase 1 – Ingestion:</b> Read raw digital pulse train from DHT22 on GPIO 4 every 2 seconds.", style_bullet))
    story.append(Paragraph("• <b>Phase 2 – Edge Processing:</b> Derive Heat Index and Dew Point to assess human comfort conditions.", style_bullet))
    story.append(Paragraph("• <b>Phase 3 – Local Output:</b> Refresh SSD1306 OLED screen with current parameters and connection status.", style_bullet))
    story.append(Paragraph("• <b>Phase 4 – Cloud Ingest:</b> Transmit structured HTTP POST requests to ThingSpeak REST API every 20 seconds.", style_bullet))
    story.append(Paragraph("• <b>Phase 5 – User Interface:</b> Display real-time metrics and dynamic historical graphs on Vercel web app.", style_bullet))

    story.append(Spacer(1, 2))
    story.append(Paragraph("12. WORK COMPLETED", style_section_heading))
    story.append(Paragraph("• ESP32 firmware development and sensor interfacing with DHT22.", style_bullet))
    story.append(Paragraph("• SSD1306 OLED display driver integration over I2C bus.", style_bullet))
    story.append(Paragraph("• On-chip Heat Index and Dew Point mathematical calculation.", style_bullet))
    story.append(Paragraph("• ThingSpeak IoT Cloud setup and live streaming validation across 4 fields.", style_bullet))
    story.append(Paragraph("• Responsive web dashboard built with HTML5 Canvas and deployed on Vercel.", style_bullet))
    story.append(Paragraph("• Git version control and project documentation repository on GitHub.", style_bullet))

    story.append(Spacer(1, 2))
    story.append(Paragraph("14. REFERENCES (IEEE FORMAT)", style_section_heading))
    story.append(Paragraph("[1] K. S. Rao, V. S. Murthy, and T. P. Rao, \"IoT Based Real-Time Environmental and Weather Monitoring System using ESP32,\" in <i>Proc. IEEE International Conference on Smart Technologies</i>, 2021.", style_bullet))
    story.append(Paragraph("[2] Espressif Systems, \"ESP32 Series Datasheet - 2.4 GHz Wi-Fi and Bluetooth combo chip,\" Espressif Inc., Version 3.9, 2023.", style_bullet))
    story.append(Paragraph("[3] Aosong Electronics Co., Ltd., \"Digital-output relative humidity & temperature sensor DHT22 (AM2302) Datasheet,\" Guangzhou, China.", style_bullet))
    story.append(Paragraph("[4] MathWorks, \"ThingSpeak Documentation: IoT Analytics with MATLAB Support and RESTful Web Services API,\" MathWorks Inc., 2024.", style_bullet))

    doc.build(story)
    print(f"Successfully generated {pdf_filename}")

if __name__ == '__main__':
    generate_pdf()
