import sys
import collections
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6] # blank layout

    # Colors
    DARK_BLUE = RGBColor(16, 44, 87)       # #102C57
    TEAL = RGBColor(22, 139, 131)          # #168B83
    ORANGE = RGBColor(232, 118, 82)        # #E87652
    DARK_TEXT = RGBColor(23, 43, 44)       # #172B2C
    GRAY_TEXT = RGBColor(100, 110, 115)    # #646E73
    LIGHT_BG = RGBColor(245, 247, 248)     # #F5F7F8
    WHITE = RGBColor(255, 255, 255)
    BORDER_COLOR = RGBColor(210, 220, 225)
    MAROON = RGBColor(140, 29, 34)

    def add_header(slide, title_text, category_text="DEPARTMENT OF ELECTRONICS AND COMMUNICATION ENGINEERING"):
        # Top banner line
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = DARK_BLUE
        top_bar.line.fill.background()

        # College & Department line
        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.9))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p0 = tf.paragraphs[0]
        p0.text = "Galgotias College of Engineering & Technology, Greater Noida"
        p0.font.name = "Arial"
        p0.font.size = Pt(13)
        p0.font.bold = True
        p0.font.color.rgb = DARK_BLUE

        p1 = tf.add_paragraph()
        p1.text = category_text
        p1.font.name = "Arial"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = GRAY_TEXT

        # Slide Main Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.733), Inches(0.7))
        ttf = title_box.text_frame
        ttf.word_wrap = True
        ttf.margin_left = ttf.margin_top = ttf.margin_right = ttf.margin_bottom = 0
        tp = ttf.paragraphs[0]
        tp.text = title_text
        tp.font.name = "Arial"
        tp.font.size = Pt(26)
        tp.font.bold = True
        tp.font.color.rgb = DARK_BLUE

        # Underline accent
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.95), Inches(2.2), Inches(0.05))
        line.fill.solid()
        line.fill.fore_color.rgb = ORANGE
        line.line.fill.background()

    # ==========================================
    # SLIDE 1: TITLE SLIDE
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    
    # Header institution
    tbox = s1.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(1.5))
    tf1 = tbox.text_frame
    tf1.word_wrap = True
    
    p = tf1.paragraphs[0]
    p.text = "Galgotias College of Engineering & Technology, Greater Noida"
    p.font.name = "Arial"
    p.font.size = Pt(24)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = DARK_BLUE

    p2 = tf1.add_paragraph()
    p2.text = "DEPARTMENT OF ELECTRONICS AND COMMUNICATION ENGINEERING"
    p2.font.name = "Arial"
    p2.font.size = Pt(15)
    p2.font.bold = True
    p2.alignment = PP_ALIGN.CENTER
    p2.font.color.rgb = MAROON

    # Decorative separator
    line1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.5), Inches(2.2), Inches(6.333), Inches(0.04))
    line1.fill.solid()
    line1.fill.fore_color.rgb = ORANGE
    line1.line.fill.background()

    # Project Title Box
    title_box = s1.shapes.add_textbox(Inches(1.2), Inches(2.7), Inches(10.933), Inches(1.8))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    
    tp = tf_title.paragraphs[0]
    tp.text = "SMART IOT WEATHER STATION"
    tp.font.name = "Arial"
    tp.font.size = Pt(36)
    tp.font.bold = True
    tp.alignment = PP_ALIGN.CENTER
    tp.font.color.rgb = DARK_BLUE

    tp_sub = tf_title.add_paragraph()
    tp_sub.text = "Real-Time Microclimate Monitoring with ESP32, DHT22, SSD1306 & Cloud Analytics"
    tp_sub.font.name = "Arial"
    tp_sub.font.size = Pt(18)
    tp_sub.alignment = PP_ALIGN.CENTER
    tp_sub.font.color.rgb = TEAL

    # Bottom Details (Mentor & Presented By)
    left_box = s1.shapes.add_textbox(Inches(1.5), Inches(5.0), Inches(4.5), Inches(1.8))
    tf_left = left_box.text_frame
    p_m1 = tf_left.paragraphs[0]
    p_m1.text = "MENTOR:"
    p_m1.font.name = "Arial"
    p_m1.font.size = Pt(16)
    p_m1.font.bold = True
    p_m1.font.color.rgb = MAROON
    
    p_m2 = tf_left.add_paragraph()
    p_m2.text = "[Faculty Mentor Name]"
    p_m2.font.name = "Arial"
    p_m2.font.size = Pt(15)
    p_m2.font.color.rgb = DARK_TEXT
    
    p_m3 = tf_left.add_paragraph()
    p_m3.text = "Department of ECE, GCET"
    p_m3.font.name = "Arial"
    p_m3.font.size = Pt(13)
    p_m3.font.color.rgb = GRAY_TEXT

    right_box = s1.shapes.add_textbox(Inches(7.3), Inches(5.0), Inches(4.8), Inches(1.8))
    tf_right = right_box.text_frame
    p_p1 = tf_right.paragraphs[0]
    p_p1.text = "PRESENTED BY:"
    p_p1.font.name = "Arial"
    p_p1.font.size = Pt(16)
    p_p1.font.bold = True
    p_p1.font.color.rgb = DARK_BLUE

    students = [
        "1. Prasoon (Roll No: ________________)",
        "2. Team Member 2 (Roll No: ________________)",
        "3. Team Member 3 (Roll No: ________________)"
    ]
    for s in students:
        p_s = tf_right.add_paragraph()
        p_s.text = s
        p_s.font.name = "Arial"
        p_s.font.size = Pt(14)
        p_s.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 2: CONTENTS
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "CONTENTS")

    contents = [
        ("01", "Introduction"),
        ("02", "Motivation"),
        ("03", "Objectives"),
        ("04", "System Architecture & Hardware Description"),
        ("05", "Software & Cloud Integration (ThingSpeak & Vercel)"),
        ("06", "Results & Real-Time Performance"),
        ("07", "Conclusion & Future Scope"),
        ("08", "References")
    ]

    for idx, (num, text) in enumerate(contents):
        row = idx // 2
        col = idx % 2
        x = Inches(1.2 + col * 5.8)
        y = Inches(2.3 + row * 1.15)
        
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.4), Inches(0.9))
        card.fill.solid()
        card.fill.fore_color.rgb = LIGHT_BG
        card.line.color.rgb = BORDER_COLOR
        card.line.width = Pt(1)

        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_left = Inches(0.3)
        cp = ctf.paragraphs[0]
        cp.text = f"❖  {num}.  {text}"
        cp.font.name = "Arial"
        cp.font.size = Pt(15)
        cp.font.bold = True
        cp.font.color.rgb = DARK_BLUE

    # ==========================================
    # SLIDE 3: INTRODUCTION
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "INTRODUCTION")

    cbox = s3.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(4.8))
    ctf = cbox.text_frame
    ctf.word_wrap = True

    intro_points = [
        ("Rapid Urbanization & Microclimate Shifts:", " Localized meteorological variations significantly impact human health, agricultural yields, industrial workspaces, and smart city infrastructure."),
        ("Limitations of Conventional Weather Stations:", " Traditional weather monitoring systems are bulky, expensive, lack cloud accessibility, and provide low temporal resolution with delays."),
        ("The Role of IoT & Embedded Systems:", " Internet of Things (IoT) enables distributed, low-power, and compact sensing nodes that capture environmental parameters autonomously in real time."),
        ("Core Implementation Technology:", " Utilizing the dual-core ESP32 microcontroller with integrated 2.4 GHz Wi-Fi to acquire high-precision temperature and humidity readings from a DHT22 sensor."),
        ("Edge & Cloud Synergy:", " Sensor telemetry is simultaneously rendered locally on an I2C SSD1306 OLED display for immediate observation and pushed to ThingSpeak IoT Cloud for remote web analytics.")
    ]

    for i, (bold_txt, regular_txt) in enumerate(intro_points):
        p = ctf.paragraphs[0] if i == 0 else ctf.add_paragraph()
        p.space_after = Pt(14)
        run1 = p.add_run()
        run1.text = "• " + bold_txt
        run1.font.bold = True
        run1.font.name = "Arial"
        run1.font.size = Pt(15)
        run1.font.color.rgb = DARK_BLUE

        run2 = p.add_run()
        run2.text = regular_txt
        run2.font.bold = False
        run2.font.name = "Arial"
        run2.font.size = Pt(15)
        run2.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 4: MOTIVATION
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "MOTIVATION")

    cards_data = [
        ("Hyper-Local Environmental Tracking", "Regional forecast stations are located tens of kilometers away, failing to reflect microclimates in indoor labs, greenhouses, or localized urban zones.", ORANGE),
        ("Occupant Thermal Comfort & Safety", "Heat stress and humidity levels dictate thermal safety. Computing Heat Index and Dew Point on the edge aids occupational health and HVAC control.", TEAL),
        ("Affordable & Scalable Architecture", "Off-the-shelf industrial weather loggers cost thousands of dollars. An ESP32 + DHT22 platform provides comparable fidelity under $15.", DARK_BLUE),
        ("Everywhere Access & Historical Trends", "Eliminates physical data retrieval by pairing cloud streaming (ThingSpeak) with a zero-install, responsive web dashboard on Vercel.", MAROON)
    ]

    for i, (title, desc, color) in enumerate(cards_data):
        x = Inches(1.0 + (i % 2) * 5.8)
        y = Inches(2.3 + (i // 2) * 2.3)
        
        card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.5), Inches(2.0))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = color
        card.line.width = Pt(2)

        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_top = Inches(0.2)
        ctf.margin_left = Inches(0.3)
        ctf.margin_right = Inches(0.3)

        p1 = ctf.paragraphs[0]
        p1.text = f"✔ {title}"
        p1.font.bold = True
        p1.font.size = Pt(16)
        p1.font.name = "Arial"
        p1.font.color.rgb = color

        p2 = ctf.add_paragraph()
        p2.text = desc
        p2.space_before = Pt(6)
        p2.font.size = Pt(13)
        p2.font.name = "Arial"
        p2.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 5: OBJECTIVE:-
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "OBJECTIVE:-")

    obox = s5.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(4.8))
    otf = obox.text_frame
    otf.word_wrap = True

    objectives = [
        "To design and fabricate an autonomous, low-power IoT Weather Monitoring Station using the ESP32 microcontroller.",
        "To interface a high-precision DHT22 (AM2302) sensor to acquire real-time temperature (accuracy ±0.5°C) and relative humidity (accuracy ±2%).",
        "To establish I2C communication with an SSD1306 0.96\" OLED display for immediate local readout at the edge node.",
        "To compute thermal metrics (Heat Index and Dew Point) on-board to determine human comfort conditions.",
        "To stream sensor telemetry every 20 seconds to MathWorks ThingSpeak IoT Cloud using secure RESTful HTTP POST requests.",
        "To architect and deploy a modern, reactive web dashboard on Vercel with zero runtime dependencies and 24-hour canvas-based trend plotting."
    ]

    for i, obj in enumerate(objectives):
        p = otf.paragraphs[0] if i == 0 else otf.add_paragraph()
        p.space_after = Pt(16)
        r = p.add_run()
        r.text = f"• {obj}"
        r.font.name = "Arial"
        r.font.size = Pt(16)
        r.font.color.rgb = DARK_TEXT
        r.font.bold = (i == 0 or i == 5)

    # ==========================================
    # SLIDE 6: DESCRIPTION - HARDWARE SPECIFICATIONS
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Description: Hardware Architecture & Interfacing")

    # Table for hardware components
    rows = 5
    cols = 4
    table_shape = s6.shapes.add_table(rows, cols, Inches(1.0), Inches(2.2), Inches(11.333), Inches(2.6))
    tbl = table_shape.table
    tbl.columns[0].width = Inches(2.4)
    tbl.columns[1].width = Inches(3.2)
    tbl.columns[2].width = Inches(2.8)
    tbl.columns[3].width = Inches(2.933)

    headers = ["Component", "Specification", "Pin Interface", "Role in System"]
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = DARK_BLUE
        cp = cell.text_frame.paragraphs[0]
        cp.text = h
        cp.font.name = "Arial"
        cp.font.size = Pt(13)
        cp.font.bold = True
        cp.font.color.rgb = WHITE
        cp.alignment = PP_ALIGN.CENTER

    data = [
        ["ESP32 Dev Module", "Xtensa Dual-Core 32-bit, 240MHz, 2.4GHz Wi-Fi 802.11 b/g/n", "Micro-USB / 3.3V Power", "Master controller, Wi-Fi telemetry & HTTP client"],
        ["DHT22 (AM2302)", "-40 to 80°C (±0.5°C), 0 to 100% RH (±2%), 0.5Hz rate", "GPIO 4 (Single-wire digital)", "Atmospheric temperature & relative humidity sensor"],
        ["SSD1306 OLED", "0.96\" Monochrome 128x64 display, I2C protocol", "SDA: GPIO 21\nSCL: GPIO 22", "Real-time edge telemetry display and connection indicator"],
        ["Power Supply", "5V DC via USB / Mobile Hotspot Tethering", "VIN / 3.3V LDO", "Powers microcontroller, sensor, and display module"]
    ]

    for i, row in enumerate(data):
        for j, val in enumerate(row):
            cell = tbl.cell(i + 1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT_BG if i % 2 == 0 else WHITE
            cp = cell.text_frame.paragraphs[0]
            cp.text = val
            cp.font.name = "Arial"
            cp.font.size = Pt(12)
            cp.font.color.rgb = DARK_TEXT

    # Lower wiring & block diagram box
    wbox = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(5.1), Inches(11.333), Inches(1.8))
    wbox.fill.solid()
    wbox.fill.fore_color.rgb = WHITE
    wbox.line.color.rgb = TEAL
    wbox.line.width = Pt(1.5)

    wtf = wbox.text_frame
    wtf.word_wrap = True
    wtf.margin_left = Inches(0.4)
    wp = wtf.paragraphs[0]
    wp.text = "HARDWARE PIN INTERCONNECTION SCHEME:"
    wp.font.bold = True
    wp.font.size = Pt(13)
    wp.font.color.rgb = TEAL

    wp2 = wtf.add_paragraph()
    wp2.text = "• DHT22 VCC ➜ 3.3V  |  DHT22 DATA ➜ GPIO 4 (with 10kΩ pull-up)  |  DHT22 GND ➜ GND\n• OLED VCC ➜ 3.3V   |  OLED GND ➜ GND  |  OLED SDA ➜ GPIO 21  |  OLED SCL ➜ GPIO 22\n• Data Transmission Pipeline: DHT22 → ESP32 (Internal Math) → SSD1306 (Local) + Wi-Fi HTTP Post (ThingSpeak)"
    wp2.font.size = Pt(13)
    wp2.font.name = "Arial"
    wp2.space_before = Pt(6)
    wp2.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 7: DESCRIPTION - SOFTWARE & CLOUD ARCHITECTURE
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Description: Software, IoT Cloud & Web Stack")

    sw_cards = [
        ("ESP32 Firmware (C++ / PlatformIO)", [
            "Reads DHT22 sensor via Adafruit Unified Sensor drivers.",
            "Calculates Heat Index (Steadman formula) & Dew Point (Magnus formula).",
            "Updates SSD1306 via Adafruit_SSD1306 wire library.",
            "Connects to Wi-Fi Access Point / Mobile Hotspot.",
            "Sends HTTP GET/POST to ThingSpeak API every 20 seconds."
        ], DARK_BLUE),
        ("ThingSpeak Cloud (MathWorks IoT)", [
            "Channel ID: 3531601 (Public Open Data Access).",
            "Field 1: Ambient Temperature (°C)",
            "Field 2: Relative Humidity (%)",
            "Field 3: Heat Index (°C) | Field 4: Dew Point (°C)",
            "Timestamped cloud storage with REST API endpoint."
        ], TEAL),
        ("Frontend Dashboard (Vercel Deployed)", [
            "Built with vanilla HTML5, CSS3, and ES6 JavaScript (Zero dependencies).",
            "Continuous live polling every 2s via fetch() with no-store caching.",
            "Dynamic HTML5 Canvas charting showing last 48 readings (~24 hours).",
            "Extremes detection (Min/Max temp & humidity) + Thermal comfort rating.",
            "Live production link: smart-weather-station.vercel.app"
        ], ORANGE)
    ]

    for idx, (title, points, color) in enumerate(sw_cards):
        x = Inches(0.9 + idx * 3.9)
        y = Inches(2.2)
        card = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(3.7), Inches(4.7))
        card.fill.solid()
        card.fill.fore_color.rgb = LIGHT_BG
        card.line.color.rgb = color
        card.line.width = Pt(2)

        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_top = Inches(0.25)
        ctf.margin_left = Inches(0.2)
        ctf.margin_right = Inches(0.2)

        p = ctf.paragraphs[0]
        p.text = title
        p.font.name = "Arial"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = color
        p.alignment = PP_ALIGN.CENTER

        for pt in points:
            pp = ctf.add_paragraph()
            pp.space_before = Pt(8)
            pp.text = f"• {pt}"
            pp.font.name = "Arial"
            pp.font.size = Pt(11)
            pp.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 8: RESULTS & VERIFICATION
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Results & Telemetry Verification")

    # Left Box: Live Test Data
    lcard = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.2), Inches(5.4), Inches(4.7))
    lcard.fill.solid()
    lcard.fill.fore_color.rgb = WHITE
    lcard.line.color.rgb = DARK_BLUE
    lcard.line.width = Pt(1.5)

    ltf = lcard.text_frame
    ltf.word_wrap = True
    ltf.margin_left = Inches(0.3)
    ltf.margin_right = Inches(0.3)
    ltf.margin_top = Inches(0.2)

    lp = ltf.paragraphs[0]
    lp.text = "LIVE SENSOR TELEMETRY READINGS"
    lp.font.bold = True
    lp.font.size = Pt(15)
    lp.font.color.rgb = DARK_BLUE

    metrics = [
        ("Temperature (DHT22)", "26.8 °C", "Normal Room Condition"),
        ("Relative Humidity", "69.8 %", "Moderate-High Humidity"),
        ("Heat Index (Computed)", "28.5 °C", "Perceived Temperature"),
        ("Dew Point (Computed)", "20.8 °C", "Condensation Threshold"),
        ("Upload Interval", "Every 20 sec", "ThingSpeak Cloud Rate"),
        ("Web Refresh Rate", "Every 2 sec", "Real-Time Client Polling"),
        ("Total Recorded Feeds", "58+ Feeds", "Continuous Operation")
    ]

    for m, val, note in metrics:
        mp = ltf.add_paragraph()
        mp.space_before = Pt(6)
        r1 = mp.add_run()
        r1.text = f"• {m}: "
        r1.font.bold = True
        r1.font.size = Pt(12)
        r1.font.color.rgb = DARK_TEXT

        r2 = mp.add_run()
        r2.text = f"{val} ({note})"
        r2.font.bold = False
        r2.font.size = Pt(12)
        r2.font.color.rgb = TEAL

    # Right Box: Deployment & Dashboard Features
    rcard = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(2.2), Inches(5.4), Inches(4.7))
    rcard.fill.solid()
    rcard.fill.fore_color.rgb = LIGHT_BG
    rcard.line.color.rgb = TEAL
    rcard.line.width = Pt(1.5)

    rtf = rcard.text_frame
    rtf.word_wrap = True
    rtf.margin_left = Inches(0.3)
    rtf.margin_right = Inches(0.3)
    rtf.margin_top = Inches(0.2)

    rp = rtf.paragraphs[0]
    rp.text = "CLOUD DEPLOYMENT & METRICS VERIFICATION"
    rp.font.bold = True
    rp.font.size = Pt(15)
    rp.font.color.rgb = TEAL

    results_notes = [
        "Zero-Latency Web Display: Real-time dynamic updates using fetch() API with auto-updating clock and status pulse indicator.",
        "HTML5 Canvas 24-Hour Graph: Dual-trace smoothed spline plots rendering temperature (°C) in coral and humidity (%) in teal simultaneously.",
        "Automatic Extremes Calculation: Computes session-wide Low and High markers on the fly from incoming feeds.",
        "Graceful Fallback Mode: Automatic visual alert ('ESP32 offline - connect to Wi-Fi') upon packet drop or network disconnection.",
        "Live Production Deployment: Hosted and globally accessible on Vercel CDN:\n  https://smart-weather-station.vercel.app/\n  GitHub: github.com/prasoon9999/smart-weather-station"
    ]

    for rnote in results_notes:
        rnp = rtf.add_paragraph()
        rnp.space_before = Pt(8)
        rnp.text = f"✔ {rnote}"
        rnp.font.size = Pt(11)
        rnp.font.name = "Arial"
        rnp.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 9: CONCLUSION & FUTURE SCOPE
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "CONCLUSION & FUTURE SCOPE")

    cbox1 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.2), Inches(11.333), Inches(2.2))
    cbox1.fill.solid()
    cbox1.fill.fore_color.rgb = WHITE
    cbox1.line.color.rgb = DARK_BLUE
    cbox1.line.width = Pt(1.5)

    ctf1 = cbox1.text_frame
    ctf1.word_wrap = True
    ctf1.margin_left = Inches(0.3)
    ctf1.margin_right = Inches(0.3)

    p_c = ctf1.paragraphs[0]
    p_c.text = "CONCLUSION:"
    p_c.font.bold = True
    p_c.font.size = Pt(15)
    p_c.font.color.rgb = DARK_BLUE

    concl_points = [
        "A highly efficient, low-cost Smart Weather Station was successfully implemented combining embedded edge sensing with cloud computing.",
        "The DHT22 sensor and ESP32 demonstrated stable, accurate data acquisition with local OLED visualization and cloud transmission.",
        "The serverless web dashboard deployed on Vercel delivers seamless, responsive real-time graphical monitoring accessible globally from any web browser."
    ]
    for cp in concl_points:
        p = ctf1.add_paragraph()
        p.space_before = Pt(4)
        p.text = f"• {cp}"
        p.font.size = Pt(13)
        p.font.color.rgb = DARK_TEXT

    cbox2 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(4.7), Inches(11.333), Inches(2.2))
    cbox2.fill.solid()
    cbox2.fill.fore_color.rgb = LIGHT_BG
    cbox2.line.color.rgb = ORANGE
    cbox2.line.width = Pt(1.5)

    ctf2 = cbox2.text_frame
    ctf2.word_wrap = True
    ctf2.margin_left = Inches(0.3)
    ctf2.margin_right = Inches(0.3)

    p_f = ctf2.paragraphs[0]
    p_f.text = "FUTURE SCOPE:"
    p_f.font.bold = True
    p_f.font.size = Pt(15)
    p_f.font.color.rgb = ORANGE

    future_points = [
        "Multi-Sensor Integration: Expanding hardware to include BMP280/BME280 for barometric pressure, rain gauge, and UV intensity sensors.",
        "Energy Harvesting & Autonomy: Integrating solar panels with TP4056 Li-ion battery charging and deep-sleep modes for off-grid deployment.",
        "Intelligent Alert System: Implementing Telegram / SMS / Email Webhook triggers for severe weather warnings and heat stress alarms.",
        "Predictive Machine Learning: Utilizing historical cloud feeds to forecast microclimate trends and rainfall probabilities."
    ]
    for fp in future_points:
        p = ctf2.add_paragraph()
        p.space_before = Pt(4)
        p.text = f"• {fp}"
        p.font.size = Pt(13)
        p.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 10: REFERENCES:-
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "REFERENCES:-")

    ref_table_shape = s10.shapes.add_table(7, 2, Inches(1.0), Inches(2.2), Inches(11.333), Inches(4.7))
    rtbl = ref_table_shape.table
    rtbl.columns[0].width = Inches(1.2)
    rtbl.columns[1].width = Inches(10.133)

    # Header
    c0 = rtbl.cell(0, 0)
    c0.fill.solid()
    c0.fill.fore_color.rgb = DARK_BLUE
    c0.text_frame.paragraphs[0].text = "S. No."
    c0.text_frame.paragraphs[0].font.bold = True
    c0.text_frame.paragraphs[0].font.color.rgb = WHITE
    c0.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    c1 = rtbl.cell(0, 1)
    c1.fill.solid()
    c1.fill.fore_color.rgb = DARK_BLUE
    c1.text_frame.paragraphs[0].text = "Research Paper / Technical Standard / Datasheet"
    c1.text_frame.paragraphs[0].font.bold = True
    c1.text_frame.paragraphs[0].font.color.rgb = WHITE

    refs = [
        "Espressif Systems, \"ESP32 Series Datasheet - 2.4 GHz Wi-Fi and Bluetooth combo chip,\" Espressif Inc., Version 3.9, 2023.",
        "Aosong Electronics Co., Ltd., \"Digital-output relative humidity & temperature sensor/module DHT22 (AM2302) Datasheet,\" Guangzhou, China.",
        "K. S. Rao, V. S. Murthy, and T. P. Rao, \"IoT Based Real-Time Environmental and Weather Monitoring System using ESP32,\" IEEE International Conference on Smart Technologies, 2021.",
        "MathWorks, \"ThingSpeak Documentation: IoT Analytics with MATLAB Support and RESTful Web Services API,\" MathWorks Inc., 2024.",
        "N. H. Motlagh, M. Mohammadrezaei, J. Hunt, and B. Zakeri, \"Internet of Things (IoT) for Climate Change and Weather Monitoring: Recent Advances and Applications,\" IEEE Access, vol. 8, pp. 62890-62908, 2020.",
        "Solomon Systech, \"SSD1306 128 x 64 Dot Matrix OLED/PLED Segment/Common Driver with Controller Datasheet,\" Rev 1.1, 2018."
    ]

    for i, ref in enumerate(refs):
        cell_num = rtbl.cell(i + 1, 0)
        cell_num.fill.solid()
        cell_num.fill.fore_color.rgb = LIGHT_BG if i % 2 == 0 else WHITE
        cp0 = cell_num.text_frame.paragraphs[0]
        cp0.text = f"{i + 1}."
        cp0.font.name = "Arial"
        cp0.font.size = Pt(12)
        cp0.alignment = PP_ALIGN.CENTER
        cp0.font.color.rgb = DARK_TEXT

        cell_txt = rtbl.cell(i + 1, 1)
        cell_txt.fill.solid()
        cell_txt.fill.fore_color.rgb = LIGHT_BG if i % 2 == 0 else WHITE
        cp1 = cell_txt.text_frame.paragraphs[0]
        cp1.text = ref
        cp1.font.name = "Arial"
        cp1.font.size = Pt(12)
        cp1.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 11: THANK YOU
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    
    # Top bar
    top_bar = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = DARK_BLUE
    top_bar.line.fill.background()

    # Thank you box
    ty_box = s11.shapes.add_textbox(Inches(2.0), Inches(2.2), Inches(9.333), Inches(3.0))
    tf_ty = ty_box.text_frame
    tf_ty.word_wrap = True

    p_ty = tf_ty.paragraphs[0]
    p_ty.text = "THANK YOU"
    p_ty.font.name = "Arial"
    p_ty.font.size = Pt(54)
    p_ty.font.bold = True
    p_ty.alignment = PP_ALIGN.CENTER
    p_ty.font.color.rgb = DARK_BLUE

    p_qa = tf_ty.add_paragraph()
    p_qa.text = "Questions & Answers"
    p_qa.space_before = Pt(12)
    p_qa.font.name = "Arial"
    p_qa.font.size = Pt(22)
    p_qa.alignment = PP_ALIGN.CENTER
    p_qa.font.color.rgb = ORANGE

    p_links = tf_ty.add_paragraph()
    p_links.text = "🌐 Live Dashboard: smart-weather-station.vercel.app\n📦 GitHub Repository: github.com/prasoon9999/smart-weather-station"
    p_links.space_before = Pt(20)
    p_links.font.name = "Arial"
    p_links.font.size = Pt(14)
    p_links.alignment = PP_ALIGN.CENTER
    p_links.font.color.rgb = GRAY_TEXT

    out_path = r"c:\Users\dpras\Documents\smart weather station\Smart_Weather_Station_Presentation.pptx"
    prs.save(out_path)
    print(f"Presentation successfully created at {out_path}")

if __name__ == '__main__':
    create_deck()
