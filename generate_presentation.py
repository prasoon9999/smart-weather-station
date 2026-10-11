import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_aesthetic_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- LUXURY / MODERN EDITORIAL COLOR PALETTE ---
    # Matching the weather dashboard design: Dark Slate & Warm Copper/Coral with Teal Glow
    BG_DARK = RGBColor(15, 23, 42)          # Deep Navy/Slate #0F172A
    CARD_DARK = RGBColor(24, 34, 58)        # Elevated Surface #18223A
    CARD_BORDER = RGBColor(45, 62, 98)      # Subtle Slate Border #2D3E62
    
    TEXT_LIGHT = RGBColor(248, 250, 252)    # Crisp Off-White #F8FAFC
    TEXT_MUTED = RGBColor(148, 163, 184)    # Cool Slate Gray #94A3B8
    
    ACCENT_TEAL = RGBColor(20, 184, 166)    # Mint / Teal #14B8A6
    ACCENT_CORAL = RGBColor(249, 115, 22)   # Vibrant Coral / Sunset #F97316
    ACCENT_GOLD = RGBColor(245, 158, 11)    # Golden Glow #F59E0B
    ACCENT_CYAN = RGBColor(56, 189, 248)    # Neon Cyan #38BDF8
    WHITE = RGBColor(255, 255, 255)

    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_DARK
        bg.line.fill.background()
        return bg

    def add_top_navigation(slide, category, title, slide_num):
        # Subtle top accent bar
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.08))
        bar.fill.solid()
        bar.fill.fore_color.rgb = ACCENT_TEAL
        bar.line.fill.background()

        # College & Category Eyebrow Pill
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(0.35), Inches(7.5), Inches(0.38))
        pill.fill.solid()
        pill.fill.fore_color.rgb = CARD_DARK
        pill.line.color.rgb = CARD_BORDER
        pill.line.width = Pt(1)
        ptf = pill.text_frame
        ptf.margin_top = Inches(0.06)
        ptf.margin_left = Inches(0.2)
        pp = ptf.paragraphs[0]
        pp.text = f"GALGOTIAS COLLEGE OF ENGG. & TECH.  •  ECE DEPARTMENT  •  {category.upper()}"
        pp.font.name = "Segoe UI"
        pp.font.size = Pt(8.5)
        pp.font.bold = True
        pp.font.color.rgb = ACCENT_TEAL

        # Slide Number Badge
        nbadge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.6), Inches(0.35), Inches(0.85), Inches(0.38))
        nbadge.fill.solid()
        nbadge.fill.fore_color.rgb = CARD_DARK
        nbadge.line.color.rgb = CARD_BORDER
        nbadge.line.width = Pt(1)
        np = nbadge.text_frame.paragraphs[0]
        np.text = f"{slide_num:02d} / 10"
        np.font.name = "Segoe UI"
        np.font.size = Pt(9)
        np.font.bold = True
        np.font.color.rgb = TEXT_MUTED
        np.alignment = PP_ALIGN.CENTER

        # Slide Title
        tbox = slide.shapes.add_textbox(Inches(0.9), Inches(0.85), Inches(11.5), Inches(0.75))
        tf = tbox.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Segoe UI"
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = TEXT_LIGHT

    # ==========================================
    # SLIDE 1: AESTHETIC HERO TITLE
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)

    # Ambient Glow card in center
    glow = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(0.8), Inches(10.933), Inches(5.9))
    glow.fill.solid()
    glow.fill.fore_color.rgb = CARD_DARK
    glow.line.color.rgb = CARD_BORDER
    glow.line.width = Pt(1.5)

    # Eyebrow tag
    tag = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.3), Inches(1.3), Inches(4.7), Inches(0.42))
    tag.fill.solid()
    tag.fill.fore_color.rgb = RGBColor(19, 78, 74) # Dark emerald
    tag.line.color.rgb = ACCENT_TEAL
    tag.line.width = Pt(1)
    tag_p = tag.text_frame.paragraphs[0]
    tag_p.text = "⚡ ESP32  •  DHT22  •  THINGSPEAK  •  VERCEL"
    tag_p.font.name = "Segoe UI"
    tag_p.font.size = Pt(10)
    tag_p.font.bold = True
    tag_p.font.color.rgb = ACCENT_TEAL
    tag_p.alignment = PP_ALIGN.CENTER

    # Main Big Title
    title_box = s1.shapes.add_textbox(Inches(1.8), Inches(1.95), Inches(9.733), Inches(1.5))
    ttf = title_box.text_frame
    ttf.word_wrap = True
    p1 = ttf.paragraphs[0]
    p1.text = "Smart Weather Station"
    p1.font.name = "Segoe UI"
    p1.font.size = Pt(44)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_LIGHT
    p1.alignment = PP_ALIGN.CENTER

    p2 = ttf.add_paragraph()
    p2.text = "Autonomous Microclimate Telemetry & Real-Time Cloud Analytics"
    p2.font.name = "Segoe UI"
    p2.font.size = Pt(18)
    p2.space_before = Pt(8)
    p2.font.color.rgb = ACCENT_CYAN
    p2.alignment = PP_ALIGN.CENTER

    # Department & Institution
    dept_box = s1.shapes.add_textbox(Inches(1.8), Inches(3.6), Inches(9.733), Inches(0.6))
    dtf = dept_box.text_frame
    dp = dtf.paragraphs[0]
    dp.text = "DEPARTMENT OF ELECTRONICS AND COMMUNICATION ENGINEERING\nGalgotias College of Engineering & Technology, Greater Noida"
    dp.font.name = "Segoe UI"
    dp.font.size = Pt(11)
    dp.font.bold = True
    dp.font.color.rgb = TEXT_MUTED
    dp.alignment = PP_ALIGN.CENTER

    # Bottom Cards for Mentor and Students
    m_card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.0), Inches(4.5), Inches(4.3), Inches(1.7))
    m_card.fill.solid()
    m_card.fill.fore_color.rgb = BG_DARK
    m_card.line.color.rgb = CARD_BORDER
    m_card.line.width = Pt(1)
    mtf = m_card.text_frame
    mtf.margin_top = Inches(0.2)
    mtf.margin_left = Inches(0.3)
    mp1 = mtf.paragraphs[0]
    mp1.text = "PROJECT MENTOR"
    mp1.font.name = "Segoe UI"
    mp1.font.size = Pt(10)
    mp1.font.bold = True
    mp1.font.color.rgb = ACCENT_CORAL
    mp2 = mtf.add_paragraph()
    mp2.text = "[Faculty Mentor Name]"
    mp2.font.name = "Segoe UI"
    mp2.font.size = Pt(14)
    mp2.font.bold = True
    mp2.space_before = Pt(4)
    mp2.font.color.rgb = TEXT_LIGHT
    mp3 = mtf.add_paragraph()
    mp3.text = "Assistant Professor, Dept of ECE"
    mp3.font.name = "Segoe UI"
    mp3.font.size = Pt(11)
    mp3.font.color.rgb = TEXT_MUTED

    p_card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.0), Inches(4.5), Inches(4.3), Inches(1.7))
    p_card.fill.solid()
    p_card.fill.fore_color.rgb = BG_DARK
    p_card.line.color.rgb = CARD_BORDER
    p_card.line.width = Pt(1)
    ptf = p_card.text_frame
    ptf.margin_top = Inches(0.2)
    ptf.margin_left = Inches(0.3)
    pp1 = ptf.paragraphs[0]
    pp1.text = "PRESENTED BY"
    pp1.font.name = "Segoe UI"
    pp1.font.size = Pt(10)
    pp1.font.bold = True
    pp1.font.color.rgb = ACCENT_TEAL
    pp2 = ptf.add_paragraph()
    pp2.text = "1. Prasoon (Roll No: _____________)\n2. Member 2 (Roll No: _____________)\n3. Member 3 (Roll No: _____________)"
    pp2.font.name = "Segoe UI"
    pp2.font.size = Pt(11)
    pp2.space_before = Pt(4)
    pp2.font.color.rgb = TEXT_LIGHT

    # ==========================================
    # SLIDE 2: MODERN CONTENTS GRID
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_top_navigation(s2, "Overview", "Table of Contents", 2)

    contents = [
        ("01", "Introduction & Context", "Environmental importance, microclimates, and real-time monitoring challenges.", ACCENT_TEAL),
        ("02", "Project Motivation", "Why conventional loggers fall short and why local human thermal comfort matters.", ACCENT_CORAL),
        ("03", "System Objectives", "End-to-end design goals: edge acquisition, math computation, and web telemetry.", ACCENT_CYAN),
        ("04", "Hardware Architecture", "ESP32 dual-core MCU, DHT22 capacitive sensor, SSD1306 I2C OLED display.", ACCENT_GOLD),
        ("05", "Software & Cloud Stack", "PlatformIO C++ firmware, MathWorks ThingSpeak API, and Vercel serverless host.", ACCENT_TEAL),
        ("06", "Results & Live Metrics", "Real-world test data, 24-hr canvas history graph, and extreme temperature tracking.", ACCENT_CORAL),
        ("07", "Conclusion & Future Scope", "System accomplishments, solar battery harvesting, and multi-sensor expansion.", ACCENT_CYAN),
        ("08", "References & Standards", "IEEE research papers, sensor datasheets, and official technical references.", ACCENT_GOLD),
    ]

    for idx, (num, heading, sub, col) in enumerate(contents):
        row = idx // 4
        col_idx = idx % 4
        x = Inches(0.9 + col_idx * 2.95)
        y = Inches(1.8 + row * 2.65)
        w = Inches(2.75)
        h = Inches(2.4)

        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_DARK
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1)

        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_top = Inches(0.2)
        ctf.margin_left = Inches(0.2)
        ctf.margin_right = Inches(0.2)

        # Number pill
        p0 = ctf.paragraphs[0]
        p0.text = num
        p0.font.name = "Segoe UI"
        p0.font.size = Pt(20)
        p0.font.bold = True
        p0.font.color.rgb = col

        p1 = ctf.add_paragraph()
        p1.text = heading
        p1.font.name = "Segoe UI"
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.space_before = Pt(8)
        p1.font.color.rgb = TEXT_LIGHT

        p2 = ctf.add_paragraph()
        p2.text = sub
        p2.font.name = "Segoe UI"
        p2.font.size = Pt(10)
        p2.space_before = Pt(6)
        p2.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 3: INTRODUCTION (SLEEK TIMELINE / PILLARS)
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_top_navigation(s3, "Background", "Introduction & Problem Statement", 3)

    intro_cards = [
        ("01", "Localized Microclimates", "Macro weather stations are located tens of kilometers apart at airports, failing to reflect microclimates inside urban corridors, agricultural greenhouses, server rooms, and academic campuses.", ACCENT_CORAL),
        ("02", "Legacy System Bottlenecks", "Conventional meteorological systems are bulky, cost-prohibitive ($1,000+), rely on manual data extraction, and lack unified instant cloud access for end-users.", ACCENT_GOLD),
        ("03", "IoT & Edge Intelligence", "Low-power microcontrollers like the ESP32 enable distributed sensing nodes with on-board edge computation (heat index, dew point) and Wi-Fi transmission.", ACCENT_TEAL),
        ("04", "Dual Telemetry Paradigm", "Local OLED display guarantees zero-latency on-site visual diagnostics even when offline, while ThingSpeak & Vercel deliver global, browser-based access.", ACCENT_CYAN)
    ]

    for i, (num, h, desc, col) in enumerate(intro_cards):
        x = Inches(0.9 + i * 2.95)
        y = Inches(1.8)
        w = Inches(2.75)
        h_box = Inches(5.1)

        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h_box)
        c.fill.solid()
        c.fill.fore_color.rgb = CARD_DARK
        c.line.color.rgb = CARD_BORDER
        c.line.width = Pt(1)

        # Top tag line
        top_line = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Inches(0.2), y, Inches(1.0), Inches(0.06))
        top_line.fill.solid()
        top_line.fill.fore_color.rgb = col
        top_line.line.fill.background()

        ctf = c.text_frame
        ctf.word_wrap = True
        ctf.margin_top = Inches(0.3)
        ctf.margin_left = Inches(0.2)
        ctf.margin_right = Inches(0.2)

        p = ctf.paragraphs[0]
        p.text = f"PILLAR {num}"
        p.font.name = "Segoe UI"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = col

        p_h = ctf.add_paragraph()
        p_h.text = h
        p_h.font.name = "Segoe UI"
        p_h.font.size = Pt(15)
        p_h.font.bold = True
        p_h.space_before = Pt(8)
        p_h.font.color.rgb = TEXT_LIGHT

        p_d = ctf.add_paragraph()
        p_d.text = desc
        p_d.font.name = "Segoe UI"
        p_d.font.size = Pt(11.5)
        p_d.space_before = Pt(14)
        p_d.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 4: OBJECTIVES (CARD GRID WITH GLOW)
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_top_navigation(s4, "Project Scope", "System Design Objectives", 4)

    objectives = [
        ("Autonomous Edge Sensing", "Interfacing ESP32 with DHT22 (AM2302) to sample temperature (-40 to 80°C) and relative humidity (0-100% RH) with high stability.", ACCENT_TEAL),
        ("On-Device Comfort Index Math", "Implementing Steadman's Heat Index and Magnus Dew Point equations in firmware to detect heat exhaustion and condensation hazards.", ACCENT_CORAL),
        ("Local OLED Feedback", "Driving a 0.96\" 128x64 SSD1306 display over I2C bus at 400kHz to deliver crisp, immediate readings directly at the physical station.", ACCENT_CYAN),
        ("Secure IoT Cloud Ingestion", "Packaging 4-field payload (Temp, Humidity, Heat Index, Dew Point) and streaming to MathWorks ThingSpeak every 20 seconds.", ACCENT_GOLD),
        ("Modern Zero-Dependency Dashboard", "Architecting a lightweight web interface with custom HTML5 Canvas dual-trace historical charts that runs smoothly on both desktop and mobile.", ACCENT_TEAL),
        ("Serverless Cloud Deployment", "Deploying the dashboard continuously via Vercel edge network and open-sourcing the firmware and frontend on GitHub.", ACCENT_CORAL)
    ]

    for i, (title, desc, col) in enumerate(objectives):
        row = i // 3
        col_idx = i % 3
        x = Inches(0.9 + col_idx * 3.95)
        y = Inches(1.8 + row * 2.6)
        w = Inches(3.75)
        h = Inches(2.35)

        card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_DARK
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1)

        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_top = Inches(0.2)
        ctf.margin_left = Inches(0.25)
        ctf.margin_right = Inches(0.25)

        p = ctf.paragraphs[0]
        p.text = f"✔  {title}"
        p.font.name = "Segoe UI"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = col

        p2 = ctf.add_paragraph()
        p2.text = desc
        p2.font.name = "Segoe UI"
        p2.font.size = Pt(11)
        p2.space_before = Pt(8)
        p2.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 5: HARDWARE SPECIFICATIONS & PINOUT
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_top_navigation(s5, "Architecture", "Hardware Engineering & Pinout Scheme", 5)

    # 3 Component Highlight Cards on Top
    hw_cards = [
        ("ESP32 Dev Module", "Xtensa Dual-Core 32-bit @ 240MHz\n• Integrated 2.4 GHz Wi-Fi (802.11 b/g/n)\n• Low-power deep sleep support\n• 520 KB SRAM + 4MB Flash Memory", "CORE PROCESSOR", ACCENT_TEAL),
        ("DHT22 Sensor (AM2302)", "Capacitive Humidity & Thermistor\n• Temp Range: -40 to +80°C (±0.5°C)\n• Humidity Range: 0 to 100% RH (±2%)\n• Digital 1-wire communication (GPIO 4)", "ENVIRONMENTAL SENSOR", ACCENT_CORAL),
        ("SSD1306 0.96\" OLED", "Monochrome 128x64 Graphic Display\n• High contrast & wide viewing angle\n• I2C Interface: SDA (GPIO 21), SCL (GPIO 22)\n• Low power consumption (0.08W)", "EDGE DISPLAY", ACCENT_CYAN)
    ]

    for i, (title, specs, role, col) in enumerate(hw_cards):
        x = Inches(0.9 + i * 3.95)
        y = Inches(1.8)
        w = Inches(3.75)
        h = Inches(3.0)

        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
        c.fill.solid()
        c.fill.fore_color.rgb = CARD_DARK
        c.line.color.rgb = CARD_BORDER
        c.line.width = Pt(1)

        ctf = c.text_frame
        ctf.word_wrap = True
        ctf.margin_top = Inches(0.2)
        ctf.margin_left = Inches(0.25)
        ctf.margin_right = Inches(0.25)

        p = ctf.paragraphs[0]
        p.text = role
        p.font.name = "Segoe UI"
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = col

        p2 = ctf.add_paragraph()
        p2.text = title
        p2.font.name = "Segoe UI"
        p2.font.size = Pt(16)
        p2.font.bold = True
        p2.space_before = Pt(4)
        p2.font.color.rgb = TEXT_LIGHT

        p3 = ctf.add_paragraph()
        p3.text = specs
        p3.font.name = "Segoe UI"
        p3.font.size = Pt(11)
        p3.space_before = Pt(8)
        p3.font.color.rgb = TEXT_MUTED

    # Bottom Full-Width Wiring Diagram Card
    wiring_card = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(5.0), Inches(11.533), Inches(1.9))
    wiring_card.fill.solid()
    wiring_card.fill.fore_color.rgb = CARD_DARK
    wiring_card.line.color.rgb = ACCENT_TEAL
    wiring_card.line.width = Pt(1.5)

    wt = wiring_card.text_frame
    wt.word_wrap = True
    wt.margin_top = Inches(0.2)
    wt.margin_left = Inches(0.3)
    wt.margin_right = Inches(0.3)

    wp1 = wt.paragraphs[0]
    wp1.text = "⚡ SCHEMATIC PIN MAPPING & POWER DISTRIBUTION"
    wp1.font.name = "Segoe UI"
    wp1.font.size = Pt(12)
    wp1.font.bold = True
    wp1.font.color.rgb = ACCENT_TEAL

    wp2 = wt.add_paragraph()
    wp2.text = (
        "• DHT22 Sensor ➜ VCC to 3.3V Pin  |  GND to Common Ground  |  Data OUT to GPIO 4 (10kΩ pull-up resistor to 3.3V)\n"
        "• SSD1306 Display ➜ VCC to 3.3V Pin  |  GND to Common Ground  |  SDA to GPIO 21  |  SCL to GPIO 22\n"
        "• Power & Tethering ➜ 5V DC via Micro-USB cable supplied by power bank or mobile hotspot connection for live uplink."
    )
    wp2.font.name = "Segoe UI"
    wp2.font.size = Pt(11)
    wp2.space_before = Pt(6)
    wp2.font.color.rgb = TEXT_LIGHT

    # ==========================================
    # SLIDE 6: SOFTWARE & CLOUD PIPELINE
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_top_navigation(s6, "Telemetry Flow", "Software Architecture & Cloud Data Pipeline", 6)

    flow_steps = [
        ("STEP 1: SENSING", "DHT22 Sensor", "Measures temperature & relative humidity every 2s via single-wire digital pulse width modulation.", ACCENT_TEAL),
        ("STEP 2: EDGE MATH", "ESP32 Firmware", "Calculates Heat Index & Dew Point on-chip. Refreshes SSD1306 OLED via I2C at 400kHz.", ACCENT_CYAN),
        ("STEP 3: CLOUD INGEST", "ThingSpeak IoT", "HTTP GET/POST pushes 4 data fields every 20s to Channel 3531601. Stores historical feeds.", ACCENT_GOLD),
        ("STEP 4: WEB DASHBOARD", "Vercel CDN", "Static dashboard polls ThingSpeak every 2s, plots 24-hr canvas graph, computes extremes.", ACCENT_CORAL)
    ]

    for i, (step, name, desc, col) in enumerate(flow_steps):
        x = Inches(0.9 + i * 2.95)
        y = Inches(1.8)
        w = Inches(2.75)
        h = Inches(5.1)

        card = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_DARK
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1)

        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_top = Inches(0.25)
        ctf.margin_left = Inches(0.2)
        ctf.margin_right = Inches(0.2)

        p = ctf.paragraphs[0]
        p.text = step
        p.font.name = "Segoe UI"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = col

        p2 = ctf.add_paragraph()
        p2.text = name
        p2.font.name = "Segoe UI"
        p2.font.size = Pt(16)
        p2.font.bold = True
        p2.space_before = Pt(6)
        p2.font.color.rgb = TEXT_LIGHT

        p3 = ctf.add_paragraph()
        p3.text = desc
        p3.font.name = "Segoe UI"
        p3.font.size = Pt(11)
        p3.space_before = Pt(12)
        p3.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 7: RESULTS & LIVE TELEMETRY VERIFICATION
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_top_navigation(s7, "Field Verification", "Experimental Results & Real-Time Performance", 7)

    # 4 Large Metric Stat Cards on Left
    stats = [
        ("TEMPERATURE", "26.8°C", "DHT22 Calibrated Reading", ACCENT_CORAL),
        ("HUMIDITY", "69.8%", "Relative Moisture Level", ACCENT_TEAL),
        ("HEAT INDEX", "28.5°C", "Human Felt-Temperature", ACCENT_GOLD),
        ("DEW POINT", "20.8°C", "Condensation Threshold", ACCENT_CYAN)
    ]

    for i, (label, val, note, col) in enumerate(stats):
        row = i // 2
        col_idx = i % 2
        x = Inches(0.9 + col_idx * 3.0)
        y = Inches(1.8 + row * 2.5)
        w = Inches(2.8)
        h = Inches(2.25)

        card = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_DARK
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1)

        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_top = Inches(0.2)
        ctf.margin_left = Inches(0.25)

        p = ctf.paragraphs[0]
        p.text = label
        p.font.name = "Segoe UI"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = TEXT_MUTED

        p2 = ctf.add_paragraph()
        p2.text = val
        p2.font.name = "Segoe UI"
        p2.font.size = Pt(32)
        p2.font.bold = True
        p2.space_before = Pt(4)
        p2.font.color.rgb = col

        p3 = ctf.add_paragraph()
        p3.text = note
        p3.font.name = "Segoe UI"
        p3.font.size = Pt(10)
        p3.font.color.rgb = TEXT_LIGHT

    # Right Wide Panel: Live Validation & Deployment Summary
    right_card = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.1), Inches(1.8), Inches(5.333), Inches(4.75))
    right_card.fill.solid()
    right_card.fill.fore_color.rgb = CARD_DARK
    right_card.line.color.rgb = ACCENT_TEAL
    right_card.line.width = Pt(1.5)

    rc_tf = right_card.text_frame
    rc_tf.word_wrap = True
    rc_tf.margin_top = Inches(0.25)
    rc_tf.margin_left = Inches(0.3)
    rc_tf.margin_right = Inches(0.3)

    rp1 = rc_tf.paragraphs[0]
    rp1.text = "LIVE SYSTEM PERFORMANCE BENCHMARKS"
    rp1.font.name = "Segoe UI"
    rp1.font.size = Pt(14)
    rp1.font.bold = True
    rp1.font.color.rgb = ACCENT_TEAL

    benchmarks = [
        ("Transmission Interval:", "ESP32 uplinks telemetry to ThingSpeak every 20 seconds flawlessly."),
        ("Client-Side Polling:", "Web dashboard pulls latest JSON feeds every 2s without screen flicker."),
        ("Dual-Trace Canvas Graph:", "Smooth Bezier spline curves rendered directly in HTML5 Canvas for the past 48 readings (~24 hours)."),
        ("Active State & Extremes:", "Dynamically parses historical session Low and High markers on the fly."),
        ("Graceful Degradation:", "Triggers connection warning if ESP32 power or Wi-Fi hotspot disconnects."),
        ("Live Web Link:", "smart-weather-station.vercel.app"),
        ("Source Code:", "github.com/prasoon9999/smart-weather-station")
    ]

    for label, text in benchmarks:
        bp = rc_tf.add_paragraph()
        bp.space_before = Pt(8)
        r1 = bp.add_run()
        r1.text = "• " + label + " "
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = TEXT_LIGHT

        r2 = bp.add_run()
        r2.text = text
        r2.font.bold = False
        r2.font.size = Pt(11)
        r2.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 8: CONCLUSION & FUTURE SCOPE
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_top_navigation(s8, "Roadmap", "Conclusion & Future Innovations", 8)

    # Left: Conclusion
    concl_card = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(1.8), Inches(5.6), Inches(4.8))
    concl_card.fill.solid()
    concl_card.fill.fore_color.rgb = CARD_DARK
    concl_card.line.color.rgb = ACCENT_CYAN
    concl_card.line.width = Pt(1.5)

    ctf = concl_card.text_frame
    ctf.word_wrap = True
    ctf.margin_top = Inches(0.25)
    ctf.margin_left = Inches(0.3)
    ctf.margin_right = Inches(0.3)

    cp1 = ctf.paragraphs[0]
    cp1.text = "PROJECT CONCLUSIONS"
    cp1.font.name = "Segoe UI"
    cp1.font.size = Pt(15)
    cp1.font.bold = True
    cp1.font.color.rgb = ACCENT_CYAN

    concl_bullets = [
        "End-to-End IoT Deployment: Seamlessly bridged physical hardware sensing (ESP32 + DHT22) with modern cloud platforms (ThingSpeak + Vercel).",
        "Edge Intelligence: Successfully demonstrated on-device computation of human thermal comfort indices (Heat Index and Dew Point).",
        "Zero-Dependency Lightweight UI: Crafted an ultra-fast, responsive dashboard without heavy external frameworks (pure HTML/CSS/Canvas).",
        "Real-World Reliability: Verified continuous 20-second sensor transmissions over mobile Wi-Fi tethering with full data integrity."
    ]

    for b in concl_bullets:
        bp = ctf.add_paragraph()
        bp.space_before = Pt(12)
        bp.text = f"✔  {b}"
        bp.font.name = "Segoe UI"
        bp.font.size = Pt(11.5)
        bp.font.color.rgb = TEXT_LIGHT

    # Right: Future Scope
    future_card = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8), Inches(5.6), Inches(4.8))
    future_card.fill.solid()
    future_card.fill.fore_color.rgb = CARD_DARK
    future_card.line.color.rgb = ACCENT_CORAL
    future_card.line.width = Pt(1.5)

    ftf = future_card.text_frame
    ftf.word_wrap = True
    ftf.margin_top = Inches(0.25)
    ftf.margin_left = Inches(0.3)
    ftf.margin_right = Inches(0.3)

    fp1 = ftf.paragraphs[0]
    fp1.text = "FUTURE EXPANSION ROADMAP"
    fp1.font.name = "Segoe UI"
    fp1.font.size = Pt(15)
    fp1.font.bold = True
    fp1.font.color.rgb = ACCENT_CORAL

    future_bullets = [
        ("Solar Energy Harvesting:", "Equipping system with 5V Solar panel + TP4056 Li-ion charge controller and ESP32 deep-sleep cycling for 100% off-grid autonomy."),
        ("Multi-Parameter Environmental Sensing:", "Integrating BMP280 (barometric pressure / altitude), MQ-135 (air quality index), and optical rain tipping bucket."),
        ("Automated Emergency Push Alerts:", "Setting up Webhooks for instant Telegram / WhatsApp / SMS push notifications during extreme heatwaves or storms."),
        ("Edge Machine Learning:", "Deploying TinyML models to predict rainfall probability locally based on falling barometric pressure and humidity slope.")
    ]

    for tag, desc in future_bullets:
        bp = ftf.add_paragraph()
        bp.space_before = Pt(10)
        r1 = bp.add_run()
        r1.text = "★  " + tag + " "
        r1.font.bold = True
        r1.font.size = Pt(11.5)
        r1.font.color.rgb = ACCENT_GOLD

        r2 = bp.add_run()
        r2.text = desc
        r2.font.bold = False
        r2.font.size = Pt(11.5)
        r2.font.color.rgb = TEXT_LIGHT

    # ==========================================
    # SLIDE 9: REFERENCES
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_top_navigation(s9, "Documentation", "References & Technical Standards", 9)

    refs_data = [
        ("01", "Espressif Systems", "ESP32 Series Datasheet - 2.4 GHz Wi-Fi and Bluetooth Combo Chip, Rev 3.9, 2023.", ACCENT_TEAL),
        ("02", "Aosong Electronics", "Digital-Output Relative Humidity & Temperature Sensor DHT22 (AM2302) Datasheet, Guangzhou.", ACCENT_CORAL),
        ("03", "IEEE Conference", "K. S. Rao, V. S. Murthy, \"IoT Based Real-Time Environmental and Weather Monitoring System using ESP32,\" IEEE ICST, 2021.", ACCENT_CYAN),
        ("04", "MathWorks Inc.", "ThingSpeak IoT Analytics Platform Documentation: RESTful API & Channels Architecture, 2024.", ACCENT_GOLD),
        ("05", "IEEE Access Journal", "N. H. Motlagh et al., \"Internet of Things (IoT) for Weather Monitoring: Recent Advances and Applications,\" IEEE Access, vol. 8, 2020.", ACCENT_TEAL),
        ("06", "Solomon Systech", "SSD1306 128 x 64 Dot Matrix OLED/PLED Driver with Controller Technical Specification, Rev 1.1, 2018.", ACCENT_CORAL),
    ]

    for i, (num, author, title, col) in enumerate(refs_data):
        y = Inches(1.8 + i * 0.82)
        row_card = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), y, Inches(11.533), Inches(0.72))
        row_card.fill.solid()
        row_card.fill.fore_color.rgb = CARD_DARK
        row_card.line.color.rgb = CARD_BORDER
        row_card.line.width = Pt(1)

        rtf = row_card.text_frame
        rtf.word_wrap = True
        rtf.margin_top = Inches(0.12)
        rtf.margin_left = Inches(0.2)
        rp = rtf.paragraphs[0]
        
        r1 = rp.add_run()
        r1.text = f"[{num}]  {author}: "
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = col

        r2 = rp.add_run()
        r2.text = f"\"{title}\""
        r2.font.bold = False
        r2.font.size = Pt(11)
        r2.font.color.rgb = TEXT_LIGHT

    # ==========================================
    # SLIDE 10: THANK YOU (MINIMALIST TECH CLOSING)
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)

    # Central spotlight container
    center_card = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.0), Inches(1.2), Inches(9.333), Inches(5.1))
    center_card.fill.solid()
    center_card.fill.fore_color.rgb = CARD_DARK
    center_card.line.color.rgb = CARD_BORDER
    center_card.line.width = Pt(1.5)

    ctf = center_card.text_frame
    ctf.word_wrap = True
    ctf.margin_top = Inches(0.8)

    cp = ctf.paragraphs[0]
    cp.text = "THANK YOU"
    cp.font.name = "Segoe UI"
    cp.font.size = Pt(50)
    cp.font.bold = True
    cp.alignment = PP_ALIGN.CENTER
    cp.font.color.rgb = TEXT_LIGHT

    cp2 = ctf.add_paragraph()
    cp2.text = "Questions & Technical Discussion"
    cp2.font.name = "Segoe UI"
    cp2.font.size = Pt(20)
    cp2.font.bold = True
    cp2.space_before = Pt(8)
    cp2.alignment = PP_ALIGN.CENTER
    cp2.font.color.rgb = ACCENT_TEAL

    cp3 = ctf.add_paragraph()
    cp3.text = "Galgotias College of Engineering & Technology, Greater Noida\nDepartment of Electronics and Communication Engineering"
    cp3.font.name = "Segoe UI"
    cp3.font.size = Pt(12)
    cp3.space_before = Pt(20)
    cp3.alignment = PP_ALIGN.CENTER
    cp3.font.color.rgb = TEXT_MUTED

    # Bottom pill badges for links
    pill_link = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.2), Inches(4.8), Inches(6.933), Inches(0.7))
    pill_link.fill.solid()
    pill_link.fill.fore_color.rgb = BG_DARK
    pill_link.line.color.rgb = CARD_BORDER
    pill_link.line.width = Pt(1)
    
    ptf = pill_link.text_frame
    ptf.margin_top = Inches(0.14)
    pp = ptf.paragraphs[0]
    pp.text = "🌐 smart-weather-station.vercel.app   |   📦 github.com/prasoon9999/smart-weather-station"
    pp.font.name = "Segoe UI"
    pp.font.size = Pt(11)
    pp.font.bold = True
    pp.alignment = PP_ALIGN.CENTER
    pp.font.color.rgb = ACCENT_CORAL

    out_path = r"c:\Users\dpras\Documents\smart weather station\Smart_Weather_Station_Presentation.pptx"
    prs.save(out_path)
    print(f"Aesthetic presentation successfully generated at {out_path}")

if __name__ == '__main__':
    create_aesthetic_deck()
