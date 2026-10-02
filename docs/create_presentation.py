#!/usr/bin/env python3
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Theme Colors
    COLOR_DARK_BG = RGBColor(22, 29, 44)       # Deep Navy
    COLOR_IBM_BLUE = RGBColor(15, 98, 254)     # Accent Blue
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_LIGHT_BG = RGBColor(244, 246, 248)   # Light Slate
    COLOR_TEXT_DARK = RGBColor(33, 37, 41)     # Dark Charcoal
    COLOR_MUTED_TEXT = RGBColor(108, 117, 125)  # Gray
    COLOR_GREEN = RGBColor(36, 161, 72)        # Active Green
    COLOR_LIGHT_BLUE = RGBColor(224, 240, 255) # Soft Blue Card
    COLOR_AWS_ORANGE = RGBColor(255, 153, 0)   # AWS Orange
    COLOR_DARK_TEAL = RGBColor(0, 93, 93)      # Teal Accent for DB Matrix

    FONT_FAMILY = "IBM Plex Sans"

    blank_layout = prs.slide_layouts[6]

    # ==========================================
    # SLIDE 1: Title Slide (Dark Background)
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    
    # Background shape
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_DARK_BG
    bg.line.fill.background()

    # Left accent line
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.833), Inches(1.5), Inches(0.08), Inches(4.5))
    accent.fill.solid()
    accent.fill.fore_color.rgb = COLOR_IBM_BLUE
    accent.line.fill.background()

    # Title & Subtitle text box
    tx_box = slide.shapes.add_textbox(Inches(1.2), Inches(1.5), Inches(11.0), Inches(4.5))
    tf = tx_box.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "Turbonomic & SQL Server\nFleet Inventory Integration"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(42)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    p.space_after = Pt(20)

    p2 = tf.add_paragraph()
    p2.text = "End-to-End Architectural Data Flow & Discovery Automation"
    p2.font.name = FONT_FAMILY
    p2.font.size = Pt(20)
    p2.font.color.rgb = RGBColor(168, 192, 255)
    p2.space_after = Pt(40)

    p3 = tf.add_paragraph()
    p3.text = "Automated WMI/WinRM Provisioning | Turbonomic REST API Integration | PSRemoting Metadata Harvester | Consolidated Fleet Reporting"
    p3.font.name = FONT_FAMILY
    p3.font.size = Pt(12)
    p3.font.color.rgb = COLOR_MUTED_TEXT

    # ==========================================
    # SLIDE 2: Executive Overview (4-Stage Process)
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    
    # Title
    tx_box = slide.shapes.add_textbox(Inches(0.833), Inches(0.5), Inches(11.667), Inches(0.8))
    tf = tx_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Executive Overview: End-to-End Orchestrated Pipeline"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK_BG

    # 4 Cards representing stages
    stages = [
        {
            "num": "01",
            "title": "Target VM Security & WinRM",
            "script": "Configure-TurbonomicWMI.ps1",
            "details": [
                "Creates local service user (turbowmi)",
                "Grants WMI namespaces remote enable",
                "Enables WinRM, NTLM & Remote Registry",
                "Opens firewall ports (5985, 135, dynamic)"
            ]
        },
        {
            "num": "02",
            "title": "WMI Target API Registration",
            "script": "create_wmi_target.py",
            "details": [
                "Connects to Turbonomic REST API (v3)",
                "Creates restricted IP-scoped groups",
                "Registers guest OS credentials target",
                "Triggers discovery validation & status check"
            ]
        },
        {
            "num": "03",
            "title": "Deep SQL & OS Harvest",
            "script": "Get-SQLServerInventory.ps1",
            "details": [
                "Remote connection via WinRM/PSRemoting",
                "Queries Windows config & host hardware",
                "Discovers SQL Server instances & state",
                "Queries engine & databases metadata via local auth"
            ]
        },
        {
            "num": "04",
            "title": "Master Report & Dashboard",
            "script": "Invoke-FleetReport.ps1",
            "details": [
                "Coordinates stage 2 & 3 scripts",
                "Reads both Turbonomic & SQL JSON data",
                "Performs dynamic merge on Name/IP key",
                "Generates unified dashboard with KPI cards"
            ]
        }
    ]

    card_width = Inches(2.7)
    card_height = Inches(5.0)
    card_gap = Inches(0.28)
    left_margin = Inches(0.833)
    top_pos = Inches(1.5)

    for i, stage in enumerate(stages):
        x = left_margin + i * (card_width + card_gap)
        
        # Outer Card
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top_pos, card_width, card_height)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_LIGHT_BG if i % 2 == 0 else COLOR_LIGHT_BLUE
        card.line.color.rgb = COLOR_IBM_BLUE if i == 3 else COLOR_MUTED_TEXT
        card.line.width = Pt(1.5 if i == 3 else 1.0)
        
        # Text Frame
        tx_box = slide.shapes.add_textbox(x + Inches(0.15), top_pos + Inches(0.15), card_width - Inches(0.3), card_height - Inches(0.3))
        tf = tx_box.text_frame
        tf.word_wrap = True
        
        # Stage Number
        p_num = tf.paragraphs[0]
        p_num.text = f"STAGE {stage['num']}"
        p_num.font.name = FONT_FAMILY
        p_num.font.size = Pt(10)
        p_num.font.bold = True
        p_num.font.color.rgb = COLOR_IBM_BLUE
        p_num.space_after = Pt(2)
        
        # Title
        p_title = tf.add_paragraph()
        p_title.text = stage["title"]
        p_title.font.name = FONT_FAMILY
        p_title.font.size = Pt(15)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_DARK_BG
        p_title.space_after = Pt(6)

        # Automation Script Badge
        p_script = tf.add_paragraph()
        p_script.text = stage["script"]
        p_script.font.name = "Courier New"
        p_script.font.size = Pt(9.5)
        p_script.font.bold = True
        p_script.font.color.rgb = COLOR_GREEN if "ps1" in stage["script"] else RGBColor(120, 50, 200)
        p_script.space_after = Pt(12)

        # Details list
        for detail in stage["details"]:
            p_det = tf.add_paragraph()
            p_det.text = "• " + detail
            p_det.font.name = FONT_FAMILY
            p_det.font.size = Pt(11)
            p_det.font.color.rgb = COLOR_TEXT_DARK
            p_det.space_after = Pt(6)

    # ==========================================
    # SLIDE 3: Stage 1 Details (Target VM Configuration)
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    
    # Title
    tx_box = slide.shapes.add_textbox(Inches(0.833), Inches(0.5), Inches(11.667), Inches(0.8))
    tf = tx_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Stage 1: Target Windows VM Security & Protocol Setup"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK_BG

    # Two column layout
    col_width = Inches(5.5)
    col_height = Inches(5.2)
    col_gap = Inches(0.667)
    
    # Left Column: User Account and Local Security Groups
    left_x = Inches(0.833)
    card1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_x, Inches(1.5), col_width, col_height)
    card1.fill.solid()
    card1.fill.fore_color.rgb = COLOR_WHITE
    card1.line.color.rgb = COLOR_LIGHT_BG
    
    tx_box1 = slide.shapes.add_textbox(left_x + Inches(0.2), Inches(1.7), col_width - Inches(0.4), col_height - Inches(0.4))
    tf1 = tx_box1.text_frame
    tf1.word_wrap = True
    
    p = tf1.paragraphs[0]
    p.text = "WMI Service Account (Option B - Hardened Non-Admin)"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_IBM_BLUE
    p.space_after = Pt(12)
    
    details_left = [
        ("Dedicated User Creation", "Creates 'turbowmi' user locally, disables password expiry to ensure uninterrupted monitoring."),
        ("Required Local Group Membership", "Adds account to 'WinRMRemoteWMIUsers__' (or Remote Management Users) and 'Performance Monitor Users' (no Local Admin required)."),
        ("WMI Namespace Permissions", "Grants 'Execute Methods', 'Enable Account', and 'Remote Enable' on Root and Root\\CIMV2 namespaces, recursively applied to all subnamespaces."),
        ("UAC Remote Access Fix (Registry)", "Sets 'LocalAccountTokenFilterPolicy = 1' in HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System. Crucial to allow non-RID 500 local accounts remote access via WinRM.")
    ]
    
    for title, text in details_left:
        p_t = tf1.add_paragraph()
        p_t.text = f"•  {title}:"
        p_t.font.name = FONT_FAMILY
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK_BG
        
        p_tx = tf1.add_paragraph()
        p_tx.text = f"    {text}"
        p_tx.font.name = FONT_FAMILY
        p_tx.font.size = Pt(11)
        p_tx.font.color.rgb = COLOR_TEXT_DARK
        p_tx.space_after = Pt(10)

    # Right Column: Network Protocols & Services
    right_x = left_x + col_width + col_gap
    card2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, right_x, Inches(1.5), col_width, col_height)
    card2.fill.solid()
    card2.fill.fore_color.rgb = COLOR_WHITE
    card2.line.color.rgb = COLOR_LIGHT_BG
    
    tx_box2 = slide.shapes.add_textbox(right_x + Inches(0.2), Inches(1.7), col_width - Inches(0.4), col_height - Inches(0.4))
    tf2 = tx_box2.text_frame
    tf2.word_wrap = True
    
    p = tf2.paragraphs[0]
    p.text = "WinRM HTTP, Remote Registry & Port Openings"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_IBM_BLUE
    p.space_after = Pt(12)

    details_right = [
        ("WS-Management Configuration", "Runs 'winrm quickconfig' to start service and set to Automatic. Configures default HTTP listener on port 5985."),
        ("Negotiate Authentication & Encr.", "Enables 'Negotiate' auth to support NTLM for WORKGROUP local identities. Set 'AllowUnencrypted = True' for SOAP messages (auth is still encrypted)."),
        ("Remote Registry Activation", "Set 'RemoteRegistry' service to Automatic startup and starts it immediately. Mandatory for SQL Server instances discovery by external tools."),
        ("Firewall Rules Deployment", "Opens ports in Windows Firewall: TCP 5985 (WinRM HTTP), TCP 135 (RPC Endport Mapper), and TCP 49152-65535 (WMI Dynamic RPC Range).")
    ]
    
    for title, text in details_right:
        p_t = tf2.add_paragraph()
        p_t.text = f"•  {title}:"
        p_t.font.name = FONT_FAMILY
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK_BG
        
        p_tx = tf2.add_paragraph()
        p_tx.text = f"    {text}"
        p_tx.font.name = FONT_FAMILY
        p_tx.font.size = Pt(11)
        p_tx.font.color.rgb = COLOR_TEXT_DARK
        p_tx.space_after = Pt(10)

    # ==========================================
    # SLIDE 4: Stage 2 Details (Turbonomic WMI API Target Setup)
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    
    # Title
    tx_box = slide.shapes.add_textbox(Inches(0.833), Inches(0.5), Inches(11.667), Inches(0.8))
    tf = tx_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Stage 2: WMI Target REST API Automation"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK_BG

    # 3 Steps card row
    step_width = Inches(3.64)
    step_gap = Inches(0.35)
    
    steps_data = [
        {
            "step": "1",
            "title": "Auth & Probe Verification",
            "endpoints": ["POST /api/v3/login", "GET /api/v3/targets/specs"],
            "desc": "Authenticates using admin credentials (turboplanuser). Queries the Target Specification endpoint to verify that the 'WMI' probe is actively enabled on the Turbonomic instance."
        },
        {
            "step": "2",
            "title": "IP Scope Resolution & Grouping",
            "endpoints": ["POST /api/v3/search", "POST /api/v3/groups"],
            "desc": "Searches for the target IP in existing VM entities. If the VM is not yet discovered, creates a dynamic, narrow IP-scoped group ('WMI-Scope-IP' with 'vmsByGuestName' rule) to contain WMI credentials target scope."
        },
        {
            "step": "3",
            "title": "POST Target & Rediscover",
            "endpoints": ["POST /api/v3/targets", "POST /api/v3/targets/{id}/rediscover"],
            "desc": "Constructs TargetApiDTO: Category='Guest OS Processes', Type='WMI', inputs: targetEntities (the resolved group UUID), username, password, NTLM=true, secure=false. Submits DTO, triggers rediscovery, and polls target state until 'Validated'."
        }
    ]

    for j, s_data in enumerate(steps_data):
        x = left_margin + j * (step_width + step_gap)
        
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.5), step_width, Inches(5.0))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_WHITE
        card.line.color.rgb = COLOR_IBM_BLUE
        card.line.width = Pt(1.5)
        
        tx_box = slide.shapes.add_textbox(x + Inches(0.15), Inches(1.65), step_width - Inches(0.3), Inches(4.7))
        tf_step = tx_box.text_frame
        tf_step.word_wrap = True
        
        # Step Circle/Indicator
        p_step = tf_step.paragraphs[0]
        p_step.text = f"STEP {s_data['step']}"
        p_step.font.name = FONT_FAMILY
        p_step.font.size = Pt(11)
        p_step.font.bold = True
        p_step.font.color.rgb = COLOR_IBM_BLUE
        p_step.space_after = Pt(4)
        
        # Title
        p_title = tf_step.add_paragraph()
        p_title.text = s_data["title"]
        p_title.font.name = FONT_FAMILY
        p_title.font.size = Pt(16)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_DARK_BG
        p_title.space_after = Pt(12)
        
        # Endpoints
        p_end_lbl = tf_step.add_paragraph()
        p_end_lbl.text = "Key API Calls:"
        p_end_lbl.font.name = FONT_FAMILY
        p_end_lbl.font.size = Pt(11)
        p_end_lbl.font.bold = True
        p_end_lbl.font.color.rgb = COLOR_TEXT_DARK
        
        for ep in s_data["endpoints"]:
            p_ep = tf_step.add_paragraph()
            p_ep.text = "  • " + ep
            p_ep.font.name = "Courier New"
            p_ep.font.size = Pt(9.5)
            p_ep.font.bold = True
            p_ep.font.color.rgb = COLOR_GREEN if "POST" in ep else COLOR_IBM_BLUE
            p_ep.space_after = Pt(4)
            
        p_ep.space_after = Pt(14)
        
        # Description
        p_desc = tf_step.add_paragraph()
        p_desc.text = s_data["desc"]
        p_desc.font.name = FONT_FAMILY
        p_desc.font.size = Pt(11.5)
        p_desc.font.color.rgb = COLOR_TEXT_DARK
        p_desc.space_after = Pt(0)

    # ==========================================
    # SLIDE 5: Stage 3 Details (Get-SQLServerInventory.ps1)
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    
    # Title
    tx_box = slide.shapes.add_textbox(Inches(0.833), Inches(0.5), Inches(11.667), Inches(0.8))
    tf = tx_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Stage 3: Deep SQL Server & DB Metadata Harvester"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK_BG

    # 3-Tier process description
    y_pos = Inches(1.5)
    row_height = Inches(1.6)
    row_gap = Inches(0.18)

    tiers = [
        {
            "lbl": "WINDOWS & VM",
            "title": "Local Server Inventory",
            "desc": "Initiates PSRemoting/WinRM to collect machine-specific configuration: FQDN, domain/workgroup membership, manufacturer, bios UUID, serial number, OS build/version, local logical disks, remaining free space, physical network adapters, assigned IP addresses, DNS servers, and default gateways."
        },
        {
            "lbl": "SQL WMI / REGISTRY",
            "title": "Service and Instance Discovery",
            "desc": "Inspects WMI namespace on target (e.g. root\\Microsoft\\SqlServer\\ComputerManagement11/12/15) to discover active database instances (default & named), service states, startup accounts, engine binaries paths. Extracts exact database patches levels, service packs, and TCP port configs directly from remote Registry."
        },
        {
            "lbl": "SQL ENGINE / T-SQL",
            "title": "Local Database and Feature Catalog",
            "desc": "Establishes a local database connection using integrated security. Runs detailed metadata catalog queries: instances clustering/HADR availability, security authentication modes, and collations. Loops through databases to extract recovery models, file sizes, transaction logs reuse status, growth specs, Query Store, and database encryption."
        }
    ]

    for k, tier in enumerate(tiers):
        cur_y = y_pos + k * (row_height + row_gap)
        
        # Row box
        row_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_margin, cur_y, Inches(11.667), row_height)
        row_bg.fill.solid()
        row_bg.fill.fore_color.rgb = COLOR_LIGHT_BG
        row_bg.line.color.rgb = COLOR_WHITE
        
        # Left Label Block
        lbl_box = slide.shapes.add_textbox(left_margin, cur_y, Inches(2.2), row_height)
        tf_lbl = lbl_box.text_frame
        tf_lbl.word_wrap = True
        p_lbl = tf_lbl.paragraphs[0]
        p_lbl.text = f"\n{tier['lbl']}"
        p_lbl.font.name = FONT_FAMILY
        p_lbl.font.size = Pt(13)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = COLOR_IBM_BLUE
        p_lbl.alignment = PP_ALIGN.CENTER
        
        # Divider line inside row
        div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_margin + Inches(2.3), cur_y + Inches(0.15), Inches(0.02), row_height - Inches(0.3))
        div.fill.solid()
        div.fill.fore_color.rgb = COLOR_MUTED_TEXT
        div.line.fill.background()
        
        # Description text
        desc_box = slide.shapes.add_textbox(left_margin + Inches(2.5), cur_y + Inches(0.15), Inches(8.9), row_height - Inches(0.3))
        tf_desc = desc_box.text_frame
        tf_desc.word_wrap = True
        
        p_t = tf_desc.paragraphs[0]
        p_t.text = tier["title"]
        p_t.font.name = FONT_FAMILY
        p_t.font.size = Pt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK_BG
        p_t.space_after = Pt(2)
        
        p_tx = tf_desc.add_paragraph()
        p_tx.text = tier["desc"]
        p_tx.font.name = FONT_FAMILY
        p_tx.font.size = Pt(10.5)
        p_tx.font.color.rgb = COLOR_TEXT_DARK

    # ==========================================
    # SLIDE 6: Stage 4 Details (Invoke-FleetReport.ps1)
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    
    # Title
    tx_box = slide.shapes.add_textbox(Inches(0.833), Inches(0.5), Inches(11.667), Inches(0.8))
    tf = tx_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Stage 4: Consolidated Master Fleet Orchestration"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK_BG

    # 3 horizontal boxes representing input, process, output
    box_w = Inches(3.64)
    box_gap = Inches(0.35)
    
    # Card 1: Input & Parameters
    x = left_margin
    b1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.5), box_w, Inches(5.0))
    b1.fill.solid()
    b1.fill.fore_color.rgb = COLOR_WHITE
    b1.line.color.rgb = COLOR_LIGHT_BG
    
    tx_b1 = slide.shapes.add_textbox(x + Inches(0.15), Inches(1.65), box_w - Inches(0.3), Inches(4.7))
    tf_b1 = tx_b1.text_frame
    tf_b1.word_wrap = True
    p = tf_b1.paragraphs[0]
    p.text = "Pipeline Inputs & Args"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_IBM_BLUE
    p.space_after = Pt(12)
    
    inputs_text = [
        ("servers.csv", "Central configuration file listing VM names and remote IPv4 addresses."),
        ("turbonomic/config.json", "REST API connection configuration: target url, credentials, SSL verification state, timeout settings."),
        ("Days (Default: 1)", "Historical query window passed to Python client to calculate peak CPU/Mem/IO utilization spikes.")
    ]
    for name, desc in inputs_text:
        p_t = tf_b1.add_paragraph()
        p_t.text = f"•  {name}:"
        p_t.font.name = FONT_FAMILY
        p_t.font.size = Pt(12)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK_BG
        
        p_tx = tf_b1.add_paragraph()
        p_tx.text = f"    {desc}"
        p_tx.font.name = FONT_FAMILY
        p_tx.font.size = Pt(10.5)
        p_tx.font.color.rgb = COLOR_TEXT_DARK
        p_tx.space_after = Pt(8)

    # Card 2: Orchestrated Merging
    x += box_w + box_gap
    b2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.5), box_w, Inches(5.0))
    b2.fill.solid()
    b2.fill.fore_color.rgb = COLOR_LIGHT_BLUE
    b2.line.color.rgb = COLOR_IBM_BLUE
    
    tx_b2 = slide.shapes.add_textbox(x + Inches(0.15), Inches(1.65), box_w - Inches(0.3), Inches(4.7))
    tf_b2 = tx_b2.text_frame
    tf_b2.word_wrap = True
    p = tf_b2.paragraphs[0]
    p.text = "Double-JSON Join Logic"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK_BG
    p.space_after = Pt(12)
    
    p_body = tf_b2.add_paragraph()
    p_body.text = "Invoke-FleetReport.ps1 triggers the sub-collectors sequentially, then ingests both structured JSON outputs:\n\n" \
                  "1. Ingests SQL inventory JSON\n" \
                  "2. Ingests Turbonomic VM metrics JSON\n\n" \
                  "Merge Key resolution:\n" \
                  "The orchestrator correlates records by mapping VM names and host IP addresses. It resolves name discrepancies by normalizing case and comparing installed network interface IP lists, creating a 1:1 unified record map per target VM."
    p_body.font.name = FONT_FAMILY
    p_body.font.size = Pt(11)
    p_body.font.color.rgb = COLOR_TEXT_DARK

    # Card 3: Outputs & Dashboards
    x += box_w + box_gap
    b3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.5), box_w, Inches(5.0))
    b3.fill.solid()
    b3.fill.fore_color.rgb = COLOR_WHITE
    b3.line.color.rgb = COLOR_LIGHT_BG
    
    tx_b3 = slide.shapes.add_textbox(x + Inches(0.15), Inches(1.65), box_w - Inches(0.3), Inches(4.7))
    tf_b3 = tx_b3.text_frame
    tf_b3.word_wrap = True
    p = tf_b3.paragraphs[0]
    p.text = "Unified Fleet Dashboard"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_IBM_BLUE
    p.space_after = Pt(12)
    
    outputs_text = [
        ("Dashboard HTML", "Generates single, fully responsive HTML dashboard. Includes inline CSS styling, sidebar navigations, and direct CSV download options."),
        ("Key Metrics Panel", "Highlights total core counts, provisioned memory vs peak utilization, host storage caps vs database file allocations, monthly cloud cost, and estimated savings."),
        ("Active Actions Queue", "Lists pending resize, scaling, buy, and move actions with exact risk categories, current vs recommended VM sizes, and monthly savings.")
    ]
    for name, desc in outputs_text:
        p_t = tf_b3.add_paragraph()
        p_t.text = f"•  {name}:"
        p_t.font.name = FONT_FAMILY
        p_t.font.size = Pt(12)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK_BG
        
        p_tx = tf_b3.add_paragraph()
        p_tx.text = f"    {desc}"
        p_tx.font.name = FONT_FAMILY
        p_tx.font.size = Pt(10.5)
        p_tx.font.color.rgb = COLOR_TEXT_DARK
        p_tx.space_after = Pt(8)

    # ==========================================
    # SLIDE 7: Visual Architecture Block Flow Map (UPDATED FOR MULTI-OS/DB)
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    
    # Title
    tx_box = slide.shapes.add_textbox(Inches(0.833), Inches(0.5), Inches(11.667), Inches(0.8))
    tf = tx_box.text_frame
    p = tf.paragraphs[0]
    p.text = "End-to-End Architectural Data Flow Map"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK_BG

    # Central Windows Mgmt VM Box
    box_mgmt = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.833), Inches(2.2), Inches(3.2), Inches(3.8))
    box_mgmt.fill.solid()
    box_mgmt.fill.fore_color.rgb = COLOR_DARK_BG
    box_mgmt.line.color.rgb = COLOR_IBM_BLUE
    box_mgmt.line.width = Pt(1.5)
    
    tx_m = slide.shapes.add_textbox(Inches(0.833), Inches(2.3), Inches(3.2), Inches(3.6))
    tf_m = tx_m.text_frame
    tf_m.word_wrap = True
    p_m = tf_m.paragraphs[0]
    p_m.text = "Central Orchestration\n(Management VM / Script)"
    p_m.font.name = FONT_FAMILY
    p_m.font.size = Pt(14)
    p_m.font.bold = True
    p_m.font.color.rgb = COLOR_WHITE
    p_m.alignment = PP_ALIGN.CENTER
    p_m.space_after = Pt(14)
    
    pm_steps = [
        "Invoke-FleetReport.ps1",
        "• Reads servers.csv & config.json",
        "• Runs Get-SQLServerInventory.ps1",
        "• Runs get_vm_metrics.py (Python)",
        "• Performs Double-JSON key merge",
        "• Writes fleet-dashboard.html"
    ]
    for s in pm_steps:
        ps = tf_m.add_paragraph()
        ps.text = s
        ps.font.name = FONT_FAMILY
        ps.font.size = Pt(10)
        ps.font.color.rgb = RGBColor(200, 220, 255) if "ps1" in s or "py" in s else COLOR_WHITE
        ps.space_after = Pt(6)

    # Arrow 1 to Target Server (WinRM & SSH Protocols)
    arrow1 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(4.2), Inches(2.8), Inches(1.1), Inches(0.4))
    arrow1.fill.solid()
    arrow1.fill.fore_color.rgb = COLOR_IBM_BLUE
    arrow1.line.fill.background()
    
    lbl1 = slide.shapes.add_textbox(Inches(4.1), Inches(2.2), Inches(1.3), Inches(0.6))
    lbl1.text_frame.word_wrap = True
    p_l1 = lbl1.text_frame.paragraphs[0]
    p_l1.text = "WinRM (5985) & SSH (22)"
    p_l1.font.name = FONT_FAMILY
    p_l1.font.size = Pt(8.5)
    p_l1.font.color.rgb = COLOR_TEXT_DARK
    p_l1.alignment = PP_ALIGN.CENTER

    # Arrow 2 to Turbonomic Server
    arrow2 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(4.2), Inches(4.8), Inches(1.1), Inches(0.4))
    arrow2.fill.solid()
    arrow2.fill.fore_color.rgb = RGBColor(120, 50, 200)
    arrow2.line.fill.background()
    
    lbl2 = slide.shapes.add_textbox(Inches(4.1), Inches(4.2), Inches(1.3), Inches(0.6))
    lbl2.text_frame.word_wrap = True
    p_l2 = lbl2.text_frame.paragraphs[0]
    p_l2.text = "REST HTTPS (443)"
    p_l2.font.name = FONT_FAMILY
    p_l2.font.size = Pt(8.5)
    p_l2.font.color.rgb = COLOR_TEXT_DARK
    p_l2.alignment = PP_ALIGN.CENTER

    # Target Server Hosts Box (UPDATED FOR WINDOWS / LINUX & SQL/ORACLE)
    box_target = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.5), Inches(1.5), Inches(3.2), Inches(2.2))
    box_target.fill.solid()
    box_target.fill.fore_color.rgb = COLOR_LIGHT_BG
    box_target.line.color.rgb = COLOR_MUTED_TEXT
    
    tx_t = slide.shapes.add_textbox(Inches(5.5), Inches(1.6), Inches(3.2), Inches(2.0))
    tf_t = tx_t.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.text = "Target Hosts & Databases"
    p_t.font.name = FONT_FAMILY
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_DARK_BG
    p_t.alignment = PP_ALIGN.CENTER
    p_t.space_after = Pt(6)
    
    t_items = [
        "OS Support: Windows & Linux",
        "WMI (WinRM) / Native SSH",
        "SQL Server & Oracle Databases",
        "Engine Metadata Probes (T-SQL/SQL)"
    ]
    for s in t_items:
        ps = tf_t.add_paragraph()
        ps.text = "• " + s
        ps.font.name = FONT_FAMILY
        ps.font.size = Pt(9.5)
        ps.font.color.rgb = COLOR_TEXT_DARK
        ps.space_after = Pt(2)

    # Turbonomic API Server Box
    box_turbo = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.5), Inches(4.1), Inches(3.2), Inches(2.2))
    box_turbo.fill.solid()
    box_turbo.fill.fore_color.rgb = COLOR_LIGHT_BLUE
    box_turbo.line.color.rgb = COLOR_IBM_BLUE
    
    tx_tb = slide.shapes.add_textbox(Inches(5.5), Inches(4.2), Inches(3.2), Inches(2.0))
    tf_tb = tx_tb.text_frame
    tf_tb.word_wrap = True
    p_tb = tf_tb.paragraphs[0]
    p_tb.text = "Turbonomic API / Server"
    p_tb.font.name = FONT_FAMILY
    p_tb.font.size = Pt(13)
    p_tb.font.bold = True
    p_tb.font.color.rgb = COLOR_DARK_BG
    p_tb.alignment = PP_ALIGN.CENTER
    p_tb.space_after = Pt(6)
    
    tb_items = [
        "REST API v3 (Endpoints specs)",
        "IP Scoped Groups (targetEntities)",
        "VCPU/VMem VM commodity metrics",
        "Pending action & recommended sizes"
    ]
    for s in tb_items:
        ps = tf_tb.add_paragraph()
        ps.text = "• " + s
        ps.font.name = FONT_FAMILY
        ps.font.size = Pt(9.5)
        ps.font.color.rgb = COLOR_TEXT_DARK
        ps.space_after = Pt(2)

    # Arrow Turbonomic -> Target (WMI Polling over WinRM)
    arrow_t2v = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(9.1), Inches(3.4), Inches(0.4), Inches(1.0))
    arrow_t2v.fill.solid()
    arrow_t2v.fill.fore_color.rgb = COLOR_GREEN
    arrow_t2v.line.fill.background()
    
    lbl3 = slide.shapes.add_textbox(Inches(8.5), Inches(2.8), Inches(1.6), Inches(0.5))
    lbl3.text_frame.word_wrap = True
    p_l3 = lbl3.text_frame.paragraphs[0]
    p_l3.text = "WMI WinRM / SSH Probes"
    p_l3.font.name = FONT_FAMILY
    p_l3.font.size = Pt(8.5)
    p_l3.font.color.rgb = COLOR_TEXT_DARK
    p_l3.alignment = PP_ALIGN.CENTER

    # Merged Output HTML Dash box on right
    box_out = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.7), Inches(2.2), Inches(2.8), Inches(3.8))
    box_out.fill.solid()
    box_out.fill.fore_color.rgb = COLOR_WHITE
    box_out.line.color.rgb = COLOR_GREEN
    box_out.line.width = Pt(2.0)
    
    tx_o = slide.shapes.add_textbox(Inches(9.7), Inches(2.3), Inches(2.8), Inches(3.6))
    tf_o = tx_o.text_frame
    tf_o.word_wrap = True
    p_o = tf_o.paragraphs[0]
    p_o.text = "Output Artifacts"
    p_o.font.name = FONT_FAMILY
    p_o.font.size = Pt(14)
    p_o.font.bold = True
    p_o.font.color.rgb = COLOR_GREEN
    p_o.alignment = PP_ALIGN.CENTER
    p_o.space_after = Pt(14)
    
    o_items = [
        "fleet-dashboard.html",
        "• High-level KPI Cards",
        "• Resource Capacity gauges",
        "• Detailed database specs",
        "• Pending scaling actions",
        "• Downloadable CSV reports"
    ]
    for s in o_items:
        ps = tf_o.add_paragraph()
        ps.text = s
        ps.font.name = FONT_FAMILY
        ps.font.size = Pt(10)
        ps.font.color.rgb = COLOR_TEXT_DARK
        ps.space_after = Pt(6)

    # Arrow from central mgmt directly to Output Artifacts
    arrow3 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(2.4), Inches(6.1), Inches(6.0), Inches(0.3))
    arrow3.fill.solid()
    arrow3.fill.fore_color.rgb = COLOR_GREEN
    arrow3.line.fill.background()
    
    lbl_mg = slide.shapes.add_textbox(Inches(4.1), Inches(6.3), Inches(2.8), Inches(0.5))
    lbl_mg.text_frame.word_wrap = True
    p_lmg = lbl_mg.text_frame.paragraphs[0]
    p_lmg.text = "Saves Merged Data & Reports"
    p_lmg.font.name = FONT_FAMILY
    p_lmg.font.size = Pt(9)
    p_lmg.font.color.rgb = COLOR_GREEN
    p_lmg.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 8: AWS-Native Style Cloud Architecture
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    
    # Title
    tx_box = slide.shapes.add_textbox(Inches(0.833), Inches(0.5), Inches(11.667), Inches(0.8))
    tf = tx_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Cloud Architecture: AWS-Native Deployment Model"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK_BG

    # 4 Cards representing AWS Services
    aws_cards = [
        {
            "service": "Amazon EC2 (Windows)",
            "lbl": "COMPUTE & TARGETS",
            "desc": [
                "Self-managed SQL Server deployed on EC2 instances in a Private VPC Subnet.",
                "EBS storage volumes monitored by Turbonomic for dynamic IOPS and storage scale alerts.",
                "VPC Security Groups restrict WinRM HTTP traffic (5985) only to the management orchestrator."
            ]
        },
        {
            "service": "AWS SSM & Secrets Manager",
            "lbl": "SECURE ORCHESTRATION",
            "desc": [
                "AWS Systems Manager Run Command automates Configure-TurbonomicWMI.ps1 execution.",
                "AWS Secrets Manager securely stores and rotates the 'turbowmi' service account credentials.",
                "SSM Parameter Store stores Turbonomic API URLs and configuration details, removing hardcoded local configs."
            ]
        },
        {
            "service": "Turbonomic on AWS",
            "lbl": "MONITORING & COGNITION",
            "desc": [
                "Turbonomic Marketplace Appliance or SaaS deployed inside a dedicated Management VPC.",
                "Uses AWS VPC Peering or AWS Transit Gateway to safely connect to and monitor Private Subnet EC2 VMs.",
                "Translates live WMI commodity stats into exact AWS EC2 resize and storage elastic scaling recommendations."
            ]
        },
        {
            "service": "Amazon S3 & CloudFront",
            "lbl": "HOSTING & COGNITIVE DASHBOARD",
            "desc": [
                "Master reporting script runs on an EC2 mgmt node and uploads fleet-dashboard.html to an S3 Bucket.",
                "Amazon CloudFront caches and securely serves the static HTML dashboard to internal users via SSL/HTTPS.",
                "AWS IAM bucket policies and VPC endpoints enforce restricted administrative access Control Lists."
            ]
        }
    ]

    card_width = Inches(2.7)
    card_height = Inches(5.0)
    card_gap = Inches(0.28)
    left_margin = Inches(0.833)
    top_pos = Inches(1.5)

    for i, acard in enumerate(aws_cards):
        x = left_margin + i * (card_width + card_gap)
        
        # Outer Card
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top_pos, card_width, card_height)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_WHITE
        card.line.color.rgb = COLOR_AWS_ORANGE
        card.line.width = Pt(1.5)
        
        # Inner text frame
        tx_box = slide.shapes.add_textbox(x + Inches(0.15), top_pos + Inches(0.15), card_width - Inches(0.3), card_height - Inches(0.3))
        tf_ac = tx_box.text_frame
        tf_ac.word_wrap = True
        
        # Label
        p_lbl = tf_ac.paragraphs[0]
        p_lbl.text = acard["lbl"]
        p_lbl.font.name = FONT_FAMILY
        p_lbl.font.size = Pt(9.5)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = COLOR_AWS_ORANGE
        p_lbl.space_after = Pt(2)
        
        # Service Title
        p_srv = tf_ac.add_paragraph()
        p_srv.text = acard["service"]
        p_srv.font.name = FONT_FAMILY
        p_srv.font.size = Pt(14)
        p_srv.font.bold = True
        p_srv.font.color.rgb = COLOR_DARK_BG
        p_srv.space_after = Pt(12)

        # Details list
        for bullet in acard["desc"]:
            p_b = tf_ac.add_paragraph()
            p_b.text = "• " + bullet
            p_b.font.name = FONT_FAMILY
            p_b.font.size = Pt(10.5)
            p_b.font.color.rgb = COLOR_TEXT_DARK
            p_b.space_after = Pt(6)

    # ==========================================
    # SLIDE 9: Multi-Platform Support: OS & Database Matrix
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    
    # Title
    tx_box = slide.shapes.add_textbox(Inches(0.833), Inches(0.5), Inches(11.667), Inches(0.8))
    tf = tx_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Multi-Platform Support: OS & Database Matrix"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK_BG

    # Split Slide Layout: Two Large Columns (Operating Systems vs Database Engines)
    half_width = Inches(5.6)
    half_gap = Inches(0.467)
    
    # Left Box: Operating Systems (Windows & Linux)
    left_x = Inches(0.833)
    os_card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_x, Inches(1.5), half_width, Inches(5.2))
    os_card.fill.solid()
    os_card.fill.fore_color.rgb = COLOR_WHITE
    os_card.line.color.rgb = COLOR_IBM_BLUE
    os_card.line.width = Pt(1.5)
    
    tx_os = slide.shapes.add_textbox(left_x + Inches(0.2), Inches(1.65), half_width - Inches(0.4), Inches(4.9))
    tf_os = tx_os.text_frame
    tf_os.word_wrap = True
    
    p = tf_os.paragraphs[0]
    p.text = "OPERATING SYSTEMS"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_IBM_BLUE
    p.space_after = Pt(6)
    
    p_os_sub = tf_os.add_paragraph()
    p_os_sub.text = "Windows Server & Linux Support"
    p_os_sub.font.name = FONT_FAMILY
    p_os_sub.font.size = Pt(18)
    p_os_sub.font.bold = True
    p_os_sub.font.color.rgb = COLOR_DARK_BG
    p_os_sub.space_after = Pt(14)
    
    os_bullets = [
        ("Microsoft Windows Server", "Monitored natively via WMI over WinRM protocols on Ports 5985/5986. Supports NTLM/Negotiate authentication for workgroups and Kerberos for Active Directory. Reads system registry, active processes, local disks geometry, and services states."),
        ("Linux Platforms (RHEL, Ubuntu, SLES)", "Monitored natively via secure SSH connections on Port 22. Supports public-key or password-based credentials. Turbonomic parses standard virtual file structures like /proc, /sys/class, and command-line execution outputs (e.g. df, top) to fetch VM and container process stats.")
    ]
    for os_name, os_desc in os_bullets:
        p_t = tf_os.add_paragraph()
        p_t.text = f"•  {os_name}:"
        p_t.font.name = FONT_FAMILY
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK_BG
        
        p_tx = tf_os.add_paragraph()
        p_tx.text = f"    {os_desc}"
        p_tx.font.name = FONT_FAMILY
        p_tx.font.size = Pt(11)
        p_tx.font.color.rgb = COLOR_TEXT_DARK
        p_tx.space_after = Pt(10)

    # Right Box: Database Platforms (SQL Server & Oracle)
    right_x = left_x + half_width + half_gap
    db_card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, right_x, Inches(1.5), half_width, Inches(5.2))
    db_card.fill.solid()
    db_card.fill.fore_color.rgb = COLOR_WHITE
    db_card.line.color.rgb = COLOR_DARK_TEAL
    db_card.line.width = Pt(1.5)
    
    tx_db = slide.shapes.add_textbox(right_x + Inches(0.2), Inches(1.65), half_width - Inches(0.4), Inches(4.9))
    tf_db = tx_db.text_frame
    tf_db.word_wrap = True
    
    p = tf_db.paragraphs[0]
    p.text = "DATABASE ENGINES"
    p.font.name = FONT_FAMILY
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK_TEAL
    p.space_after = Pt(6)
    
    p_db_sub = tf_db.add_paragraph()
    p_db_sub.text = "MS SQL Server & Oracle Database"
    p_db_sub.font.name = FONT_FAMILY
    p_db_sub.font.size = Pt(18)
    p_db_sub.font.bold = True
    p_db_sub.font.color.rgb = COLOR_DARK_BG
    p_db_sub.space_after = Pt(14)
    
    db_bullets = [
        ("Microsoft SQL Server", "Supports both Windows Integrated and SQL-specific credentials. Collects instances clustering configurations, HADR availability, system buffers, database size specs, log reuse status, and transactional log growth metrics."),
        ("Oracle Database (on Windows/Linux)", "Supports connections via Oracle Listener (Port 1521) or local bequeath loops. Connects to Oracle DB instances and Pluggable Databases (PDBs) to query tablespaces, SGA/PGA memory structures, active processes, and ASM disk-group allocations.")
    ]
    for db_name, db_desc in db_bullets:
        p_t = tf_db.add_paragraph()
        p_t.text = f"•  {db_name}:"
        p_t.font.name = FONT_FAMILY
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK_BG
        
        p_tx = tf_db.add_paragraph()
        p_tx.text = f"    {db_desc}"
        p_tx.font.name = FONT_FAMILY
        p_tx.font.size = Pt(11)
        p_tx.font.color.rgb = COLOR_TEXT_DARK
        p_tx.space_after = Pt(10)

    prs.save("Turbonomic_SQL_Fleet_Architecture.pptx")
    print("SUCCESS: Turbonomic_SQL_Fleet_Architecture.pptx updated with Multi-Platform Flow Map details.")

if __name__ == "__main__":
    create_presentation()
