from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── Colour palette ──────────────────────────────────────────────
BG_DARK       = RGBColor(0x0A, 0x0D, 0x1A)   # near-black navy
ACCENT_RED    = RGBColor(0xC0, 0x00, 0x00)   # threat-red
ACCENT_ORANGE = RGBColor(0xE8, 0x6C, 0x00)   # warning orange
TEXT_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_LIGHT    = RGBColor(0xCC, 0xD6, 0xE8)   # soft blue-white
TEXT_MUTED    = RGBColor(0x88, 0x99, 0xAA)
RULE_COLOR    = RGBColor(0xC0, 0x00, 0x00)
TAG_BG        = RGBColor(0x1A, 0x22, 0x3A)   # dark card bg
BULLET_DOT    = RGBColor(0xE8, 0x6C, 0x00)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]   # fully blank layout

# ══════════════════════════════════════════════════════════════════
# Helper utilities
# ══════════════════════════════════════════════════════════════════

def add_slide():
    return prs.slides.add_slide(BLANK)

def fill_bg(slide, color=BG_DARK):
    """Solid background rectangle covering the whole slide."""
    bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    return bg

def add_rect(slide, l, t, w, h, fill, line_color=None, line_width=0):
    s = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line_color:
        s.line.color.rgb = line_color
        s.line.width = Pt(line_width)
    else:
        s.line.fill.background()
    return s

def add_text(slide, text, l, t, w, h,
             font_size=18, bold=False, color=TEXT_WHITE,
             align=PP_ALIGN.LEFT, italic=False):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf  = txb.text_frame
    tf.word_wrap = True
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size  = Pt(font_size)
    run.font.bold  = bold
    run.font.color.rgb = color
    run.font.italic = italic
    return txb

def add_rule(slide, l, t, w, color=RULE_COLOR, thickness=2):
    line = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Pt(thickness))
    line.fill.solid()
    line.fill.fore_color.rgb = color
    line.line.fill.background()

def add_bullet_block(slide, items, l, t, w, h,
                     font_size=15, color=TEXT_LIGHT, dot_color=ACCENT_ORANGE,
                     line_spacing=0.38):
    """Render a list of strings as individual bullet rows."""
    y = t
    for item in items:
        # bullet dot
        dot = slide.shapes.add_shape(
            1,
            Inches(l),
            Inches(y + 0.05),
            Inches(0.12), Inches(0.12)
        )
        dot.fill.solid()
        dot.fill.fore_color.rgb = dot_color
        dot.line.fill.background()

        # text
        txb = slide.shapes.add_textbox(
            Inches(l + 0.2), Inches(y),
            Inches(w - 0.2), Inches(line_spacing + 0.05)
        )
        tf = txb.text_frame
        tf.word_wrap = True
        p  = tf.paragraphs[0]
        run = p.add_run()
        run.text = item
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
        y += line_spacing

def tag_box(slide, label, l, t, w=2.2, h=0.45,
            bg=ACCENT_RED, text_color=TEXT_WHITE, font_size=11):
    add_rect(slide, l, t, w, h, bg)
    add_text(slide, label, l, t + 0.04, w, h,
             font_size=font_size, bold=True, color=text_color,
             align=PP_ALIGN.CENTER)

def section_card(slide, title, items, l, t, w, h,
                 title_size=13, body_size=13.5):
    add_rect(slide, l, t, w, h, TAG_BG)
    add_rule(slide, l, t, w, ACCENT_RED, thickness=3)
    add_text(slide, title, l + 0.15, t + 0.08, w - 0.3, 0.35,
             font_size=title_size, bold=True, color=ACCENT_ORANGE)
    add_bullet_block(slide, items,
                     l + 0.15, t + 0.5, w - 0.3, h - 0.6,
                     font_size=body_size, line_spacing=0.36)

def slide_header(slide, title, subtitle=None,
                 title_y=0.28, sub_y=0.78, rule_y=0.72):
    add_text(slide, title,
             0.55, title_y, 12.2, 0.6,
             font_size=28, bold=True, color=TEXT_WHITE)
    add_rule(slide, 0.55, rule_y, 12.2)
    if subtitle:
        add_text(slide, subtitle,
                 0.55, sub_y, 12.2, 0.45,
                 font_size=13, color=ACCENT_ORANGE, italic=True)

# ══════════════════════════════════════════════════════════════════
# SLIDE 1 – Title
# ══════════════════════════════════════════════════════════════════
s1 = add_slide()
fill_bg(s1)

# left accent bar
add_rect(s1, 0, 0, 0.35, 7.5, ACCENT_RED)

# classification tag
tag_box(s1, "THREAT ACTOR PROFILE", 1.0, 0.6, 3.2, 0.42,
        bg=ACCENT_RED, font_size=11)

# main title
add_text(s1, "MuddyWater",
         1.0, 1.25, 10.5, 1.4,
         font_size=68, bold=True, color=TEXT_WHITE)

# rule
add_rule(s1, 1.0, 2.75, 11.2, ACCENT_RED, thickness=4)

# subtitle / aliases
add_text(s1, "AKA  Seedworm  |  Mango Sandstorm  |  Static Kitten",
         1.0, 2.95, 10.5, 0.5,
         font_size=18, color=ACCENT_ORANGE, bold=True)

# descriptor line
add_text(s1, "Iranian State-Sponsored Advanced Persistent Threat",
         1.0, 3.55, 10.5, 0.45,
         font_size=16, color=TEXT_LIGHT)

# bottom meta row
add_rule(s1, 1.0, 6.6, 11.2, TEXT_MUTED, thickness=1)
add_text(s1, "Classification: TLP:AMBER   |   Date: March 2026   |   MITRE ATT&CK Aligned",
         1.0, 6.7, 11.2, 0.4,
         font_size=11, color=TEXT_MUTED)

# ══════════════════════════════════════════════════════════════════
# SLIDE 2 – Identity & Strategic Motivation
# ══════════════════════════════════════════════════════════════════
s2 = add_slide()
fill_bg(s2)
add_rect(s2, 0, 0, 0.35, 7.5, ACCENT_RED)
slide_header(s2, "Identity & Strategic Motivation",
             subtitle="Who are they and what do they want?")

# three info cards
cards = [
    ("AFFILIATION",
     ["Iranian Ministry of Intelligence",
      "and Security (MOIS)",
      "Operational since ~2017"]),
    ("PRIMARY OBJECTIVE",
     ["Strategic Espionage",
      "Long-term network persistence",
      "Geopolitical intelligence gathering"]),
    ("TARGET SECTORS",
     ["Telecommunications & ISPs",
      "Government & Defence",
      "Middle East & United States"]),
]
for i, (title, items) in enumerate(cards):
    section_card(s2, title, items,
                 l=0.55 + i * 4.25, t=1.55, w=4.0, h=2.8,
                 title_size=13, body_size=13)

# key insight box
add_rect(s2, 0.55, 4.65, 12.4, 2.35, TAG_BG)
add_rule(s2, 0.55, 4.65, 12.4, ACCENT_ORANGE, thickness=3)
add_text(s2, "KEY INSIGHT",
         0.75, 4.72, 4, 0.4,
         font_size=12, bold=True, color=ACCENT_ORANGE)
add_text(s2,
         'MuddyWater does not hack for financial gain. They conduct long-term reconnaissance — '
         'staying "quiet" inside networks for months — to extract political, economic, and '
         'military intelligence aligned with Iran\'s foreign policy goals.',
         0.75, 5.1, 12.0, 1.6,
         font_size=14, color=TEXT_LIGHT)

# ══════════════════════════════════════════════════════════════════
# SLIDE 3 – Case Study: Operation Epic Fury
# ══════════════════════════════════════════════════════════════════
s3 = add_slide()
fill_bg(s3)
add_rect(s3, 0, 0, 0.35, 7.5, ACCENT_RED)
slide_header(s3, "Case Study — Operation Epic Fury",
             subtitle="High-Profile Breaches & Tactical Shift  |  March 2026")

# left column: incidents
section_card(s3,
             "HIGH-PROFILE INCIDENTS",
             ["FBI Director Kash Patel — personal email breach (Mar 27)",
              "Documents & photos leaked, confirmed by FBI via BBC",
              "U.S. Banking & Aerospace infrastructure targeted",
              "Major international airport systems compromised"],
             l=0.55, t=1.55, w=5.9, h=3.4,
             title_size=13, body_size=13)

# right column: new malware + shift
section_card(s3,
             "NEW MALWARE: DINOO / DINDOOR",
             ["JavaScript-based backdoor (2026 debut)",
              "Deployed via Deno runtime — evades AV",
              "Confirmed by Broadcom Symantec / Carbon Black",
              "Active against U.S. defence contractors"],
             l=6.75, t=1.55, w=6.0, h=1.85,
             title_size=13, body_size=12.5)

section_card(s3,
             "STRATEGIC SHIFT",
             ["From pure espionage → psychological influence ops",
              "Weaponising 'admin noise' (e.g. VOA/HMRC merger phishing)",
              "Targets distracted employees during org change"],
             l=6.75, t=3.6, w=6.0, h=1.35,
             title_size=13, body_size=12.5)

# bottom timeline bar
add_rect(s3, 0.55, 5.25, 12.4, 0.08, ACCENT_RED)
for i, (x, label) in enumerate([
    (0.65,  "Jan 2026\nInitial recon"),
    (3.5,   "Feb 2026\nDinoo deployed"),
    (7.0,   "Mar 15\nBank & Airport hit"),
    (10.2,  "Mar 27\nFBI Director breach"),
]):
    add_text(s3, label, x, 5.35, 2.5, 0.7,
             font_size=10, color=TEXT_MUTED, align=PP_ALIGN.LEFT)
    dot = s3.shapes.add_shape(
        1, Inches(x - 0.08), Inches(5.18), Inches(0.18), Inches(0.18))
    dot.fill.solid()
    dot.fill.fore_color.rgb = ACCENT_RED
    dot.line.fill.background()

# ══════════════════════════════════════════════════════════════════
# SLIDE 4 – SOC Detection
# ══════════════════════════════════════════════════════════════════
s4 = add_slide()
fill_bg(s4)
add_rect(s4, 0, 0, 0.35, 7.5, ACCENT_RED)
slide_header(s4, "SOC Detection — Advanced Indicators of Concern",
             subtitle="SFIA: SCAD — Security Operations")

detection_cards = [
    ("DINOO SIGNATURE",
     ["Execution via Deno.exe (JS runtime)",
      "Unusual child process from dev tool",
      "Obfuscated .js payload on disk"]),
    ("RMM ABUSE",
     ["Unauthorised SimpleHelp sessions",
      "Atera / ScreenConnect anomalies",
      "Admin tool → external IP egress"]),
    ("CLOUD HIJACKING",
     ["Malicious scripts via MS Intune",
      "Rogue 'Global Admin' account creation",
      "Entra ID sign-in from TOR / VPN"]),
    ("C2 TRAFFIC",
     ["Telegram Bot API calls from endpoints",
      "Dropbox / cloud storage exfil",
      "Beaconing on non-standard intervals"]),
]
for i, (title, items) in enumerate(detection_cards):
    col = i % 2
    row = i // 2
    section_card(s4, title, items,
                 l=0.55 + col * 6.2, t=1.55 + row * 2.4,
                 w=5.9, h=2.1,
                 title_size=12.5, body_size=12.5)

# priority flag
add_rect(s4, 0.55, 6.38, 12.4, 0.65, RGBColor(0x2A, 0x06, 0x06))
add_rule(s4, 0.55, 6.38, 12.4, ACCENT_RED, thickness=3)
add_text(s4,
         "PRIORITY INDICATOR — Admin tool (SimpleHelp/Atera) communicating with Telegram Bot API "
         "= IMMEDIATE HIGH-PRIORITY INCIDENT. Maps directly to MuddyWater playbook.",
         0.75, 6.45, 12.0, 0.55,
         font_size=12, bold=True, color=RGBColor(0xFF, 0x80, 0x80))

# ══════════════════════════════════════════════════════════════════
# SLIDE 5 – Recognition & Attribution
# ══════════════════════════════════════════════════════════════════
s5 = add_slide()
fill_bg(s5)
add_rect(s5, 0, 0, 0.35, 7.5, ACCENT_RED)
slide_header(s5, "Recognition & Attribution",
             subtitle="SFIA: TECH — Specialist Advice  |  MITRE ATT&CK Mapping")

# left: TTP list
section_card(s5,
             "KNOWN TACTICS, TECHNIQUES & PROCEDURES",
             ["PowerShell-heavy lateral movement (T1059.001)",
              "Ligolo-ng for internal network tunnelling",
              "LSASS credential dumping (T1003.001)",
              "Spearphishing with lure documents (T1566.001)",
              "Living-off-the-Land — no custom malware in early stages",
              "Persistent RMM implant for long-haul access"],
             l=0.55, t=1.55, w=6.1, h=4.5,
             title_size=13, body_size=13.5)

# right: MITRE tiles
add_text(s5, "MITRE ATT&CK TECHNIQUES",
         7.0, 1.55, 5.8, 0.4,
         font_size=13, bold=True, color=ACCENT_ORANGE)
add_rule(s5, 7.0, 1.98, 5.8, ACCENT_RED)

mitre = [
    ("T1059.001", "PowerShell Execution"),
    ("T1572",     "Protocol Tunnelling (Ligolo-ng)"),
    ("T1078",     "Valid Accounts — RMM"),
    ("T1071.001", "Web Protocols — Telegram C2"),
    ("T1003.001", "OS Credential Dumping"),
    ("T1566.001", "Spearphishing Attachment"),
    ("T1021.001", "Remote Desktop Protocol"),
]
for i, (tid, tdesc) in enumerate(mitre):
    y = 2.1 + i * 0.52
    add_rect(s5, 7.0, y, 1.6, 0.4, ACCENT_RED)
    add_text(s5, tid, 7.0, y + 0.04, 1.6, 0.38,
             font_size=11, bold=True, color=TEXT_WHITE, align=PP_ALIGN.CENTER)
    add_text(s5, tdesc, 8.7, y + 0.04, 4.1, 0.38,
             font_size=12, color=TEXT_LIGHT)

# attribution insight
add_rect(s5, 0.55, 6.25, 12.4, 0.78, TAG_BG)
add_rule(s5, 0.55, 6.25, 12.4, ACCENT_ORANGE, thickness=2)
add_text(s5,
         "High-confidence attribution enables predictive defence — if we see PowerShell + Ligolo-ng + "
         "SimpleHelp, we know their next target will be database servers. We move to protect them proactively.",
         0.75, 6.33, 12.0, 0.65,
         font_size=13, color=TEXT_LIGHT, italic=True)

# ══════════════════════════════════════════════════════════════════
# SLIDE 6 – Tactical Response & Containment
# ══════════════════════════════════════════════════════════════════
s6 = add_slide()
fill_bg(s6)
add_rect(s6, 0, 0, 0.35, 7.5, ACCENT_RED)
slide_header(s6, "Tactical Response & Containment",
             subtitle="SFIA: USUP — Incident Management  |  Speed of Attacker = Speed of Response")

response_steps = [
    ("1  IDENTITY-FIRST",
     ["Revoke ALL session tokens immediately",
      "Force password resets for admin accounts",
      "Disable compromised service principals"]),
    ("2  ENDPOINT LOCKDOWN",
     ["Quarantine Deno.exe runtime",
      "Isolate host from network segment",
      "Kill & remove Dindoor JS backdoor"]),
    ("3  CLOUD CONTROL",
     ["Audit Intune for rogue scripts",
      "Remove unauthorised Global Admin roles",
      "Review Entra ID Conditional Access logs"]),
    ("4  NETWORK / C2 BLOCK",
     ["Block Telegram API at perimeter firewall",
      "Null-route known Dinoo/Dindoor domains",
      "Disable unauthorised RMM (SimpleHelp)"]),
]
for i, (title, items) in enumerate(response_steps):
    col = i % 2
    row = i // 2
    section_card(s6, title, items,
                 l=0.55 + col * 6.2, t=1.55 + row * 2.25,
                 w=5.9, h=2.0,
                 title_size=13, body_size=13)

# bottom note
add_rect(s6, 0.55, 6.18, 12.4, 0.85, TAG_BG)
add_rule(s6, 0.55, 6.18, 12.4, ACCENT_RED, thickness=2)
add_text(s6,
         "STRYKER PRECEDENT — MuddyWater used Microsoft Intune to wipe systems company-wide. "
         "Revoking session tokens locks the attacker out of the cloud console instantly before "
         "a single wipe command can propagate.",
         0.75, 6.26, 12.0, 0.72,
         font_size=12, color=RGBColor(0xFF, 0xCC, 0x88))

# ══════════════════════════════════════════════════════════════════
# SLIDE 7 – Recovery & Stakeholder Support
# ══════════════════════════════════════════════════════════════════
s7 = add_slide()
fill_bg(s7)
add_rect(s7, 0, 0, 0.35, 7.5, ACCENT_RED)
slide_header(s7, "Recovery & Stakeholder Support",
             subtitle="SFIA: SLMO / RLMT — Service Level & Stakeholder Management")

# three pillars
pillars = [
    ("ERADICATION",
     ["Threat hunt for 'sleeping' backdoors",
      "Full forensic sweep of all endpoints",
      "Validate clean re-images via hashing",
      "Confirm no persistence in scheduled tasks",
      "Verify cloud config post-remediation"]),
    ("CREDENTIAL HARDENING",
     ["Enforce MFA resets org-wide",
      "Rotate all privileged service accounts",
      "Implement Conditional Access policies",
      "Review & tighten admin role assignments",
      "Enable Entra ID Identity Protection"]),
    ("STAKEHOLDER SUPPORT",
     ["Clear, jargon-free comms to impacted teams",
      "Explain WHY security measures are needed",
      "Just-in-Time phishing simulation training",
      "Turn breach into a learning opportunity",
      "Strengthen SOC–business relationship"]),
]
for i, (title, items) in enumerate(pillars):
    section_card(s7, title, items,
                 l=0.55 + i * 4.25, t=1.55, w=4.0, h=4.5,
                 title_size=13, body_size=13)

# bottom quote
add_rect(s7, 0.55, 6.3, 12.4, 0.72, TAG_BG)
add_rule(s7, 0.55, 6.3, 12.4, ACCENT_ORANGE, thickness=2)
add_text(s7,
         '"Recovery is more than a reboot. It is rebuilding trust — with the network, with users, '
         'and with the business."',
         0.75, 6.38, 12.0, 0.58,
         font_size=14, color=TEXT_LIGHT, italic=True)

# ══════════════════════════════════════════════════════════════════
# SLIDE 8 – Conclusion
# ══════════════════════════════════════════════════════════════════
s8 = add_slide()
fill_bg(s8)
add_rect(s8, 0, 0, 0.35, 7.5, ACCENT_RED)

# large section title
add_text(s8, "CONCLUSION",
         0.55, 0.3, 5.0, 0.5,
         font_size=13, bold=True, color=ACCENT_RED, italic=True)
add_text(s8, "Key Takeaways",
         0.55, 0.75, 12.0, 0.75,
         font_size=32, bold=True, color=TEXT_WHITE)
add_rule(s8, 0.55, 1.55, 12.2, ACCENT_RED, thickness=3)

# three takeaway boxes
takeaways = [
    ("ADAPTIVE & STATE-ALIGNED",
     "MuddyWater evolves its tooling — from simple PowerShell to custom JS "
     "backdoors — but its espionage mission never changes."),
    ("RMM ABUSE IS THE KILL CHAIN",
     "Early detection of unauthorised RMM tools (SimpleHelp, Atera) "
     "communicating with Telegram is the single highest-value detection."),
    ("PROACTIVE SOC WINS",
     "Attribution-driven threat hunting and strong user relationships let "
     "the SOC predict attacker movement and protect assets before they're hit."),
]
for i, (title, body) in enumerate(takeaways):
    l = 0.55 + i * 4.25
    add_rect(s8, l, 1.85, 4.0, 3.6, TAG_BG)
    add_rule(s8, l, 1.85, 4.0, ACCENT_RED, thickness=3)
    # number
    num = s8.shapes.add_shape(
        1, Inches(l + 0.15), Inches(1.98), Inches(0.5), Inches(0.5))
    num.fill.solid()
    num.fill.fore_color.rgb = ACCENT_RED
    num.line.fill.background()
    add_text(s8, str(i + 1),
             l + 0.15, 1.98, 0.5, 0.5,
             font_size=15, bold=True, color=TEXT_WHITE, align=PP_ALIGN.CENTER)
    add_text(s8, title,
             l + 0.75, 2.0, 3.0, 0.5,
             font_size=12, bold=True, color=ACCENT_ORANGE)
    add_text(s8, body,
             l + 0.2, 2.6, 3.65, 2.6,
             font_size=13, color=TEXT_LIGHT)

# closing statement
add_rect(s8, 0.55, 5.75, 12.4, 1.0, RGBColor(0x12, 0x04, 0x04))
add_rule(s8, 0.55, 5.75, 12.4, ACCENT_RED, thickness=2)
add_text(s8,
         "MuddyWater uses our own tools against us. A SOC that monitors for unusual behaviour "
         "and maintains strong relationships with its users can effectively neutralise this threat.",
         0.75, 5.85, 12.0, 0.82,
         font_size=14, color=TEXT_WHITE, italic=True)

# bottom bar
add_rule(s8, 0.55, 7.05, 12.2, TEXT_MUTED, thickness=1)
add_text(s8, "Classification: TLP:AMBER   |   March 2026   |   MITRE ATT&CK Aligned",
         0.55, 7.1, 12.2, 0.35,
         font_size=10, color=TEXT_MUTED, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════
# Save
# ══════════════════════════════════════════════════════════════════
OUT = "/home/user/AI-Projects/MuddyWater_Threat_Profile.pptx"
prs.save(OUT)
print(f"Saved: {OUT}")
