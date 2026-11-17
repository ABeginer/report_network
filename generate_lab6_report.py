#!/usr/bin/env python3
"""
Lab 6 Report Generator - CSRF Attacks and Countermeasures
Following strict formatting requirements:
- LEFT alignment only (no JUSTIFY)
- Each paragraph as separate doc.add_paragraph() calls
- No empty section headers
- Code blocks in gray tables with Consolas font
- Images at dpi=200, centered with captions
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import FancyArrowPatch

# Paths
OUTPUT_DIR = "/workspace/report_network"
IMAGES_DIR = os.path.join(OUTPUT_DIR, "lab6_images")
os.makedirs(IMAGES_DIR, exist_ok=True)

def create_document():
    """Create a new Word document with proper styling."""
    doc = Document()
    
    # Define styles - all LEFT aligned
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    return doc

def add_title_page(doc):
    """Add title page with student information."""
    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("LABORATORY REPORT")
    run.bold = True
    run.font.size = Pt(16)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Lab 6: Web CSRF Attacks and Countermeasures")
    run.bold = True
    run.font.size = Pt(14)
    
    # Student info table
    table = doc.add_table(rows=5, cols=2)
    table.style = 'Table Grid'
    for row in table.rows:
        for cell in row.cells:
            cell.width = Inches(3.5)
    
    info = [
        ("Student Name:", "Nguyễn Văn A"),
        ("Student ID:", "20261117"),
        ("Course:", "Network Security Laboratory"),
        ("Instructor:", "Dr. Wenliang Du"),
        ("Date:", "November 2024")
    ]
    
    for i, (label, value) in enumerate(info):
        cell = table.cell(i, 0)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = cell.paragraphs[0].add_run(label)
        run.bold = True
        
        cell = table.cell(i, 1)
        run = cell.paragraphs[0].add_run(value)
    
    return doc

def add_section_header(doc, level, title):
    """Add section header with proper formatting."""
    if level == 1:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(14)
        run.font.name = 'Calibri'
    elif level == 2:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(12)
        run.font.name = 'Calibri'
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(11)
        run.font.name = 'Calibri'
    
    return doc

def add_code_block(doc, code_text, language=""):
    """Add code block in gray table with Consolas font."""
    # Create 1x1 table
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    table.width = Inches(6.5)
    
    # Set background color to light gray (#F4F4F4)
    cell = table.cell(0, 0)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Apply gray background
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), 'F4F4F4')
    tcPr.append(shd)
    
    # Add code text with Consolas font
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    if language:
        run = p.add_run(f"# {language}\n")
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(100, 100, 100)
    
    # Split code into lines and add each line
    for line in code_text.strip().split('\n'):
        run = p.add_run(line + '\n')
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0, 0, 0)
    
    return doc

def add_figure(doc, image_path, caption):
    """Add figure with caption below."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Figure {len([p for p in doc.paragraphs if 'Figure' in p.text]) + 1}: ")
    run.font.italic = True
    run.font.name = 'Calibri'
    
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(caption)
    run2.font.italic = True
    run2.font.name = 'Calibri'
    
    # Add image centered
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    img_run = p_img.add_run()
    img_run.add_picture(image_path, width=Inches(5.8))
    
    return doc

def draw_csrf_attack_diagram():
    """Draw CSRF attack diagram."""
    fig, ax = plt.subplots(figsize=(8, 6), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # Draw boxes
    colors = ['#E3F2FD', '#FFF3E0', '#F3E5F5']
    positions = [(20, 70), (50, 70), (80, 70)]
    labels = ['Victim Browser', 'Attacker Site', 'Target Site (Elgg)']
    
    for pos, color, label in zip(positions, colors, labels):
        x, y = pos
        rect = plt.Rectangle([x-15, y-20], 30, 40, facecolor=color, edgecolor='black', linewidth=2)
        ax.add_patch(rect)
        ax.text(x, y+10, label, ha='center', va='center', fontsize=9, fontweight='bold')
    
    # Draw arrows for CSRF attack using FancyArrowPatch
    arrow = FancyArrowPatch((35, 85), (65, 85), arrowstyle='->', color='red', lw=2)
    ax.add_patch(arrow)
    ax.text(50, 92, 'Malicious Link', ha='center', fontsize=8, color='red')
    
    arrow2 = FancyArrowPatch((65, 85), (95, 85), arrowstyle='->', color='red', lw=2)
    ax.add_patch(arrow2)
    ax.text(80, 92, 'Forged Request', ha='center', fontsize=8, color='red')
    
    plt.title('CSRF Attack Flow', fontsize=10, fontweight='bold')
    plt.tight_layout()
    path = os.path.join(IMAGES_DIR, "csrf_attack_diagram.png")
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    return path

def draw_request_comparison():
    """Draw GET vs POST request comparison."""
    fig, axes = plt.subplots(1, 2, figsize=(10, 5), dpi=200)
    
    # GET Request
    ax1 = axes[0]
    ax1.set_xlim(0, 100)
    ax1.set_ylim(0, 80)
    ax1.axis('off')
    
    rect = plt.Rectangle([10, 20], 70, 50, facecolor='#E3F2FD', edgecolor='black', linewidth=2)
    ax1.add_patch(rect)
    ax1.text(45, 60, 'GET Request', ha='center', va='center', fontsize=10, fontweight='bold')
    ax1.text(45, 45, 'URL: /friend/add?friend_id=123', ha='center', fontsize=8)
    
    # POST Request
    ax2 = axes[1]
    ax2.set_xlim(0, 100)
    ax2.set_ylim(0, 80)
    ax2.axis('off')
    
    rect = plt.Rectangle([10, 20], 70, 50, facecolor='#FFF3E0', edgecolor='black', linewidth=2)
    ax2.add_patch(rect)
    ax2.text(45, 60, 'POST Request', ha='center', va='center', fontsize=10, fontweight='bold')
    ax2.text(45, 45, 'Body: friend_id=123\nCSRF Token: abc123', ha='center', fontsize=8)
    
    plt.suptitle('GET vs POST Request Comparison', fontsize=12, fontweight='bold')
    plt.tight_layout()
    path = os.path.join(IMAGES_DIR, "request_comparison.png")
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    return path

def draw_samesite_diagram():
    """Draw SameSite cookie behavior diagram."""
    fig, ax = plt.subplots(figsize=(10, 6), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # Three columns for SameSite modes
    modes = ['None', 'Lax', 'Strict']
    colors = ['#FFEbee', '#E8F5E9', '#FFFDE7']
    descriptions = [
        'Cookies sent on all requests\n(Vulnerable to CSRF)',
        'Cookies sent on same-site only\nand top-level navigations',
        'Cookies sent only on same-site requests'
    ]
    
    for i, (mode, color, desc) in enumerate(zip(modes, colors, descriptions)):
        x = 20 + i * 25
        rect = plt.Rectangle([x-18, 30], 36, 50, facecolor=color, edgecolor='black', linewidth=2)
        ax.add_patch(rect)
        ax.text(x, 75, mode, ha='center', va='center', fontsize=10, fontweight='bold')
        ax.text(x, 45, desc, ha='center', fontsize=7, linespacing=1.3)
    
    plt.title('SameSite Cookie Behavior Comparison', fontsize=12, fontweight='bold')
    plt.tight_layout()
    path = os.path.join(IMAGES_DIR, "samesite_diagram.png")
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    return path

def generate_report():
    """Generate the complete Lab 6 report."""
    doc = create_document()
    
    # Title page
    add_title_page(doc)
    doc.add_paragraph()  # Empty line after title
    
    # Section 1: Introduction
    add_section_header(doc, 1, "1. Introduction")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("Cross-Site Request Forgery (CSRF) is a web security vulnerability that allows attackers to induce users to perform actions they did not intend to perform. This lab explores CSRF attacks using the Elgg social networking platform and examines various countermeasures.")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("The primary goal of this laboratory is to understand how CSRF attacks exploit active user sessions and to learn effective defense mechanisms including CSRF tokens and SameSite cookie attributes.")
    
    # Section 2: Lab Setup
    add_section_header(doc, 1, "2. Laboratory Environment Setup")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("The lab environment uses Docker containers to host three distinct websites:")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- Vulnerable Elgg site (www.seed-server.com)")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- Attacker's malicious site (www.attacker32.com)")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- Defense demonstration site (www.example32.com)")
    
    # Add code block for setup commands
    add_section_header(doc, 2, "2.1 Setup Commands")
    
    code = """alias dcbuild='docker-compose build'
alias dcdown='docker-compose down'
alias dcup='docker-compose up'
alias dockps='docker ps --format "{{.ID}}  {{.Names}}"'

# Host configuration in /etc/hosts:
10.9.0.5 www.seed-server.com
10.9.0.5 www.example32.com
10.9.0.105 www.attacker32.com"""
    
    add_code_block(doc, code, "Bash")
    
    # Section 3: CSRF Attacks
    add_section_header(doc, 1, "3. CSRF Attack Experiments")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("This section documents the three main attack vectors explored during this laboratory exercise.")
    
    # Task 1: Observing HTTP Requests
    add_section_header(doc, 2, "3.1 Task 1: Observing HTTP Requests")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("Using Firefox add-on HTTP Header Live and browser developer tools, we examined the GET and POST requests made by the Elgg platform.")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("Key observations:")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- Friend request endpoint: /friend/add")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- User ID parameter: user_id in GET requests")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- Friend identifier: friend_id as the target of action")
    
    # Add diagram
    img_path = draw_csrf_attack_diagram()
    add_figure(doc, img_path, "CSRF Attack Flow Diagram showing victim browser, attacker site, and target site interaction")
    
    # Task 2: CSRF via GET Request
    add_section_header(doc, 2, "3.2 Task 2: CSRF via GET Request")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("The first attack vector exploits the fact that GET requests can be triggered silently through HTML img tags.")
    
    code = """<!-- Malicious HTML on attacker site -->
<img src="http://www.seed-server.com/friend/add?friend_id=123" 
     width="0" height="0" style="display:none;">"""
    
    add_code_block(doc, code, "HTML")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("When a victim visits the attacker's site while logged into Elgg, the browser automatically sends the GET request with the victim's session cookie, adding friend_id=123 to their friends list without their knowledge.")
    
    # Task 3: CSRF via POST Request
    add_section_header(doc, 2, "3.3 Task 3: CSRF via POST Request")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("POST-based CSRF attacks require form submission and can be automated using JavaScript.")
    
    code = """<!-- Malicious HTML with auto-submitting form -->
<form id="csrf_form" action="http://www.seed-server.com/user/edit" method="POST">
  <input type="hidden" name="fullname" value="Hacked User">
  <input type="hidden" name="email" value="attacker@evil.com">
</form>

<script>
  document.getElementById('csrf_form').submit();
</script>"""
    
    add_code_block(doc, code, "HTML/JavaScript")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("The JavaScript automatically submits the forged POST request when the page loads, modifying the victim's profile without their consent.")
    
    # Add comparison diagram
    img_path2 = draw_request_comparison()
    add_figure(doc, img_path2, "GET vs POST Request Comparison showing different attack vectors")
    
    # Section 4: CSRF Defenses
    add_section_header(doc, 1, "4. CSRF Defense Mechanisms")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("This section examines the countermeasures implemented to protect against CSRF attacks.")
    
    # Task 4: CSRF Token Protection
    add_section_header(doc, 2, "4.1 Task 4: Enabling Elgg's CSRF Countermeasures")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("Elgg implements token-based CSRF protection where each form includes a unique, unpredictable CSRF token that must be validated by the server.")
    
    code = """<!-- Legitimate Elgg form with CSRF token -->
<form action="/user/edit" method="POST">
  <input type="hidden" name="csrf_token" value="a1b2c3d4e5f6g7h8">
  <input type="text" name="fullname" placeholder="Full Name">
  <button type="submit">Save Profile</button>
</form>"""
    
    add_code_block(doc, code, "HTML")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("When CSRF protection is enabled:")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- The server generates a unique token per session")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- Token is embedded in all forms and AJAX requests")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- Attacker cannot forge valid tokens without access to victim's session")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- Server rejects requests with missing or invalid tokens")
    
    # Task 5: SameSite Cookie Attribute
    add_section_header(doc, 2, "4.2 Task 5: Experimenting with SameSite Cookies")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("The SameSite cookie attribute provides an additional layer of CSRF protection by controlling when cookies are sent with cross-site requests.")
    
    code = """# Cookie with different SameSite values:

Set-Cookie: sessionid=abc123; SameSite=None  # Vulnerable to CSRF
Set-Cookie: sessionid=abc123; SameSite=Lax  # Partial protection
Set-Cookie: sessionid=abc123; SameSite=Strict  # Maximum protection"""
    
    add_code_block(doc, code, "HTTP Headers")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("SameSite cookie behavior analysis:")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- SameSite=None: Cookies sent on all cross-site requests (vulnerable)")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- SameSite=Lax: Cookies sent only on same-site and top-level navigations")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- SameSite=Strict: Cookies sent only on same-site requests (maximum protection)")
    
    # Add SameSite diagram
    img_path3 = draw_samesite_diagram()
    add_figure(doc, img_path3, "SameSite Cookie Behavior Comparison showing None, Lax, and Strict modes")
    
    # Section 5: Conclusion
    add_section_header(doc, 1, "5. Conclusion")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("This laboratory exercise demonstrated the mechanics of CSRF attacks and their potential impact on web applications. Through hands-on experimentation with the Elgg platform, we observed how attackers can exploit active user sessions to perform unauthorized actions.")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("Key findings from this lab:")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- CSRF attacks can be executed via both GET and POST requests")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- HTML img tags enable silent GET-based CSRF attacks")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- JavaScript can automate POST-based CSRF attacks through form submission")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- Token-based CSRF protection effectively blocks forged requests when properly implemented")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("- SameSite cookie attributes provide defense-in-depth against CSRF attacks")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("The combination of CSRF tokens and appropriate SameSite cookie settings represents the current best practice for CSRF mitigation in modern web applications.")
    
    return doc

def main():
    """Main function to generate and save the report."""
    print("Generating Lab 6 Report...")
    
    doc = generate_report()
    
    output_path = os.path.join(OUTPUT_DIR, "Lab6_Report_Solution.docx")
    doc.save(output_path)
    
    print(f"Report saved to: {output_path}")
    print(f"File size: {os.path.getsize(output_path)} bytes")

if __name__ == "__main__":
    main()
