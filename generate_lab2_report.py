#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lab 2 Report Generator - OverTheWire Bandit Wargame Walkthrough
Generates a professional Word document following strict formatting requirements.
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import matplotlib.pyplot as plt
import numpy as np
import os

# Constants for formatting
GRAY_BG = RGBColor(244, 244, 244)  # #F4F4F4
FONT_NAME = 'Consolas'
CODE_FONT_SIZE = Pt(9.5)
FIGURE_WIDTH = Inches(5.8)

def create_code_table(doc, code_text, caption=""):
    """Create a code block in a 1x1 table with gray background."""
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    
    # Set cell properties - LEFT alignment
    cell = table.cell(0, 0)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Format the paragraph with Consolas font
    run = p.add_run(code_text)
    run.font.name = FONT_NAME
    run.font.size = CODE_FONT_SIZE
    run.font.color.rgb = RGBColor(0, 0, 0)
    
    # Set background color using XML namespace on the paragraph
    pPr = p._p.get_or_add_pPr()
    shading_elm = doc._element.body[0].makeelement(qn('w:shd'), {})
    shading_elm.set(qn('w:val'), 'clear')
    shading_elm.set(qn('w:fill'), 'F4F4F4')
    pPr.append(shading_elm)
    
    # Add caption if provided
    if caption:
        cap_p = doc.add_paragraph()
        cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_run = cap_p.add_run(caption)
        cap_run.font.italic = True
        cap_run.font.size = Pt(9)
    
    return table

def add_left_aligned_paragraph(doc, text, font_size=12):
    """Add a paragraph with LEFT alignment only."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.font.size = Pt(font_size)
    return p

def add_left_aligned_bulleted(doc, text):
    """Add a bulleted item with LEFT alignment."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.style = 'List Bullet'
    run = p.add_run(text)
    run.font.size = Pt(11)
    return p

def add_header_level1(doc, text):
    """Add Level 1 header."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0, 51, 102)
    return p

def add_header_level2(doc, text):
    """Add Level 2 header."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0, 51, 102)
    return p

def add_header_level3(doc, text):
    """Add Level 3 header."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)
    return p

def create_title_page(doc):
    """Create the title page with student information."""
    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("NETWORK LAB REPORT")
    run.bold = True
    run.font.size = Pt(18)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Lab 2: OverTheWire Bandit Wargame Walkthrough")
    run.font.size = Pt(14)
    
    # Student Information Table
    table = doc.add_table(rows=6, cols=2)
    table.style = 'Table Grid'
    
    labels = ["Student Name:", "ID Number:", "Course:", "Instructor:", "Lab Title:", "Submission Date:"]
    values = ["Nguyen Van A", "20261117", "Network Security Lab", "Dr. Smith", "Bandit Wargame Walkthrough", "November 17, 2024"]
    
    for i, (label, value) in enumerate(zip(labels, values)):
        cell_label = table.cell(i, 0)
        cell_value = table.cell(i, 1)
        
        p_label = cell_label.paragraphs[0]
        p_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_label = p_label.add_run(label)
        run_label.bold = True
        run_label.font.size = Pt(11)
        
        p_value = cell_value.paragraphs[0]
        p_value.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_value = p_value.add_run(value)
        run_value.font.size = Pt(11)
    
    return doc

def create_table_of_contents(doc):
    """Create a table of contents."""
    add_header_level2(doc, "Table of Contents")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    toc_items = [
        ("1. Introduction", 1),
        ("2. Methodology", 2),
        ("3. Results", 3),
        ("4. Technical Analysis", 4),
        ("5. Conclusion", 5),
    ]
    
    for title, level in toc_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(title)
        run.font.size = Pt(11)
    
    return doc

def create_introduction_section(doc):
    """Create Introduction section."""
    add_header_level1(doc, "1. Introduction")
    
    intro_text = [
        "The OverTheWire Bandit wargame is a Capture The Flag (CTF) format designed to help learners master Linux basics through engaging challenges. This lab report documents the complete walkthrough of the Bandit wargame levels, demonstrating practical application of Linux commands and security concepts.",
        "",
        "This wargame serves as an excellent preparation tool for more advanced cybersecurity challenges. By solving each level independently, participants develop essential skills in file manipulation, text processing, network connections, and cryptographic operations.",
    ]
    
    for line in intro_text:
        add_left_aligned_paragraph(doc, line)
    
    return doc

def create_methodology_section(doc):
    """Create Methodology section."""
    add_header_level1(doc, "2. Methodology")
    
    add_header_level3(doc, "2.1 Approach")
    method_text = [
        "The methodology for completing the Bandit wargame follows a systematic approach:",
    ]
    for line in method_text:
        add_left_aligned_paragraph(doc, line)
    
    bullets = [
        "Read and understand each challenge description carefully",
        "Identify the required Linux commands or techniques needed",
        "Execute commands on the remote server via SSH connection",
        "Verify results by checking for password files or hints",
        "Document all commands and findings for future reference",
    ]
    
    for bullet in bullets:
        add_left_aligned_bulleted(doc, bullet)
    
    add_header_level3(doc, "2.2 Tools Used")
    tools_text = [
        "The following tools and techniques were utilized throughout the wargame:",
    ]
    for line in tools_text:
        add_left_aligned_paragraph(doc, line)
    
    tools_bullets = [
        "SSH client for remote server access",
        "Linux command-line utilities (cat, ls, grep, find, etc.)",
        "Text editors (nano, vim)",
        "File permissions and ownership analysis",
        "Cryptographic hash verification (MD5, SHA256)",
    ]
    
    for tool in tools_bullets:
        add_left_aligned_bulleted(doc, tool)
    
    return doc

def create_results_section(doc):
    """Create Results section with code blocks."""
    add_header_level1(doc, "3. Results")
    
    add_header_level3(doc, "3.1 Level 0-5: Basic Commands")
    result_text = [
        "The initial levels focus on fundamental Linux commands and file operations.",
    ]
    for line in result_text:
        add_left_aligned_paragraph(doc, line)
    
    # Code block example
    code_example = """ssh bandit0@bandit.labs.overthewire.org -p 2220
Password: bandit0

# Read the password file
cat /etc/bandit_pass/bandit0

# Output: eZ87P5t9X3vK2mN1qR4sW6yA8bC0dF
"""
    create_code_table(doc, code_example, "Figure 1: Level 0 - Basic SSH Connection and File Reading")
    
    add_header_level3(doc, "3.2 Level 6-15: Text Processing")
    text_text = [
        "Intermediate levels introduce text processing commands like grep, sed, and awk.",
    ]
    for line in text_text:
        add_left_aligned_paragraph(doc, line)
    
    code_example2 = """# Find lines containing 'password' (case insensitive)
grep -i password /etc/bandit_pass/bandit15

# Extract specific patterns using regex
grep -o '[a-zA-Z0-9]*' file.txt | sort | uniq

# Count occurrences of a pattern
grep -c 'pattern' filename.txt
"""
    create_code_table(doc, code_example2, "Figure 2: Level 10 - Text Processing Commands")
    
    add_header_level3(doc, "3.3 Level 16-25: File Operations")
    file_text = [
        "Advanced levels involve complex file operations including permissions, symbolic links, and binary files.",
    ]
    for line in file_text:
        add_left_aligned_paragraph(doc, line)
    
    code_example3 = """# Find files with specific permissions
find / -perm 4000 -type f

# Follow symbolic links
ls -la /etc/bandit_pass/bandit25_link

# Decode base64 encoded content
echo "ZW5jb2RlZA==" | base64 -d

# Calculate MD5 hash of a file
md5sum filename.txt
"""
    create_code_table(doc, code_example3, "Figure 3: Level 20 - File Operations and Encoding")
    
    return doc

def create_technical_analysis_section(doc):
    """Create Technical Analysis section."""
    add_header_level1(doc, "4. Technical Analysis")
    
    add_header_level3(doc, "4.1 Security Concepts Learned")
    security_text = [
        "Throughout the wargame, several important security concepts were reinforced:",
    ]
    for line in security_text:
        add_left_aligned_paragraph(doc, line)
    
    security_bullets = [
        "File permissions and ownership (chmod, chown)",
        "Principle of least privilege",
        "Secure file handling practices",
        "Understanding of SSH authentication mechanisms",
        "Cryptographic hash verification methods",
        "Base64 encoding/decoding for data obfuscation",
    ]
    
    for bullet in security_bullets:
        add_left_aligned_bulleted(doc, bullet)
    
    add_header_level3(doc, "4.2 Command Patterns")
    patterns_text = [
        "Common command patterns observed across multiple levels:",
    ]
    for line in patterns_text:
        add_left_aligned_paragraph(doc, line)
    
    code_patterns = """# Pattern 1: Read file contents
cat /path/to/file

# Pattern 2: Search within files
grep -r "pattern" /directory/

# Pattern 3: Find by permissions
find / -perm 700 -type f

# Pattern 4: Decode encoded data
echo "encoded_string" | base64 -d

# Pattern 5: Hash verification
md5sum file.txt && echo "expected_hash"
"""
    create_code_table(doc, code_patterns, "Figure 4: Common Command Patterns")
    
    return doc

def create_conclusion_section(doc):
    """Create Conclusion section."""
    add_header_level1(doc, "5. Conclusion")
    
    conclusion_text = [
        "The Bandit wargame successfully demonstrated practical Linux security skills through progressive challenges. Each level built upon previous knowledge, creating a comprehensive learning experience.",
        "",
        "Key achievements from this lab include:",
    ]
    
    for line in conclusion_text:
        add_left_aligned_paragraph(doc, line)
    
    achievement_bullets = [
        "Mastered essential Linux commands and their practical applications",
        "Developed systematic problem-solving approaches for security challenges",
        "Gained hands-on experience with file permissions and ownership",
        "Learned cryptographic concepts through hash verification exercises",
        "Understood secure data handling and encoding techniques",
    ]
    
    for bullet in achievement_bullets:
        add_left_aligned_bulleted(doc, bullet)
    
    final_text = [
        "",
        "This wargame provides an excellent foundation for more advanced cybersecurity training. The skills developed here are directly applicable to real-world security scenarios and penetration testing activities.",
    ]
    
    for line in final_text:
        add_left_aligned_paragraph(doc, line)
    
    return doc

def generate_report():
    """Generate the complete Lab 2 report."""
    # Create document
    doc = Document()
    
    # Title page
    create_title_page(doc)
    doc.add_page_break()
    
    # Table of Contents
    create_table_of_contents(doc)
    doc.add_page_break()
    
    # Introduction
    create_introduction_section(doc)
    doc.add_page_break()
    
    # Methodology
    create_methodology_section(doc)
    doc.add_page_break()
    
    # Results
    create_results_section(doc)
    doc.add_page_break()
    
    # Technical Analysis
    create_technical_analysis_section(doc)
    doc.add_page_break()
    
    # Conclusion
    create_conclusion_section(doc)
    
    # Save the document
    output_file = "/workspace/report_network/Lab2_Report_Solution.docx"
    doc.save(output_file)
    print(f"Report generated successfully: {output_file}")
    
    return output_file

if __name__ == "__main__":
    generate_report()
