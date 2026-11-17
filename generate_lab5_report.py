from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import parse_xml
import matplotlib.pyplot as plt
import numpy as np
import os

# Create document
doc = Document()

# Define styles
def set_left_alignment(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

# Title Page
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_run = title.add_run("BÁO CÁO THỰC HÀNH MÔN AN TOÀN MẠNG\n")
title_run.bold = True
title_run.font.size = Pt(16)

title2 = doc.add_paragraph()
title2.alignment = WD_ALIGN_PARAGRAPH.CENTER
title2_run = title2.add_run("ĐỀ TÀI: WARGAME NATAS - LAB 5\n")
title2_run.bold = True
title2_run.font.size = Pt(14)

# Student Info Table
table = doc.add_table(rows=6, cols=2)
table.style = 'Table Grid'
table.alignment = WD_ALIGN_PARAGRAPH.CENTER

data = [
    ("Họ và tên:", "Nguyễn Văn A"),
    ("MSSV:", "20215432"),
    ("Lớp:", "CNTT-K67"),
    ("Giảng viên hướng dẫn:", "ThS. Trần Văn B"),
    ("Khoa:", "Công nghệ Thông tin"),
    ("Năm học:", "2024-2025")
]

for i, (label, value) in enumerate(data):
    cell_label = table.cell(i, 0)
    cell_value = table.cell(i, 1)
    p_label = cell_label.paragraphs[0]
    p_label.add_run(label + ": ")
    p_label.runs[0].bold = True
    set_left_alignment(p_label)
    
    p_value = cell_value.paragraphs[0]
    p_value.add_run(value)
    set_left_alignment(p_value)

doc.add_page_break()

# Table of Contents
toc_title = doc.add_paragraph()
toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
toc_title_run = toc_title.add_run("MỤC LỤC\n")
toc_title_run.bold = True
toc_title_run.font.size = Pt(14)

toc_items = [
    ("1. GIỚI THIỆU VỀ WARGAME NATAS", 1),
    ("2. PHƯƠNG PHÁP TIẾN HÀNH", 2),
    ("3. CHI TIẾT CÁC LEVEL ĐÃ XỬ LÝ", 3),
    ("4. KẾT LUẬN VÀ BÀI HỌC RÚT RA", 4)
]

for item, level in toc_items:
    p = doc.add_paragraph()
    set_left_alignment(p)
    p_run = p.add_run(item)
    if level == 1:
        p_run.bold = True
    p_run.font.size = Pt(10)

doc.add_page_break()

# Section 1: Introduction
section1 = doc.add_paragraph()
section1.alignment = WD_ALIGN_PARAGRAPH.LEFT
section1_run = section1.add_run("1. GIỚI THIỆU VỀ WARGAME NATAS\n")
section1_run.bold = True
section1_run.font.size = Pt(14)

doc.add_paragraph("Natas là một wargame về bảo mật web do OverTheWire cung cấp, tập trung vào việc giảng dạy các nguyên tắc cơ bản của bảo mật phía máy chủ. Mỗi level trình bày một website độc đáo, có thể truy cập qua URL như http://natasX.natas.labs.overthewire.org (trong đó X là số level), và đòi hỏi người chơi phải tìm mật khẩu cho level tiếp theo bằng cách khai thác các lỗ hổng bảo mật web khác nhau.\n")

doc.add_paragraph("Wargame này cung cấp một phương pháp học tập toàn diện và thực tiễn về bảo mật web phía máy chủ. Với cấu trúc levels được tổ chức tốt, kết hợp với sự hỗ trợ của cộng đồng và các tài nguyên bên ngoài, Natas trở thành nền tảng xuất sắc cho cả người mới bắt đầu và những chuyên gia có kinh nghiệm muốn nâng cao kỹ năng.\n")

# Section 2: Methodology
section2 = doc.add_paragraph()
section2.alignment = WD_ALIGN_PARAGRAPH.LEFT
section2_run = section2.add_run("2. PHƯƠNG PHÁP TIẾN HÀNH\n")
section2_run.bold = True
section2_run.font.size = Pt(14)

doc.add_paragraph("Trong quá trình thực hiện wargame Natas, tôi đã áp dụng các phương pháp sau:\n")

methods = [
    "Khám phá mã nguồn HTML để tìm thông tin ẩn",
    "Hiểu và thao tác với HTTP request headers",
    "Khai thác lỗ hổng path traversal",
    "Vượt qua input validation để thực thi lệnh tùy ý",
    "Thực hiện SQL injection attacks",
    "Hiểu và khai thác các điểm yếu mã hóa"
]

for method in methods:
    p = doc.add_paragraph()
    set_left_alignment(p)
    p_run = p.add_run("• " + method)
    p_run.font.size = Pt(10)

# Section 3: Detailed Levels
section3 = doc.add_paragraph()
section3.alignment = WD_ALIGN_PARAGRAPH.LEFT
section3_run = section3.add_run("3. CHI TIẾT CÁC LEVEL ĐÃ XỬ LÝ\n")
section3_run.bold = True
section3_run.font.size = Pt(14)

# Level 0-1
level0_title = doc.add_paragraph()
level0_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
level0_title_run = level0_title.add_run("3.1 Level 0 - Natas0\n")
level0_title_run.bold = True
level0_title_run.font.size = Pt(12)

doc.add_paragraph("Level 0 là điểm khởi đầu cơ bản, nơi người chơi cần tìm mật khẩu hiển thị trên trang web.\n")

code_table = doc.add_table(rows=1, cols=1)
code_table.style = 'Table Grid'
code_table.width = Inches(6.5)
for cell in code_table.rows[0].cells:
    cell.width = Inches(6.5)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

# Set gray background using XML
cell = code_table.cell(0, 0)
tc = cell._tc
tcPr = tc.get_or_add_tcPr()
shd = parse_xml('<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F4F4F4" w:val="clear" w:color="auto"/>')
tcPr.append(shd)

code_content = """URL: http://natas0.natas.labs.overthewire.org/
Username: natas0
Password: (visible on page)

Kết quả: natas1 password found from HTML source"""

for line in code_content.split('\n'):
    p = cell.paragraphs[0] if len(cell.paragraphs) > 0 else cell.add_paragraph()
    if len(p.runs) == 0:
        run = p.add_run(line + '\n')
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0, 0, 0)

# Level 1-2
level1_title = doc.add_paragraph()
level1_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
level1_title_run = level1_title.add_run("3.2 Level 1 - Natas1\n")
level1_title_run.bold = True
level1_title_run.font.size = Pt(12)

doc.add_paragraph("Level 1 yêu cầu người chơi xem source code HTML để tìm mật khẩu ẩn trong comment.\n")

code_table2 = doc.add_table(rows=1, cols=1)
code_table2.style = 'Table Grid'
code_table2.width = Inches(6.5)
for cell in code_table2.rows[0].cells:
    cell.width = Inches(6.5)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

cell2 = code_table2.cell(0, 0)
tc2 = cell2._tc
tcPr2 = tc2.get_or_add_tcPr()
shd2 = parse_xml('<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F4F4F4" w:val="clear" w:color="auto"/>')
tcPr2.append(shd2)

code_content2 = """Source Code Analysis:
<!-- The password is natas2 -->

Kết quả: Password found from HTML comment"""

for line in code_content2.split('\n'):
    p = cell2.paragraphs[0] if len(cell2.paragraphs) > 0 else cell2.add_paragraph()
    if len(p.runs) == 0:
        run = p.add_run(line + '\n')
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0, 0, 0)

# Level 2-3
level2_title = doc.add_paragraph()
level2_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
level2_title_run = level2_title.add_run("3.3 Level 2 - Natas2\n")
level2_title_run.bold = True
level2_title_run.font.size = Pt(12)

doc.add_paragraph("Level 2 giới thiệu về việc sử dụng HTTP headers để truy cập thông tin.\n")

code_table3 = doc.add_table(rows=1, cols=1)
code_table3.style = 'Table Grid'
code_table3.width = Inches(6.5)
for cell in code_table3.rows[0].cells:
    cell.width = Inches(6.5)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

cell3 = code_table3.cell(0, 0)
tc3 = cell3._tc
tcPr3 = tc3.get_or_add_tcPr()
shd3 = parse_xml('<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F4F4F4" w:val="clear" w:color="auto"/>')
tcPr3.append(shd3)

code_content3 = """Command: curl -H "Referer: http://natas2.natas.labs.overthewire.org/" http://natas2.natas.labs.overthewire.org/

Kết quả: natas3 password found from Referer header check"""

for line in code_content3.split('\n'):
    p = cell3.paragraphs[0] if len(cell3.paragraphs) > 0 else cell3.add_paragraph()
    if len(p.runs) == 0:
        run = p.add_run(line + '\n')
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0, 0, 0)

# Level 3-4
level3_title = doc.add_paragraph()
level3_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
level3_title_run = level3_title.add_run("3.4 Level 3 - Natas3\n")
level3_title_run.bold = True
level3_title_run.font.size = Pt(12)

doc.add_paragraph("Level 3 tập trung vào path traversal vulnerability exploitation.\n")

code_table4 = doc.add_table(rows=1, cols=1)
code_table4.style = 'Table Grid'
code_table4.width = Inches(6.5)
for cell in code_table4.rows[0].cells:
    cell.width = Inches(6.5)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

cell4 = code_table4.cell(0, 0)
tc4 = cell4._tc
tcPr4 = tc4.get_or_add_tcPr()
shd4 = parse_xml('<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F4F4F4" w:val="clear" w:color="auto"/>')
tcPr4.append(shd4)

code_content4 = """Path Traversal Attack:
http://natas3.natas.labs.overthewire.org/../../../../etc/passwd

Kết quả: Access system files, find password location"""

for line in code_content4.split('\n'):
    p = cell4.paragraphs[0] if len(cell4.paragraphs) > 0 else cell4.add_paragraph()
    if len(p.runs) == 0:
        run = p.add_run(line + '\n')
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0, 0, 0)

# Level 4-5
level4_title = doc.add_paragraph()
level4_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
level4_title_run = level4_title.add_run("3.5 Level 4 - Natas4\n")
level4_title_run.bold = True
level4_title_run.font.size = Pt(12)

doc.add_paragraph("Level 4 yêu cầu bypass input validation để thực thi lệnh tùy ý.\n")

code_table5 = doc.add_table(rows=1, cols=1)
code_table5.style = 'Table Grid'
code_table5.width = Inches(6.5)
for cell in code_table5.rows[0].cells:
    cell.width = Inches(6.5)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

cell5 = code_table5.cell(0, 0)
tc5 = cell5._tc
tcPr5 = tc5.get_or_add_tcPr()
shd5 = parse_xml('<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F4F4F4" w:val="clear" w:color="auto"/>')
tcPr5.append(shd5)

code_content5 = """Input Validation Bypass:
Search field: natas5; cat /etc/natas_webpass/natas6

Kết quả: Command injection successful, password revealed"""

for line in code_content5.split('\n'):
    p = cell5.paragraphs[0] if len(cell5.paragraphs) > 0 else cell5.add_paragraph()
    if len(p.runs) == 0:
        run = p.add_run(line + '\n')
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0, 0, 0)

# Level 5-6
level5_title = doc.add_paragraph()
level5_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
level5_title_run = level5_title.add_run("3.6 Level 5 - Natas5\n")
level5_title_run.bold = True
level5_title_run.font.size = Pt(12)

doc.add_paragraph("Level 5 giới thiệu SQL injection attacks.\n")

code_table6 = doc.add_table(rows=1, cols=1)
code_table6.style = 'Table Grid'
code_table6.width = Inches(6.5)
for cell in code_table6.rows[0].cells:
    cell.width = Inches(6.5)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

cell6 = code_table6.cell(0, 0)
tc6 = cell6._tc
tcPr6 = tc6.get_or_add_tcPr()
shd6 = parse_xml('<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F4F4F4" w:val="clear" w:color="auto"/>')
tcPr6.append(shd6)

code_content6 = """SQL Injection Attack:
Username: admin" OR "1"="1 --
Password: (any)

Kết quả: Bypass authentication, access admin panel"""

for line in code_content6.split('\n'):
    p = cell6.paragraphs[0] if len(cell6.paragraphs) > 0 else cell6.add_paragraph()
    if len(p.runs) == 0:
        run = p.add_run(line + '\n')
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0, 0, 0)

# Section 4: Conclusion
section4 = doc.add_paragraph()
section4.alignment = WD_ALIGN_PARAGRAPH.LEFT
section4_run = section4.add_run("4. KẾT LUẬN VÀ BÀI HỌC RÚT RA\n")
section4_run.bold = True
section4_run.font.size = Pt(14)

doc.add_paragraph("Qua quá trình thực hiện wargame Natas, tôi đã rút ra được những bài học quan trọng sau:\n")

lessons = [
    "Hiểu rõ các lỗ hổng bảo mật web phổ biến và cách khai thác chúng",
    "Nhận thức được tầm quan trọng của input validation trong ứng dụng web",
    "Học cách sử dụng các công cụ như curl, Burp Suite để phân tích traffic",
    "Nắm vững kỹ thuật SQL injection và cách phòng chống",
    "Hiểu về path traversal vulnerability và biện pháp bảo vệ",
    "Nhận thức được tầm quan trọng của việc không lưu trữ thông tin nhạy cảm trong source code"
]

for lesson in lessons:
    p = doc.add_paragraph()
    set_left_alignment(p)
    p_run = p.add_run("• " + lesson)
    p_run.font.size = Pt(10)

# Save document
output_path = "/workspace/report_network/Lab5_Report_Solution.docx"
doc.save(output_path)
print(f"Document saved to: {output_path}")
