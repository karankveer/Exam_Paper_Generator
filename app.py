import streamlit as st
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

# --- helper function for borderless tables ---
def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = tblPr.first_child_found_in("w:tblBorders")
    if tblBorders is not None:
        tblPr.remove(tblBorders)
    new_borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="none"/>\n'
        f'  <w:left w:val="none"/>\n'
        f'  <w:bottom w:val="none"/>\n'
        f'  <w:right w:val="none"/>\n'
        f'  <w:insideH w:val="none"/>\n'
        f'  <w:insideV w:val="none"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(new_borders)

# --- docx engine logic ---
def build_docx(header_data, questions_list):
    doc = docx.Document()
    
    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # Global Font Setup
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(10.5)

    # 1. School Header
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(header_data['school_name']).bold = True
    p.runs[0].font.size = Pt(15)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.add_run(f"{header_data['address']}\nPh: {header_data['phone']} | Email: {header_data['email']}").font.size = Pt(9)

    # 2. Metadata Grid (Time, Marks, Name, Roll)
    meta_table = doc.add_table(rows=2, cols=2)
    remove_table_borders(meta_table)
    for row in meta_table.rows:
        row.cells[0].width = Inches(3.5)
        row.cells[1].width = Inches(3.5)
    
    meta_table.cell(0, 0).paragraphs[0].text = f"Time: {header_data['time']}"
    m_cell = meta_table.cell(0, 1).paragraphs[0]
    m_cell.text = f"M.M: {header_data['max_marks']}"
    m_cell.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    meta_table.cell(1, 0).paragraphs[0].text = "Name: ______________________"
    r_cell = meta_table.cell(1, 1).paragraphs[0]
    r_cell.text = "Roll No: ____________"
    r_cell.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Assessment Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(12)
    r_title = p_title.add_run(f"{header_data['assessment_name']}\nSUBJECT - {header_data['subject']} | CLASS - {header_data['class_name']}")
    r_title.bold = True

    # 3. Dynamic Question Generation Loop
    for idx, q in enumerate(questions_list, 1):
        q_type = q['type']
        
        # Base Question layout table (Numbering protection)
        q_table = doc.add_table(rows=1, cols=2)
        remove_table_borders(q_table)
        q_table.rows[0].cells[0].width = Inches(0.4)
        q_table.rows[0].cells[1].width = Inches(6.6)
        
        q_table.cell(0, 0).paragraphs[0].text = f"Q.{idx}"
        q_text_p = q_table.cell(0, 1).paragraphs[0]
        q_text_p.text = q['text']
        
        if q.get('marks'):
            m_run = q_text_p.add_run(f"   ({q['marks']})")
            m_run.bold = True

        # Render Specific Layout Structures safely
        if q_type == "MCQ":
            opt_table = doc.add_table(rows=1, cols=4)
            remove_table_borders(opt_table)
            for i, cell in enumerate(opt_table.rows[0].cells):
                cell.width = Inches(1.75)
                p_opt = cell.paragraphs[0]
                p_opt.paragraph_format.left_indent = Inches(0.4)
                p_opt.text = q['options'][i]
            opt_table.rows[0].cells[0].paragraphs[0].paragraph_format.space_after = Pt(8)

        elif q_type == "Match the Following":
            match_table = doc.add_table(rows=len(q['pairs']), cols=2)
            remove_table_borders(match_table)
            for r_idx, pair in enumerate(q['pairs']):
                # Sub-indent column positioning securely
                c1 = match_table.cell(r_idx, 0).paragraphs[0]
                c2 = match_table.cell(r_idx, 1).paragraphs[0]
                c1.paragraph_format.left_indent = Inches(0.4)
                c2.paragraph_format.left_indent = Inches(0.4)
                c1.text = pair[0]
                c2.text = pair[1]
            match_table.rows[-1].cells[0].paragraphs[0].paragraph_format.space_after = Pt(8)

        elif q_type == "Image/Source Based":
            if q.get('image_file'):
                img_p = doc.add_paragraph()
                img_p.paragraph_format.left_indent = Inches(0.4)
                img_run = img_p.add_run()
                img_run.add_picture(q['image_file'], width=Inches(3.0)) # Scaled correctly
            
            # Sub-questions nested under image context
            for sub_idx, sub_q in enumerate(q['sub_questions'], 1):
                sub_table = doc.add_table(rows=1, cols=2)
                remove_table_borders(sub_table)
                sub_table.rows[0].cells[0].width = Inches(0.7)
                sub_table.rows[0].cells[1].width = Inches(6.3)
                sub_table.cell(0, 0).paragraphs[0].text = f"  ({sub_idx})"
                sub_table.cell(0, 1).paragraphs[0].text = sub_q
            doc.add_paragraph().paragraph_format.space_after = Pt(4)
            
        else: # Fill in blanks or short questions
            doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Save out to buffer stream
    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# --- streamlit frontend interface ---
st.set_page_config(page_title="Exam Creator Pro", layout="wide")
st.title("📝 Automatic Exam Sheet Template Generator")
st.caption("Input your contents safely. Formatting, indentation alignments, and grid structural limits are forced automatically.")

# Sidebar Configuration for Header Details
with st.sidebar:
    st.header("🏫 School & Header Info")
    h_data = {
        "school_name": st.text_input("School/Institution Name", "VSI GLOBAL SR. SEC. SCHOOL"),
        "address": st.text_area("Address Line", "Sec. 5, Pratap Nagar, Tonk Road, Jaipur"),
        "phone": st.text_input("Phone Number", "9309305656"),
        "email": st.text_input("Email ID", "vsiglobalschool@gmail.com"),
        "assessment_name": st.text_input("Assessment Title", "ASSESSMENT SHEET - 2026-27"),
        "subject": st.text_input("Subject", "EVS"),
        "class_name": st.text_input("Class Level", "III"),
        "time": st.text_input("Time Duration Limit", "1 HOUR"),
        "max_marks": st.text_input("Maximum Marks (M.M)", "20")
    }

# Main Application Dynamic Input Area
if 'questions' not in st.session_state:
    st.session_state.questions = []

st.subheader("🛠️ Step 2: Build Paper Structure")
col_num, col_type, col_add = st.columns([2, 4, 2])

with col_type:
    q_type_sel = st.selectbox("Choose Question Type to Add", ["MCQ", "Fill in Blanks / Short Ans", "Match the Following", "Image/Source Based"])

with col_add:
    st.write("##")
    if st.button("➕ Add This Question Element"):
        new_q = {"type": q_type_sel, "text": "", "marks": ""}
        if q_type_sel == "MCQ":
            new_q["options"] = ["(a) ", "(b) ", "(c) ", "(d) "]
        elif q_type_sel == "Match the Following":
            new_q["pairs"] = [["i) Left Item A", "a. Right Item X"], ["ii) Left Item B", "b. Right Item Y"]]
        elif q_type_sel == "Image/Source Based":
            new_q["image_file"] = None
            new_q["sub_questions"] = [""]
        st.session_state.questions.append(new_q)

# Loop to draw forms sequentially on screen based on what she added
for idx, question in enumerate(st.session_state.questions):
    with st.expander(f"Question N°{idx+1} — Form Category Layout: **{question['type']}**", expanded=True):
        c_q, c_m = st.columns([6, 2])
        question['text'] = c_q.text_area(f"Question/Instruction Text {idx+1}", value=question['text'], key=f"txt_{idx}")
        question['marks'] = c_m.text_input(f"Marks allocation string", value=question['marks'], key=f"mrk_{idx}", placeholder="e.g. 0.5 × 5 = 2.5")
        
        if question['type'] == "MCQ":
            st.markdown("**Enter Option Values below:**")
            opt_cols = st.columns(4)
            for o_idx in range(4):
                question['options'][o_idx] = opt_cols[o_idx].text_input(f"Option {chr(97+o_idx)}", value=question['options'][o_idx], key=f"opt_{idx}_{o_idx}")
                
        elif question['type'] == "Match the Following":
            st.markdown("**Define Match Pairs:**")
            if st.button("➕ Append row pair line", key=f"add_pair_{idx}"):
                question['pairs'].append(["", ""])
            for p_idx, pair in enumerate(question['pairs']):
                cp1, cp2 = st.columns(2)
                question['pairs'][p_idx][0] = cp1.text_input(f"Left Line {p_idx+1}", value=pair[0], key=f"p1_{idx}_{p_idx}")
                question['pairs'][p_idx][1] = cp2.text_input(f"Right Line {p_idx+1}", value=pair[1], key=f"p2_{idx}_{p_idx}")
                
        elif question['type'] == "Image/Source Based":
            question['image_file'] = st.file_uploader("Upload reference diagram source asset", type=["png", "jpg", "jpeg"], key=f"img_{idx}")
            st.markdown("**Sub-questions attached below context:**")
            if st.button("➕ Append sub-question step line", key=f"add_sub_{idx}"):
                question['sub_questions'].append("")
            for s_idx, sub_q in enumerate(question['sub_questions']):
                question['sub_questions'][s_idx] = st.text_input(f"Sub-question ({s_idx+1}) Text", value=sub_q, key=f"sub_{idx}_{s_idx}")

        if st.button("🗑️ Delete Question Element", key=f"del_{idx}"):
            st.session_state.questions.pop(idx)
            st.rerun()

# --- Compile File Triggers ---
st.write("---")
if st.session_state.questions:
    if st.button("🚀 Render and Compile Document Grid Template"):
        with st.spinner("Assembling layout structure properties safely..."):
            docx_buffer = build_docx(h_data, st.session_state.questions)
            st.success("Compilation Success! Your flawlessly aligned document is ready.")
            st.download_button(
                label="📥 Download Microsoft Word (.docx) File",
                data=docx_buffer,
                file_name=f"{h_data['subject']}_Exam_Template.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
else:
    st.info("Add some question structural components above to begin generating the interactive file document template.")
