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

    # 1. School Header Layout
    if header_data.get('logo_file'):
        header_table = doc.add_table(rows=1, cols=2)
        remove_table_borders(header_table)
        header_table.rows[0].cells[0].width = Inches(1.2)
        header_table.rows[0].cells[1].width = Inches(5.8)
        
        logo_p = header_table.cell(0, 0).paragraphs[0]
        logo_run = logo_p.add_run()
        logo_run.add_picture(header_data['logo_file'], width=Inches(1.0))
        
        text_cell = header_table.cell(0, 1)
        p = text_cell.paragraphs[0]
    else:
        p = doc.add_paragraph()
        
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(header_data['school_name']).bold = True
    p.runs[0].font.size = Pt(15)

    if header_data.get('logo_file'):
        p_sub = text_cell.add_paragraph()
    else:
        p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.add_run(f"{header_data['address']}\nPh: {header_data['phone']} | Email: {header_data['email']}").font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 2. Metadata Grid (Time, Marks, Name, Roll)[span_0](start_span)[span_0](end_span)
    meta_table = doc.add_table(rows=2, cols=2)
    remove_table_borders(meta_table)
    for row in meta_table.rows:
        row.cells[0].width = Inches(3.5)
        row.cells[1].width = Inches(3.5)
    
    meta_table.cell(0, 0).paragraphs[0].text = f"Time: {header_data['time']}[span_1](start_span)"[span_1](end_span)
    m_cell = meta_table.cell(0, 1).paragraphs[0]
    m_cell.text = f"M.M: {header_data['max_marks']}[span_2](start_span)"[span_2](end_span)
    m_cell.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    meta_table.cell(1, 0).paragraphs[0].text = "Name: ______________________[span_3](start_span)"[span_3](end_span)
    r_cell = meta_table.cell(1, 1).paragraphs[0]
    r_cell.text = "Roll No: ____________[span_4](start_span)"[span_4](end_span)
    r_cell.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Assessment Title[span_5](start_span)[span_5](end_span)
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(12)
    r_title = p_title.add_run(f"{header_data['assessment_name']}\nSUBJECT - {header_data['subject']} | CLASS - {header_data['class_name']}")[span_6](start_span)[span_6](end_span)
    r_title.bold = True

    # 3. Dynamic Question Generation Loop[span_7](start_span)[span_7](end_span)
    for idx, q in enumerate(questions_list, 1):
        q_type = q['type']
        
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
            
            # Lock sizes for columns to balance layout evenly
            for row in match_table.rows:
                row.cells[0].width = Inches(3.5)
                row.cells[1].width = Inches(3.5)
                
            for r_idx, pair in enumerate(q['pairs']):
                c1 = match_table.cell(r_idx, 0).paragraphs[0]
                c2 = match_table.cell(r_idx, 1).paragraphs[0]
                c1.paragraph_format.left_indent = Inches(0.4)
                c2.paragraph_format.left_indent = Inches(0.4)
                
                # Render Left Side (Text or Image)
                if pair.get('left_type') == "Image" and pair.get('left_img'):
                    c1.text = pair.get('left_prefix', '') + " "
                    run_img1 = c1.add_run()
                    run_img1.add_picture(pair['left_img'], width=Inches(1.0))
                else:
                    c1.text = pair.get('left_text', '')
                    
                # Render Right Side (Text or Image)
                if pair.get('right_type') == "Image" and pair.get('right_img'):
                    c2.text = pair.get('right_prefix', '') + " "
                    run_img2 = c2.add_run()
                    run_img2.add_picture(pair['right_img'], width=Inches(1.0))
                else:
                    c2.text = pair.get('right_text', '')
                    
            match_table.rows[-1].cells[0].paragraphs[0].paragraph_format.space_after = Pt(8)

        elif q_type == "Image/Source Based":
            if q.get('image_file'):
                img_p = doc.add_paragraph()
                img_p.paragraph_format.left_indent = Inches(0.4)
                img_run = img_p.add_run()
                img_run.add_picture(q['image_file'], width=Inches(3.0))
            
            for sub_idx, sub_q in enumerate(q['sub_questions'], 1):
                sub_table = doc.add_table(rows=1, cols=2)
                remove_table_borders(sub_table)
                sub_table.rows[0].cells[0].width = Inches(0.7)
                sub_table.rows[0].cells[1].width = Inches(6.3)
                sub_table.cell(0, 0).paragraphs[0].text = f"  ({sub_idx})"
                sub_table.cell(0, 1).paragraphs[0].text = sub_q
            doc.add_paragraph().paragraph_format.space_after = Pt(4)
            
        else:
            doc.add_paragraph().paragraph_format.space_after = Pt(4)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# --- streamlit frontend interface ---
st.set_page_config(page_title="Exam Creator Pro", layout="wide")
st.title("📝 Automatic Exam Sheet Template Generator")

with st.sidebar:
    st.header("🏫 School & Header Info")
    uploaded_logo = st.file_uploader("Upload School Logo", type=["png", "jpg", "jpeg"])
    h_data = {
        "logo_file": uploaded_logo,
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
            new_q["pairs"] = [{
                "left_type": "Text", "left_text": "i) Item", "left_prefix": "i)", "left_img": None,
                "right_type": "Text", "right_text": "a. Target", "right_prefix": "a.", "right_img": None
            }]
        elif q_type_sel == "Image/Source Based":
            new_q["image_file"] = None
            new_q["sub_questions"] = [""]
        st.session_state.questions.append(new_q)

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
                p_count = len(question['pairs'])
                question['pairs'].append({
                    "left_type": "Text", "left_text": f"i) Item", "left_prefix": "i)", "left_img": None,
                    "right_type": "Text", "right_text": "a. Target", "right_prefix": "a.", "right_img": None
                })
            
            for p_idx, pair in enumerate(question['pairs']):
                st.markdown(f"--- **Pair Row {p_idx+1}** ---")
                l_col, r_col = st.columns(2)
                
                # Left Item Configuration
                with l_col:
                    pair['left_type'] = st.radio(f"Left Type ({p_idx+1})", ["Text", "Image"], index=0 if pair['left_type']=="Text" else 1, key=f"ltype_{idx}_{p_idx}")
                    pair['left_prefix'] = st.text_input(f"Left Index (e.g. i)", value=pair.get('left_prefix', ''), key=f"lpref_{idx}_{p_idx}")
                    if pair['left_type'] == "Text":
                        pair['left_text'] = st.text_input(f"Left Text String", value=pair.get('left_text', ''), key=f"ltxt_{idx}_{p_idx}")
                    else:
                        pair['left_img'] = st.file_uploader(f"Upload Left Image", type=["png","jpg","jpeg"], key=f"limg_{idx}_{p_idx}")
                
                # Right Item Configuration
                with r_col:
                    pair['right_type'] = st.radio(f"Right Type ({p_idx+1})", ["Text", "Image"], index=0 if pair['right_type']=="Text" else 1, key=f"rtype_{idx}_{p_idx}")
                    pair['right_prefix'] = st.text_input(f"Right Index (e.g. a.)", value=pair.get('right_prefix', ''), key=f"rpref_{idx}_{p_idx}")
                    if pair['right_type'] == "Text":
                        pair['right_text'] = st.text_input(f"Right Text String", value=pair.get('right_text', ''), key=f"rtxt_{idx}_{p_idx}")
                    else:
                        pair['right_img'] = st.file_uploader(f"Upload Right Image", type=["png","jpg","jpeg"], key=f"rimg_{idx}_{p_idx}")
                
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
