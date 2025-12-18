import tkinter as tk

from tkinter import filedialog, messagebox, ttk

import re

import random

import zipfile

import io

import os

import csv

from xml.dom import minidom



# ==================== PHẦN 1: LOGIC XỬ LÝ (CORE) ====================



W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"



def parse_range_string(s):

    """Chuyển chuỗi '1, 2, 5-8' thành set {1, 2, 5, 6, 7, 8}"""

    res = set()

    if not s: return res

    parts = s.split(',')

    for part in parts:

        part = part.strip()

        if not part: continue

        if '-' in part:

            try:

                start, end = map(int, part.split('-'))

                res.update(range(start, end + 1))

            except: pass

        else:

            try:

                res.add(int(part))

            except: pass

    return res



def escape_xml(text):

    if not text: return ""

    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")



def is_correct_option(block):

    r_nodes = block.getElementsByTagNameNS(W_NS, "r")

    for r in r_nodes:

        rPr_list = r.getElementsByTagNameNS(W_NS, "rPr")

        for rPr in rPr_list:

            u_list = rPr.getElementsByTagNameNS(W_NS, "u")

            if u_list:

                val = u_list[0].getAttributeNS(W_NS, "val")

                if val and val != "none": return True

            color_list = rPr.getElementsByTagNameNS(W_NS, "color")

            if color_list:

                val = color_list[0].getAttributeNS(W_NS, "val")

                if val and val.upper() in ["FF0000", "RED"]: return True

    return False



def extract_short_answer_key(question_blocks):

    key = ""

    clean_blocks = []

    for block in question_blocks:

        txt = get_text(block)

        m = re.match(r'^\s*(?:Đáp án|DA|Lời giải|HD|Hướng dẫn)\s*[:\.]?\s*(.*)', txt, re.IGNORECASE)

        if m:

            key = m.group(1).strip()

            continue

        clean_blocks.append(block)

    return clean_blocks, key



def get_text(block):

    texts = []

    t_nodes = block.getElementsByTagNameNS(W_NS, "t")

    for t in t_nodes:

        if t.firstChild and t.firstChild.nodeValue:

            texts.append(t.firstChild.nodeValue)

    return "".join(texts).strip()



# --- TAB LAYOUT HELPERS ---



def set_paragraph_tabs(paragraph, tab_positions):

    doc = paragraph.ownerDocument

    pPr_list = paragraph.getElementsByTagNameNS(W_NS, "pPr")

    if not pPr_list:

        pPr = doc.createElementNS(W_NS, "w:pPr")

        paragraph.insertBefore(pPr, paragraph.firstChild)

    else: pPr = pPr_list[0]

    tabs_list = pPr.getElementsByTagNameNS(W_NS, "tabs")

    for tabs in tabs_list: pPr.removeChild(tabs)

    w_tabs = doc.createElementNS(W_NS, "w:tabs")

    for pos in tab_positions:

        w_tab = doc.createElementNS(W_NS, "w:tab")

        w_tab.setAttributeNS(W_NS, "w:val", "left")

        w_tab.setAttributeNS(W_NS, "w:pos", str(pos))

        w_tabs.appendChild(w_tab)

    pPr.appendChild(w_tabs)



def merge_paragraphs(p_dest, p_src):

    doc = p_dest.ownerDocument

    r_tab = doc.createElementNS(W_NS, "w:r")

    tab = doc.createElementNS(W_NS, "w:tab")

    r_tab.appendChild(tab)

    p_dest.appendChild(r_tab)

    children = []

    for child in p_src.childNodes:

        if child.localName not in ["pPr", "proofErr", "bookmarkStart", "bookmarkEnd"]:

            children.append(child)

    for child in children: p_dest.appendChild(child)

    return p_dest



def format_mcq_layout(question_blocks):

    option_indices = []

    for i, block in enumerate(question_blocks):

        if re.match(r'^\s*[A-D][\.\)]', get_text(block), re.IGNORECASE):

            option_indices.append(i)

    if len(option_indices) != 4: return question_blocks

    opt_blocks = [question_blocks[i] for i in option_indices]

    lengths = [len(get_text(b)) for b in opt_blocks]

    max_len = max(lengths)

    layout_mode = 1

    if max_len < 20: layout_mode = 4

    elif max_len < 45: layout_mode = 2

    else: layout_mode = 1

    if layout_mode == 1: return question_blocks

    new_question_blocks = []

    for i in range(option_indices[0]): new_question_blocks.append(question_blocks[i])

    if layout_mode == 4:

        p_root = opt_blocks[0]

        merge_paragraphs(p_root, opt_blocks[1])

        merge_paragraphs(p_root, opt_blocks[2])

        merge_paragraphs(p_root, opt_blocks[3])

        set_paragraph_tabs(p_root, [3000, 6000, 9000])

        new_question_blocks.append(p_root)

    elif layout_mode == 2:

        row1 = opt_blocks[0]

        merge_paragraphs(row1, opt_blocks[1])

        set_paragraph_tabs(row1, [6000])

        new_question_blocks.append(row1)

        row2 = opt_blocks[2]

        merge_paragraphs(row2, opt_blocks[3])

        set_paragraph_tabs(row2, [6000])

        new_question_blocks.append(row2)

    last_opt_idx = option_indices[-1]

    for i in range(last_opt_idx + 1, len(question_blocks)): new_question_blocks.append(question_blocks[i])

    return new_question_blocks



# --- EXISTING HELPERS ---



def style_run_blue_bold(run):

    doc = run.ownerDocument

    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")

    if rPr_list: rPr = rPr_list[0]

    else:

        rPr = doc.createElementNS(W_NS, "w:rPr")

        run.insertBefore(rPr, run.firstChild)

    color_list = rPr.getElementsByTagNameNS(W_NS, "color")

    if color_list: color_el = color_list[0]

    else:

        color_el = doc.createElementNS(W_NS, "w:color")

        rPr.appendChild(color_el)

    color_el.setAttributeNS(W_NS, "w:val", "0000FF")

    b_list = rPr.getElementsByTagNameNS(W_NS, "b")

    if not b_list:

        b_el = doc.createElementNS(W_NS, "w:b")

        rPr.appendChild(b_el)



def update_mcq_label(paragraph, new_label):

    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")

    if not t_nodes: return

    new_letter = new_label[0].upper()

    for i, t in enumerate(t_nodes):

        if not t.firstChild: continue

        txt = t.firstChild.nodeValue

        m = re.match(r'^(\s*)([A-D])([\.\)])?', txt, re.IGNORECASE)

        if not m: continue

        leading_space = m.group(1) or ""

        old_punct = m.group(3) or ""

        after_match = txt[m.end():]

        t.firstChild.nodeValue = leading_space + new_letter + ("." if not old_punct else old_punct) + " " + after_match.strip()

        run = t.parentNode

        if run and run.localName == "r": style_run_blue_bold(run)

        for j in range(i + 1, len(t_nodes)):

            t2 = t_nodes[j]

            if not t2.firstChild: continue

            val2 = t2.firstChild.nodeValue

            if re.match(r'^[\s\.]+$', val2): t2.firstChild.nodeValue = ""

            elif re.match(r'^\.', val2): 

                t2.firstChild.nodeValue = val2[1:]

                break

            else: break

        break



def update_tf_label(paragraph, new_label):

    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")

    if not t_nodes: return

    new_letter = new_label[0].lower()

    for i, t in enumerate(t_nodes):

        if not t.firstChild: continue

        txt = t.firstChild.nodeValue

        m = re.match(r'^(\s*)([a-d])(\))?', txt, re.IGNORECASE)

        if not m: continue

        leading_space = m.group(1) or ""

        after_match = txt[m.end():]

        t.firstChild.nodeValue = leading_space + new_letter + ")" + after_match

        run = t.parentNode

        if run and run.localName == "r": style_run_blue_bold(run)

        for j in range(i + 1, len(t_nodes)):

            t2 = t_nodes[j]

            if not t2.firstChild: continue

            val2 = t2.firstChild.nodeValue

            if re.match(r'^[\s\)]+$', val2): t2.firstChild.nodeValue = ""

            elif re.match(r'^\s*\)', val2):

                t2.firstChild.nodeValue = re.sub(r'^\s*\)', '', val2, count=1)

                break

            else: break

        break



def update_question_label(paragraph, new_label):

    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")

    if not t_nodes: return

    for i, t in enumerate(t_nodes):

        if not t.firstChild: continue

        txt = t.firstChild.nodeValue

        m = re.match(r'^(\s*)(Câu\s*)(\d+)(\.)?', txt, re.IGNORECASE)

        if not m: continue

        leading_space = m.group(1) or ""

        after_match = txt[m.end():]

        t.firstChild.nodeValue = leading_space + new_label + after_match

        run = t.parentNode

        if run and run.localName == "r": style_run_blue_bold(run)

        for j in range(i + 1, len(t_nodes)):

            t2 = t_nodes[j]

            if not t2.firstChild: continue

            if re.match(r'^[\s0-9\.]*$', t2.firstChild.nodeValue): t2.firstChild.nodeValue = ""

            else: break

        break



def find_part_index(blocks, part_number):

    pattern = re.compile(rf'PHẦN\s*{part_number}\b', re.IGNORECASE)

    for i, block in enumerate(blocks):

        if pattern.search(get_text(block)): return i

    return -1



# --- UPDATED PARSER WITH MARKER-BASED CLUSTER DETECTION ---



def parse_questions_in_range(blocks, start, end):

    part_blocks = blocks[start:end]

    items = [] 

    intro = []

    

    i = 0

    # 1. Lấy Intro đầu phần

    while i < len(part_blocks):

        text = get_text(part_blocks[i])

        if re.match(r'^Câu\s*\d+\b', text, re.IGNORECASE): break

        if "@BẮT ĐẦU DÙNG CHUNG@" in text.upper(): break

        intro.append(part_blocks[i])

        i += 1

        

    # 2. Quét chính

    while i < len(part_blocks):

        block = part_blocks[i]

        text = get_text(block)

        

        # A. Kiểm tra Marker BẮT ĐẦU

        if "@BẮT ĐẦU DÙNG CHUNG@" in text.upper():

            cluster_header = []

            cluster_questions = []

            i += 1 # Bỏ qua dòng chứa marker

            

            while i < len(part_blocks):

                b_curr = part_blocks[i]

                t_curr = get_text(b_curr)

                

                if "@KẾT THÚC DÙNG CHUNG@" in t_curr.upper():

                    i += 1 # Bỏ qua dòng chứa marker

                    break

                

                if re.match(r'^Câu\s*\d+\b', t_curr, re.IGNORECASE):

                    one_q = [b_curr]

                    i += 1

                    while i < len(part_blocks):

                        b_next = part_blocks[i]

                        t_next = get_text(b_next)

                        if "@KẾT THÚC DÙNG CHUNG@" in t_next.upper(): break

                        if re.match(r'^Câu\s*\d+\b', t_next, re.IGNORECASE): break

                        one_q.append(b_next)

                        i += 1

                    cluster_questions.append(one_q)

                else:

                    if cluster_questions:

                        cluster_questions[-1].append(b_curr)

                    else:

                        cluster_header.append(b_curr)

                    i += 1

            

            items.append({

                "type": "cluster",

                "header": cluster_header,

                "questions": cluster_questions

            })

            continue



        # B. Nếu là câu hỏi thường (Lẻ)

        if re.match(r'^Câu\s*\d+\b', text, re.IGNORECASE):

            group = [block]

            i += 1

            while i < len(part_blocks):

                t2 = get_text(part_blocks[i])

                if re.match(r'^Câu\s*\d+\b', t2, re.IGNORECASE): break

                if "@BẮT ĐẦU DÙNG CHUNG@" in t2.upper(): break

                if re.match(r'^PHẦN\s*\d\b', t2, re.IGNORECASE): break

                

                group.append(part_blocks[i])

                i += 1

            items.append({

                "type": "question",

                "blocks": group

            })

        else:

            if items and items[-1]["type"] == "question":

                 items[-1]["blocks"].append(block)

            elif not items:

                intro.append(block)

            i += 1

            

    return intro, items



def shuffle_array(arr):

    out = arr.copy()

    for i in range(len(out) - 1, 0, -1):

        j = random.randint(0, i)

        out[i], out[j] = out[j], out[i]

    return out



# --- HEADER AND FOOTER FUNCTIONS ---



def create_header_xml(doc, info):

    so_gd = escape_xml(info.get("so_gd", "").upper())

    truong = escape_xml(info.get("truong", ""))

    ky_thi = escape_xml(info.get("ky_thi", "").upper())

    mon_thi = escape_xml(info.get("mon_thi", "").upper())

    thoi_gian = escape_xml(info.get("thoi_gian", ""))

    nam_hoc = escape_xml(info.get("nam_hoc", ""))



    xml_str = f"""

    <w:tbl xmlns:w="{W_NS}">

        <w:tblPr>

            <w:tblW w:w="0" w:type="auto"/>

            <w:jc w:val="center"/>

            <w:tblBorders>

                <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>

                <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>

                <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>

                <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>

                <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>

                <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>

            </w:tblBorders>

        </w:tblPr>

        <w:tr>

            <w:tc>

                <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>

                <w:p>

                    <w:pPr><w:jc w:val="center"/></w:pPr>

                    <w:r><w:rPr><w:b/></w:rPr><w:t>{so_gd}</w:t></w:r>

                </w:p>

                <w:p>

                    <w:pPr><w:jc w:val="center"/></w:pPr>

                    <w:r><w:rPr><w:b/></w:rPr><w:t>{truong}</w:t></w:r>

                </w:p>

                <w:p>

                    <w:pPr><w:jc w:val="center"/></w:pPr>

                    <w:r><w:t>------------------</w:t></w:r>

                </w:p>

            </w:tc>

            <w:tc>

                <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>

                <w:p>

                    <w:pPr><w:jc w:val="center"/></w:pPr>

                    <w:r><w:rPr><w:b/></w:rPr><w:t>{ky_thi}</w:t></w:r>

                </w:p>

                <w:p>

                    <w:pPr><w:jc w:val="center"/></w:pPr>

                    <w:r><w:rPr><w:b/></w:rPr><w:t>MÔN: {mon_thi}</w:t></w:r>

                </w:p>

                <w:p>

                    <w:pPr><w:jc w:val="center"/></w:pPr>

                    <w:r><w:t>Thời gian làm bài: {thoi_gian}</w:t></w:r>

                </w:p>

                 <w:p>

                    <w:pPr><w:jc w:val="center"/></w:pPr>

                    <w:r><w:t>(Năm học: {nam_hoc})</w:t></w:r>

                </w:p>

            </w:tc>

        </w:tr>

    </w:tbl>

    """

    return minidom.parseString(xml_str).documentElement



def create_footer_xml_content(ma_de):

    xml_str = f"""

    <w:ftr xmlns:w="{W_NS}">

        <w:p>

            <w:pPr>

                <w:pStyle w:val="Footer"/>

                <w:jc w:val="right"/>

                <w:pBdr>

                    <w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/>

                </w:pBdr>

            </w:pPr>

            <w:r>

                <w:t xml:space="preserve">Mã đề {ma_de} - Trang </w:t>

            </w:r>

            <w:fldSimple w:instr="PAGE"/>

        </w:p>

    </w:ftr>

    """

    return xml_str.strip()



def add_header_to_body(dom, body, header_info):

    if not header_info.get("enable", False): return

    try:

        tbl_node = create_header_xml(dom, header_info)

        if body.firstChild: body.insertBefore(tbl_node, body.firstChild)

        else: body.appendChild(tbl_node)

        p_empty = dom.createElementNS(W_NS, "w:p")

        if body.childNodes.length > 1: body.insertBefore(p_empty, body.childNodes[1])

    except: pass



# --- UPDATED SHUFFLE FUNCTIONS ---



def relabel_mcq_options(question_blocks):

    letters = ["A", "B", "C", "D"]

    count = 0

    for block in question_blocks:

        if re.match(r'^\s*[A-D][\.\)]', get_text(block), re.IGNORECASE):

            l = letters[count] if count < 4 else "D"

            update_mcq_label(block, f"{l}.")

            count += 1



def relabel_tf_options(question_blocks):

    letters = ["a", "b", "c", "d"]

    count = 0

    for block in question_blocks:

        if re.match(r'^\s*[a-d]\)', get_text(block), re.IGNORECASE):

            l = letters[count] if count < 4 else "d"

            update_tf_label(block, f"{l})")

            count += 1



def shuffle_mcq_options(question_blocks, allow_shuffle=True):

    indices = []

    correct_indices_before = []

    for i, block in enumerate(question_blocks):

        if re.match(r'^\s*[A-D][\.\)]', get_text(block), re.IGNORECASE):

            indices.append(i)

            if is_correct_option(block): correct_indices_before.append(i)

    if len(indices) < 2: return question_blocks, ""

    options = [question_blocks[idx] for idx in indices]

    perm = list(range(len(options)))

    if allow_shuffle: random.shuffle(perm)

    shuffled_options = [options[p] for p in perm]

    new_correct_char = ""

    if correct_indices_before:

        orig_correct_idx_in_options = -1

        for k, val in enumerate(indices):

            if val == correct_indices_before[0]:

                orig_correct_idx_in_options = k

                break

        if orig_correct_idx_in_options != -1:

            for new_pos, old_pos in enumerate(perm):

                if old_pos == orig_correct_idx_in_options:

                    letters = ["A", "B", "C", "D", "E", "F"]

                    if new_pos < len(letters): new_correct_char = letters[new_pos]

                    break

    min_idx, max_idx = min(indices), max(indices)

    before = question_blocks[:min_idx]

    after = question_blocks[max_idx + 1:]

    return before + shuffled_options + after, new_correct_char



def shuffle_tf_options(question_blocks, allow_shuffle=True):

    option_indices = {}

    for i, block in enumerate(question_blocks):

        m = re.match(r'^\s*([a-d])\)', get_text(block), re.IGNORECASE)

        if m: option_indices[m.group(1).lower()] = i

    abc_idx = [option_indices.get(k) for k in ["a", "b", "c"] if option_indices.get(k) is not None]

    if len(abc_idx) < 2: return question_blocks, ["", "", "", ""]

    abc_nodes = [question_blocks[idx] for idx in abc_idx]

    if allow_shuffle: shuffled_abc = shuffle_array(abc_nodes)

    else: shuffled_abc = abc_nodes.copy()

    all_vals = [v for v in option_indices.values() if v is not None]

    min_idx, max_idx = min(all_vals), max(all_vals)

    before = question_blocks[:min_idx]

    after = question_blocks[max_idx + 1:]

    d_node = question_blocks[option_indices["d"]] if "d" in option_indices else None

    middle = shuffled_abc.copy()

    if d_node: middle.append(d_node)

    current_key_status = []

    for block in middle:

        status = "D" if is_correct_option(block) else "S"

        current_key_status.append(status)

    return before + middle + after, current_key_status



def process_single_question_logic(q, part_type, allow_shuffle_opt):

    """Xử lý 1 câu hỏi (trộn opt, lấy key, format layout)"""

    new_block = []

    key = ""

    if part_type == "PHAN1":

        new_block, key = shuffle_mcq_options(q, allow_shuffle_opt)

    elif part_type == "PHAN2":

        new_block, key = shuffle_tf_options(q, allow_shuffle_opt)

    elif part_type == "PHAN3":

        new_block, key = extract_short_answer_key(q)

    else:

        new_block = q.copy()

    return new_block, key



def process_part(blocks, start, end, part_type, global_q_idx_start, config):

    intro, items = parse_questions_in_range(blocks, start, end)

    

    # 1. Expand items và xử lý nội dung bên trong

    processed_items = []

    current_q_counter = global_q_idx_start

    

    fixed_pos_set = config.get("fixed_pos_set", set())

    fixed_opt_set = config.get("fixed_opt_set", set())

    fix_group_pos = config.get("fix_group_pos", False) # Cờ mới

    

    for item in items:

        if item["type"] == "question":

            q_idx = current_q_counter + 1

            allow_opt = config.get("shuffle_opt_global", True)

            if q_idx in fixed_opt_set: allow_opt = False

            

            new_q, key = process_single_question_logic(item["blocks"], part_type, allow_opt)

            processed_items.append({

                "type": "question",

                "blocks": new_q,

                "keys": [key],

                "original_idx": q_idx

            })

            current_q_counter += 1

            

        elif item["type"] == "cluster":

            header = item["header"]

            sub_qs = item["questions"]

            sub_items_data = []

            sub_keys = []

            

            # Trộn nội dung các câu hỏi con

            for sub_q_blocks in sub_qs:

                q_idx = current_q_counter + 1

                allow_opt = config.get("shuffle_opt_global", True)

                if q_idx in fixed_opt_set: allow_opt = False

                

                new_q, key = process_single_question_logic(sub_q_blocks, part_type, allow_opt)

                sub_items_data.append((new_q, key))

                current_q_counter += 1

            

            # Trộn thứ tự câu hỏi con bên trong nhóm (Mặc định có, trừ khi tắt global shuffle)

            if config.get("shuffle_pos_global", True):

                random.shuffle(sub_items_data)

                

            # Ghép lại thành khối blocks

            cluster_final_blocks = header.copy()

            for sq, k in sub_items_data:

                cluster_final_blocks.extend(sq)

                sub_keys.append(k)

                

            processed_items.append({

                "type": "cluster",

                "blocks": cluster_final_blocks,

                "keys": sub_keys,

                "original_idx": current_q_counter - len(sub_qs) + 1 

            })



    # 2. Trộn vị trí các Item (Question vs Cluster)

    fixed_map = {}

    movable = []

    

    for i, item_data in enumerate(processed_items):

        is_fixed = False

        if not config.get("shuffle_pos_global", True): is_fixed = True

        if item_data["original_idx"] in fixed_pos_set: is_fixed = True

        

        # LOGIC MỚI: Nếu chọn "Cố định nhóm" và đây là nhóm -> Fix vị trí

        if fix_group_pos and item_data["type"] == "cluster":

            is_fixed = True

            

        if is_fixed:

            fixed_map[i] = item_data

        else:

            movable.append(item_data)

            

    random.shuffle(movable)

    

    final_blocks = intro.copy()

    final_keys = []

    

    movable_idx = 0

    total_items = len(processed_items)

    

    # Ghép lại danh sách items theo thứ tự mới

    final_item_list = []

    for i in range(total_items):

        if i in fixed_map:

            final_item_list.append(fixed_map[i])

        else:

            final_item_list.append(movable[movable_idx])

            movable_idx += 1

            

    # 3. Duyệt danh sách items để gộp blocks và đánh lại số câu

    q_counter = 0

    

    # Helper format layout

    def flush_q_group(group, p_type):

        if not group: return []

        if p_type == "PHAN1":

            relabel_mcq_options(group)

            return format_mcq_layout(group)

        elif p_type == "PHAN2":

            relabel_tf_options(group)

            return group

        return group



    # Do các item đã được sắp xếp, ta chỉ cần duỗi blocks ra và đánh số

    # Vẫn phải cẩn thận header của cluster

    

    for item in final_item_list:

        final_keys.extend(item["keys"])

        

        if item["type"] == "question":

            # Câu đơn: format layout và đánh số

            q_blocks = item["blocks"]

            if q_blocks:

                q_counter += 1

                update_question_label(q_blocks[0], f"Câu {q_counter}.")

                # Format layout

                formatted_blocks = flush_q_group(q_blocks, part_type)

                final_blocks.extend(formatted_blocks)

                

        elif item["type"] == "cluster":

            # Cụm: Header + nhiều câu hỏi

            # Header không cần format layout

            # Các câu hỏi con cần format và đánh số

            

            # Tách lại header và câu hỏi con trong blocks đã gộp

            # Dựa vào "Câu ..." để nhận biết điểm bắt đầu câu hỏi con

            

            # Cách đơn giản: Duyệt qua blocks của cluster

            c_blocks = item["blocks"]

            current_sub_q = []

            

            for blk in c_blocks:

                txt = get_text(blk)

                if re.match(r'^Câu\s*\d+\b', txt):

                    # Nếu có câu trước đó đang gom -> flush

                    if current_sub_q:

                        final_blocks.extend(flush_q_group(current_sub_q, part_type))

                        current_sub_q = []

                    

                    q_counter += 1

                    update_question_label(blk, f"Câu {q_counter}.")

                    current_sub_q.append(blk)

                else:

                    if current_sub_q:

                        current_sub_q.append(blk)

                    else:

                        # Đây là header

                        final_blocks.append(blk)

            

            if current_sub_q:

                final_blocks.extend(flush_q_group(current_sub_q, part_type))



    return final_blocks, final_keys



def shuffle_docx_logic(file_bytes, shuffle_mode, header_info, ma_de_str="", config=None):

    if config is None: config = {}

    input_buffer = io.BytesIO(file_bytes)

    keys_by_part = {}

    

    with zipfile.ZipFile(input_buffer, 'r') as zin:

        doc_xml = zin.read("word/document.xml").decode('utf-8')

        dom = minidom.parseString(doc_xml)

        body = dom.getElementsByTagNameNS(W_NS, "body")[0]

        

        blocks = []

        other_nodes = []

        for child in list(body.childNodes):

            if child.nodeType == child.ELEMENT_NODE and child.localName in ["p", "tbl"]: blocks.append(child)

            elif child.nodeType == child.ELEMENT_NODE: other_nodes.append(child)

            body.removeChild(child)

            

        new_blocks = []

        

        p1 = find_part_index(blocks, 1)

        p2 = find_part_index(blocks, 2)

        p3 = find_part_index(blocks, 3)

        

        if shuffle_mode != "auto" or (p1 == -1 and p2 == -1 and p3 == -1):

            intro, qs = parse_questions_in_range(blocks, 0, len(blocks)) # qs is structure, but process_part handles blocks

            # Reuse logic

            p_type = "PHAN1" if shuffle_mode == "mcq" or shuffle_mode == "auto" else "PHAN2"

            nb, k = process_part(blocks, 0, len(blocks), p_type, 0, config)

            new_blocks = nb

            keys_by_part['MCQ_ALL' if p_type == "PHAN1" else 'TF_ALL'] = k

            

        else:

            cursor = 0

            current_global_q_idx = 0 

            

            if p1 >= 0:

                new_blocks.extend(blocks[cursor:p1+1])

                cursor = p1 + 1

                end1 = p2 if p2 >= 0 else len(blocks)

                nb, k = process_part(blocks, cursor, end1, "PHAN1", current_global_q_idx, config)

                new_blocks.extend(nb)

                keys_by_part['PHAN1'] = k

                current_global_q_idx += len(k)

                cursor = end1

                

            if p2 >= 0:

                new_blocks.append(blocks[p2])

                cursor = p2 + 1

                end2 = p3 if p3 >= 0 else len(blocks)

                nb, k = process_part(blocks, cursor, end2, "PHAN2", current_global_q_idx, config)

                new_blocks.extend(nb)

                keys_by_part['PHAN2'] = k

                current_global_q_idx += len(k)

                cursor = end2

                

            if p3 >= 0:

                new_blocks.append(blocks[p3])

                cursor = p3 + 1

                nb, k = process_part(blocks, cursor, len(blocks), "PHAN3", current_global_q_idx, config)

                new_blocks.extend(nb)

                keys_by_part['PHAN3'] = k



        # --- Header / Footer / Zip ---

        if ma_de_str:

            p_ma = dom.createElementNS(W_NS, "w:p")

            p_ma_pr = dom.createElementNS(W_NS, "w:pPr")

            jc = dom.createElementNS(W_NS, "w:jc")

            jc.setAttributeNS(W_NS, "w:val", "right")

            p_ma_pr.appendChild(jc)

            p_ma.appendChild(p_ma_pr)

            r = dom.createElementNS(W_NS, "w:r")

            t = dom.createElementNS(W_NS, "w:t")

            rPr = dom.createElementNS(W_NS, "w:rPr")

            b = dom.createElementNS(W_NS, "w:b")

            rPr.appendChild(b)

            r.appendChild(rPr)

            t.appendChild(dom.createTextNode(f"Mã đề: {ma_de_str}"))

            r.appendChild(t)

            p_ma.appendChild(r)

            

            add_header_to_body(dom, body, header_info)

            if header_info.get("enable"):

                if body.childNodes.length > 1: body.insertBefore(p_ma, body.childNodes[1])

                else: body.appendChild(p_ma)

            else:

                if body.firstChild: body.insertBefore(p_ma, body.firstChild)

                else: body.appendChild(p_ma)

        else:

            add_header_to_body(dom, body, header_info)



        footer_rel_id = "rIdFooterNew"

        footer_fname = "word/footer_new.xml"

        sectPrs = body.getElementsByTagNameNS(W_NS, "sectPr")

        if sectPrs: sectPr = sectPrs[-1]

        else:

            sectPr = dom.createElementNS(W_NS, "w:sectPr")

            body.appendChild(sectPr)

        for child in list(sectPr.childNodes):

            if child.localName == "footerReference": sectPr.removeChild(child)

        fr = dom.createElementNS(W_NS, "w:footerReference")

        fr.setAttributeNS(W_NS, "w:type", "default")

        fr.setAttributeNS(R_NS, "r:id", footer_rel_id)

        sectPr.appendChild(fr)



        for b in new_blocks: body.appendChild(b)

        for n in other_nodes: body.appendChild(n)

        

        output_buffer = io.BytesIO()

        with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:

            footer_xml = create_footer_xml_content(ma_de_str)

            zout.writestr(footer_fname, footer_xml.encode('utf-8'))

            for item in zin.infolist():

                if item.filename == "word/document.xml":

                    zout.writestr(item, dom.toxml().encode('utf-8'))

                elif item.filename == "[Content_Types].xml":

                    ct_xml = zin.read(item).decode('utf-8')

                    ct_dom = minidom.parseString(ct_xml)

                    types = ct_dom.getElementsByTagName("Types")[0]

                    ov = ct_dom.createElement("Override")

                    ov.setAttribute("PartName", "/word/footer_new.xml")

                    ov.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")

                    types.appendChild(ov)

                    zout.writestr(item, ct_dom.toxml().encode('utf-8'))

                elif item.filename == "word/_rels/document.xml.rels":

                    rels_xml = zin.read(item).decode('utf-8')

                    rels_dom = minidom.parseString(rels_xml)

                    relationships = rels_dom.getElementsByTagName("Relationships")[0]

                    rel = rels_dom.createElement("Relationship")

                    rel.setAttribute("Id", footer_rel_id)

                    rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")

                    rel.setAttribute("Target", "footer_new.xml")

                    relationships.appendChild(rel)

                    zout.writestr(item, rels_dom.toxml().encode('utf-8'))

                else:

                    zout.writestr(item, zin.read(item.filename))

                    

        return output_buffer.getvalue(), keys_by_part



# --- EXCEL GENERATION ---

def generate_real_excel_xlsx(all_answers_dict):

    ma_des = sorted(list(all_answers_dict.keys()))

    if not ma_des: return b""



    headers = ["Đề \\ Câu"]

    headers.extend([str(i) for i in range(1, 41)])

    for q in range(1, 9):

        for char in ['a', 'b', 'c', 'd']: headers.append(f"{q}{char}")

    headers.extend([str(i) for i in range(1, 7)])

    

    rows_data = [headers]

    for md in ma_des:

        row = [str(md)]

        keys = all_answers_dict[md]

        mcq_list = []

        if 'PHAN1' in keys: mcq_list = keys['PHAN1']

        elif 'MCQ_ALL' in keys: mcq_list = keys['MCQ_ALL']

        row.extend((mcq_list + [""] * 40)[:40])

        

        tf_data = []

        if 'PHAN2' in keys: tf_data = keys['PHAN2']

        elif 'TF_ALL' in keys: tf_data = keys['TF_ALL']

        tf_flat = []

        for i in range(8):

            if i < len(tf_data): tf_flat.extend((tf_data[i] + [""] * 4)[:4])

            else: tf_flat.extend(["", "", "", ""])

        row.extend(tf_flat)

        

        sa_list = []

        if 'PHAN3' in keys: sa_list = keys['PHAN3']

        row.extend((sa_list + [""] * 6)[:6])

        rows_data.append(row)



    def get_col_name(n):

        string = ""

        while n >= 0:

            string = chr(n % 26 + 65) + string

            n = n // 26 - 1

        return string



    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">

<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>

<Default Extension="xml" ContentType="application/xml"/>

<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>

<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>

<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>

</Types>"""



    rels_root = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>

</Relationships>"""



    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"""



    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>

<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>

</Relationships>"""



    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>"""



    sheet_data_xml = ""

    for r_idx, row_val in enumerate(rows_data):

        row_str = f'<row r="{r_idx+1}">'

        for c_idx, cell_val in enumerate(row_val):

            col_letter = get_col_name(c_idx)

            safe_val = str(cell_val).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

            row_str += f'<c r="{col_letter}{r_idx+1}" t="inlineStr"><is><t>{safe_val}</t></is></c>'

        row_str += '</row>'

        sheet_data_xml += row_str



    sheet1_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

<sheetData>{sheet_data_xml}</sheetData>

</worksheet>"""



    output = io.BytesIO()

    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:

        zf.writestr("[Content_Types].xml", content_types)

        zf.writestr("_rels/.rels", rels_root)

        zf.writestr("xl/workbook.xml", workbook_xml)

        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)

        zf.writestr("xl/styles.xml", styles_xml)

        zf.writestr("xl/worksheets/sheet1.xml", sheet1_xml)

    return output.getvalue()



def create_summary_table_xml(all_answers_dict):

    ma_des = sorted(list(all_answers_dict.keys()))

    if not ma_des: return None

    mcq_keys_map = {}

    tf_keys_map = {}

    sa_keys_map = {}

    for md in ma_des:

        k = all_answers_dict[md]

        if 'PHAN1' in k: mcq_keys_map[md] = k['PHAN1']

        elif 'MCQ_ALL' in k: mcq_keys_map[md] = k['MCQ_ALL']

        if 'PHAN2' in k: tf_keys_map[md] = k['PHAN2']

        elif 'TF_ALL' in k: tf_keys_map[md] = k['TF_ALL']

        if 'PHAN3' in k: sa_keys_map[md] = k['PHAN3']



    def make_p(text, bold=False, align='center', size=None):

        sz_tag = f'<w:sz w:val="{size}"/>' if size else ''

        b_tag = '<w:b/>' if bold else ''

        safe_text = str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        return f'<w:p><w:pPr><w:jc w:val="{align}"/></w:pPr><w:r><w:rPr>{b_tag}{sz_tag}</w:rPr><w:t>{safe_text}</w:t></w:r></w:p>'

    def make_tc(content, width=None):

        w_tag = f'<w:tcW w:w="{width}" w:type="dxa"/>' if width else '<w:tcW w:w="0" w:type="auto"/>'

        return f'<w:tc><w:tcPr>{w_tag}</w:tcPr>{content}</w:tc>'



    body_content = ""

    if mcq_keys_map:

        num_mcq = len(mcq_keys_map[ma_des[0]])

        body_content += make_p("PHẦN I: TRẮC NGHIỆM", bold=True, align='left', size='28')

        row_cells = make_tc(make_p("Câu \\ Mã", bold=True), width=1200)

        for md in ma_des: row_cells += make_tc(make_p(str(md), bold=True), width=800)

        tbl1_rows = f'<w:tr>{row_cells}</w:tr>'

        for i in range(num_mcq):

            row_cells = make_tc(make_p(str(i+1), bold=True))

            for md in ma_des:

                ans = mcq_keys_map[md][i] if i < len(mcq_keys_map[md]) else ""

                row_cells += make_tc(make_p(ans))

            tbl1_rows += f'<w:tr>{row_cells}</w:tr>'

        body_content += f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>{tbl1_rows}</w:tbl><w:p/>'



    if tf_keys_map:

        body_content += make_p("PHẦN II: ĐÚNG SAI", bold=True, align='left', size='28')

        row_cells = ""

        headers = ["Mã đề", "Câu", "Ý a", "Ý b", "Ý c", "Ý d"]

        widths = [1000, 800, 800, 800, 800, 800]

        for idx, h in enumerate(headers): row_cells += make_tc(make_p(h, bold=True), width=widths[idx])

        tbl2_rows = f'<w:tr>{row_cells}</w:tr>'

        for md in ma_des:

            tf_data = tf_keys_map[md]

            for i, ans_list in enumerate(tf_data):

                md_text = str(md)

                row_cells = make_tc(make_p(md_text)) + make_tc(make_p(str(i+1), bold=True))

                for char_idx in range(4):

                    val = ans_list[char_idx] if char_idx < len(ans_list) else ""

                    row_cells += make_tc(make_p(val))

                tbl2_rows += f'<w:tr>{row_cells}</w:tr>'

        body_content += f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>{tbl2_rows}</w:tbl><w:p/>'



    if sa_keys_map:

        body_content += make_p("PHẦN III: TRẢ LỜI NGẮN", bold=True, align='left', size='28')

        row_cells = make_tc(make_p("Câu \\ Mã", bold=True), width=1200)

        for md in ma_des: row_cells += make_tc(make_p(str(md), bold=True), width=1500)

        tbl3_rows = f'<w:tr>{row_cells}</w:tr>'

        num_sa = len(sa_keys_map[ma_des[0]])

        for i in range(num_sa):

            row_cells = make_tc(make_p(str(i+1), bold=True))

            for md in ma_des:

                ans = sa_keys_map[md][i] if i < len(sa_keys_map[md]) else ""

                row_cells += make_tc(make_p(ans))

            tbl3_rows += f'<w:tr>{row_cells}</w:tr>'

        body_content += f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>{tbl3_rows}</w:tbl>'



    doc_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

    <w:document xmlns:w="{W_NS}">

        <w:body>

            <w:p><w:pPr><w:jc w:val="center"/><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>BẢNG ĐÁP ÁN TỔNG HỢP</w:t></w:r></w:p>

            {body_content}

        </w:body>

    </w:document>

    """

    return doc_xml



def generate_summary_docx(file_bytes, all_answers_dict):

    input_buffer = io.BytesIO(file_bytes)

    output_buffer = io.BytesIO()

    table_xml_str = create_summary_table_xml(all_answers_dict)

    if not table_xml_str: return io.BytesIO(b"") 

    with zipfile.ZipFile(input_buffer, 'r') as zin:

        with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:

            for item in zin.infolist():

                if item.filename == "word/document.xml":

                    zout.writestr(item, table_xml_str.encode('utf-8'))

                else:

                    zout.writestr(item, zin.read(item.filename))

    return output_buffer.getvalue()



# ==================== PHẦN 2: GIAO DIỆN NGƯỜI DÙNG (GUI) ====================



class TronDeApp:

    def __init__(self, root):

        self.root = root

        self.root.title("Phần Mềm Trộn Đề Word Pro - AIOMT")

        self.root.geometry("800x750")

        

        style = ttk.Style()

        style.theme_use('clam')

        style.configure("TLabel", font=("Arial", 10))

        style.configure("TButton", font=("Arial", 10, "bold"))

        style.configure("TRadiobutton", font=("Arial", 10))

        style.configure("TCheckbutton", font=("Arial", 10))

        

        self.file_path = tk.StringVar()

        self.mode_var = tk.StringVar(value="auto")

        self.num_ver_var = tk.IntVar(value=4)

        

        # Header vars

        self.use_header = tk.BooleanVar(value=True)

        self.so_gd = tk.StringVar(value="SỞ GD&ĐT HÀ NỘI")

        self.truong = tk.StringVar(value="TRƯỜNG THPT CHU VĂN AN")

        self.ky_thi = tk.StringVar(value="ĐỀ KIỂM TRA GIỮA KỲ I")

        self.mon_thi = tk.StringVar(value="TOÁN 12")

        self.thoi_gian = tk.StringVar(value="90 phút")

        self.nam_hoc = tk.StringVar(value="2024 - 2025")

        

        # New Config vars

        self.ma_de_mode = tk.StringVar(value="auto") # auto / manual

        self.ma_de_manual_str = tk.StringVar(value="")

        

        self.shuffle_pos_global = tk.BooleanVar(value=True)

        self.shuffle_opt_global = tk.BooleanVar(value=True)

        self.fixed_pos_str = tk.StringVar(value="")

        self.fixed_opt_str = tk.StringVar(value="")

        self.fix_group_pos = tk.BooleanVar(value=True) # Checkbox mới



        self._build_ui()



    def _build_ui(self):

        # 1. File

        frame_file = tk.LabelFrame(self.root, text="1. Chọn File Đề Gốc (.docx)", font=("Arial", 10, "bold"))

        frame_file.pack(fill="x", padx=10, pady=5)

        entry = tk.Entry(frame_file, textvariable=self.file_path, width=60)

        entry.pack(side="left", padx=10, pady=10)

        tk.Button(frame_file, text="Chọn File...", command=self.select_file, bg="#0d9488", fg="white").pack(side="left", padx=5)



        # 2. Header

        frame_header = tk.LabelFrame(self.root, text="2. Cấu hình Tiêu đề (Header)", font=("Arial", 10, "bold"))

        frame_header.pack(fill="x", padx=10, pady=5)

        cb = tk.Checkbutton(frame_header, text="Tự động thêm bảng tiêu đề vào file kết quả", variable=self.use_header, command=self.toggle_header)

        cb.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=5)

        self.header_widgets = []

        ttk.Label(frame_header, text="Sở/Phòng GD&ĐT:").grid(row=1, column=0, sticky="e", padx=5)

        e1 = ttk.Entry(frame_header, textvariable=self.so_gd, width=35)

        e1.grid(row=1, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(frame_header, text="Tên Trường:").grid(row=2, column=0, sticky="e", padx=5)

        e2 = ttk.Entry(frame_header, textvariable=self.truong, width=35)

        e2.grid(row=2, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(frame_header, text="Tên Kỳ Thi:").grid(row=1, column=2, sticky="e", padx=5)

        e3 = ttk.Entry(frame_header, textvariable=self.ky_thi, width=35)

        e3.grid(row=1, column=3, sticky="w", padx=5, pady=2)

        ttk.Label(frame_header, text="Môn Thi:").grid(row=2, column=2, sticky="e", padx=5)

        e4 = ttk.Entry(frame_header, textvariable=self.mon_thi, width=35)

        e4.grid(row=2, column=3, sticky="w", padx=5, pady=2)

        ttk.Label(frame_header, text="Thời gian:").grid(row=3, column=2, sticky="e", padx=5)

        e5 = ttk.Entry(frame_header, textvariable=self.thoi_gian, width=35)

        e5.grid(row=3, column=3, sticky="w", padx=5, pady=2)

        ttk.Label(frame_header, text="Năm học:").grid(row=3, column=0, sticky="e", padx=5)

        e6 = ttk.Entry(frame_header, textvariable=self.nam_hoc, width=35)

        e6.grid(row=3, column=1, sticky="w", padx=5, pady=2)

        self.header_widgets.extend([e1, e2, e3, e4, e5, e6])



        # 3. Basic Config

        frame_cfg = tk.LabelFrame(self.root, text="3. Cấu hình Cơ bản", font=("Arial", 10, "bold"))

        frame_cfg.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_cfg, text="Chế độ:").grid(row=0, column=0, sticky="e", padx=5)

        cbox = ttk.Combobox(frame_cfg, textvariable=self.mode_var, values=["auto", "mcq", "tf"], state="readonly", width=15)

        cbox.grid(row=0, column=1, sticky="w", padx=5)

        

        # Mã đề config

        ttk.Label(frame_cfg, text="Cách tạo Mã đề:").grid(row=1, column=0, sticky="e", padx=5)

        rb_auto = ttk.Radiobutton(frame_cfg, text="Tự động (Ngẫu nhiên)", variable=self.ma_de_mode, value="auto", command=self.toggle_ma_de_input)

        rb_auto.grid(row=1, column=1, sticky="w")

        rb_manual = ttk.Radiobutton(frame_cfg, text="Tự nhập:", variable=self.ma_de_mode, value="manual", command=self.toggle_ma_de_input)

        rb_manual.grid(row=1, column=2, sticky="w")

        

        self.entry_manual_ma_de = ttk.Entry(frame_cfg, textvariable=self.ma_de_manual_str, width=30, state="disabled")

        self.entry_manual_ma_de.grid(row=1, column=3, sticky="w", padx=5)

        ttk.Label(frame_cfg, text="(VD: 101, 102, 201)").grid(row=1, column=4, sticky="w")



        # Số lượng (chỉ hiện khi auto)

        self.lbl_num_ver = ttk.Label(frame_cfg, text="Số lượng đề:")

        self.lbl_num_ver.grid(row=0, column=2, sticky="e", padx=5)

        self.spin_num_ver = tk.Spinbox(frame_cfg, from_=1, to=50, textvariable=self.num_ver_var, width=5)

        self.spin_num_ver.grid(row=0, column=3, sticky="w", padx=5)



        # 4. Advanced Config

        frame_adv = tk.LabelFrame(self.root, text="4. Cấu hình Nâng cao (Tùy chọn cố định)", font=("Arial", 10, "bold"))

        frame_adv.pack(fill="x", padx=10, pady=5)

        

        # Row 1: Global switches

        ttk.Checkbutton(frame_adv, text="Trộn thứ tự Câu hỏi", variable=self.shuffle_pos_global).grid(row=0, column=0, sticky="w", padx=10)

        ttk.Checkbutton(frame_adv, text="Trộn thứ tự Đáp án (A,B,C,D)", variable=self.shuffle_opt_global).grid(row=0, column=1, sticky="w", padx=10)

        

        # Checkbox mới

        ttk.Checkbutton(frame_adv, text="Cố định vị trí các Nhóm câu hỏi dùng chung", variable=self.fix_group_pos).grid(row=0, column=2, sticky="w", padx=10)

        

        # Row 2: Fixed Positions

        ttk.Label(frame_adv, text="Câu hỏi KHÔNG trộn vị trí (VD: 1, 2, 40):").grid(row=1, column=0, sticky="e", padx=5, pady=5)

        ttk.Entry(frame_adv, textvariable=self.fixed_pos_str, width=40).grid(row=1, column=1, columnspan=2, sticky="w", padx=5)

        

        # Row 3: Fixed Options

        ttk.Label(frame_adv, text="Câu hỏi KHÔNG trộn đáp án (VD: 1-5, 10):").grid(row=2, column=0, sticky="e", padx=5, pady=5)

        ttk.Entry(frame_adv, textvariable=self.fixed_opt_str, width=40).grid(row=2, column=1, columnspan=2, sticky="w", padx=5)



        # 5. Run

        self.btn_run = tk.Button(self.root, text="🚀 BẮT ĐẦU TRỘN ĐỀ", command=self.run_process, 

                                 bg="#0f766e", fg="white", font=("Arial", 12, "bold"), height=2)

        self.btn_run.pack(fill="x", padx=50, pady=10)

        

        tk.Label(self.root, text="Gạch chân hoặc Tô đỏ đáp án đúng để tạo bảng đáp án", fg="red").pack(side="bottom", pady=5)



    def select_file(self):

        f = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])

        if f: self.file_path.set(f)

        

    def toggle_header(self):

        state = "normal" if self.use_header.get() else "disabled"

        for w in self.header_widgets: w.config(state=state)

        

    def toggle_ma_de_input(self):

        mode = self.ma_de_mode.get()

        if mode == "manual":

            self.entry_manual_ma_de.config(state="normal")

            self.spin_num_ver.config(state="disabled")

        else:

            self.entry_manual_ma_de.config(state="disabled")

            self.spin_num_ver.config(state="normal")



    def run_process(self):

        inp = self.file_path.get()

        if not inp or not os.path.exists(inp):

            messagebox.showerror("Lỗi", "Vui lòng chọn file hợp lệ!")

            return

            

        try:

            self.btn_run.config(state="disabled", text="Đang xử lý...")

            self.root.update()

            

            with open(inp, "rb") as f:

                file_bytes = f.read()

            

            base_name = os.path.basename(inp).replace(".docx", "")

            mode = self.mode_var.get()

            

            header_info = {

                "enable": self.use_header.get(),

                "so_gd": self.so_gd.get(),

                "truong": self.truong.get(),

                "ky_thi": self.ky_thi.get(),

                "mon_thi": self.mon_thi.get(),

                "thoi_gian": self.thoi_gian.get(),

                "nam_hoc": self.nam_hoc.get()

            }

            

            config = {

                "shuffle_pos_global": self.shuffle_pos_global.get(),

                "shuffle_opt_global": self.shuffle_opt_global.get(),

                "fixed_pos_set": parse_range_string(self.fixed_pos_str.get()),

                "fixed_opt_set": parse_range_string(self.fixed_opt_str.get()),

                "fix_group_pos": self.fix_group_pos.get()

            }

            

            ma_de_list = []

            if self.ma_de_mode.get() == "manual":

                raw_str = self.ma_de_manual_str.get()

                parts = [s.strip() for s in raw_str.split(',') if s.strip()]

                if not parts:

                    messagebox.showerror("Lỗi", "Vui lòng nhập ít nhất 1 mã đề!")

                    return

                ma_de_list = parts

            else:

                num = self.num_ver_var.get()

                start_code = 101

                ma_de_list = [str(start_code + i) for i in range(num)]

            

            all_answers_summary = {}

            

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:

                ans_key_text = "MÃ ĐỀ\n"

                

                for ma_de in ma_de_list:

                    out_bytes, keys_by_part = shuffle_docx_logic(file_bytes, mode, header_info, ma_de, config)

                    all_answers_summary[ma_de] = keys_by_part

                    

                    fname = f"{base_name}_MaDe{ma_de}.docx"

                    zout.writestr(fname, out_bytes)

                    ans_key_text += f"- Đề {fname}: (Đã trộn)\n"

                

                zout.writestr("DS_MaDe.txt", ans_key_text)

                

                try:

                    summary_bytes = generate_summary_docx(file_bytes, all_answers_summary)

                    zout.writestr("Dap_an_tong_hop.docx", summary_bytes)

                except Exception as e: print(f"Lỗi tạo bảng Word: {e}")

                    

                try:

                    excel_bytes = generate_real_excel_xlsx(all_answers_summary)

                    zout.writestr("Dap_an_Excel_Chuan.xlsx", excel_bytes)

                except Exception as e: print(f"Lỗi tạo Excel: {e}")

            

            s_path = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("Zip", "*.zip")], initialfile=f"{base_name}_TronDe_Pro")

            if s_path:

                with open(s_path, "wb") as f: f.write(zip_buffer.getvalue())

                messagebox.showinfo("Thành công", f"Đã trộn {len(ma_de_list)} mã đề!\nLưu tại: {s_path}")

            

        except Exception as e:

            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {str(e)}")

            print(e)

        finally:

            self.btn_run.config(state="normal", text="🚀 BẮT ĐẦU TRỘN ĐỀ")



if __name__ == "__main__":

    root = tk.Tk()

    app = TronDeApp(root)

    root.mainloop()
