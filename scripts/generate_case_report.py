# -*- coding: utf-8 -*-
"""
generate_case_report.py
========================
按「案例检索报告」模板格式，从裁判文书文本 + 检索要求生成 Word 文档。

严格对齐模板（【混同用工支持承担连带责任】案例检索报告.docx）的单元格排版：
  · 争议焦点：精简为一句问/述（优先取同名 .md 笔记里手写的争议焦点）
  · 裁判观点：按文书层级拆分为「再审申请人意见 / 一审法院查明事实 /
              一审法院裁判观点 / 二审法院裁判观点 / 再审审查裁判观点」等
              子段落，标签行加粗、内容行常规，各自成行
  · 本案可参考观点：按要点分列（优先取同名 .md 笔记里的「可参考观点」多行）

用法示例：
  python generate_case_report.py 文书1.docx 文书2.docx --req "混同用工连带责任检索" --out 报告.docx

  # 为每份文书配一份同名笔记（.md/.txt），可精细控制字段：
  #   争议焦点：xxx（精简一句）
  #   可参考观点：
  #   第一点……
  #   第二点……
  python generate_case_report.py 文书1.docx --out 报告.docx

依赖：python-docx  (Anaconda 自带；其他环境: pip install python-docx)
"""
import argparse
import os
import re

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 默认模板：用户提供的案例检索报告。可在 --template 覆盖。
DEFAULT_TEMPLATE = r"C:/Users/茶声/Desktop/裕景律所/【混同用工支持承担连带责任】案例检索报告.docx"

FONT = "仿宋"  # 模板全篇字体
SIZE_BODY = 12  # 正文 12pt
SIZE_HEAD = 14  # 关键词/列名 14pt


# ---------- 格式工具 ----------
def set_run_font(run, size=None, bold=None, name=FONT):
    """设置 run 的字体（含中文 eastAsia）、字号、加粗。"""
    if name:
        run.font.name = name
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        rfonts.set(qn("w:eastAsia"), name)
        rfonts.set(qn("w:ascii"), name)
        rfonts.set(qn("w:hAnsi"), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


def set_cell_margins(tbl, top=0, bottom=0, left=108, right=108):
    """设置表格单元格左右边距（模板：左右约 0.19cm = 108 twips）。"""
    tbl_pr = tbl._tbl.tblPr
    margins = OxmlElement("w:tblCellMar")
    for tag, val in (("w:top", top), ("w:bottom", bottom), ("w:left", left), ("w:right", right)):
        node = OxmlElement(tag)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        margins.append(node)
    tbl_pr.append(margins)


def set_table_borders(tbl, sz=4):
    """设置表格边框：全部 single，0.5pt（sz 单位为 1/8 pt）。"""
    tbl_pr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(sz))
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        borders.append(e)
    tbl_pr.append(borders)


def clear_para(p):
    """删除段落内现有 run（用于复用首个空段落时清空残留空 run）。"""
    for r in list(p.runs):
        r._element.getparent().remove(r._element)


def set_cell_vcenter(cell):
    tcpr = cell._tc.get_or_add_tcPr()
    v = OxmlElement("w:vAlign")
    v.set(qn("w:val"), "center")
    tcpr.append(v)


def fill_single(cell, text, size, bold, align):
    """写入单元格（单段落单 run）：用于序号/法院/案号/争议焦点/列名/关键词。"""
    cell.text = text if text else ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.runs[0] if p.runs else p.add_run("")
    set_run_font(run, size=size, bold=bold)
    set_cell_vcenter(cell)


def fill_opinion(cell, sections):
    """
    裁判观点单元格：sections = [(标签, 内容), ...]
    每个子段 = 一个加粗标签段 + 一个常规内容段，各自成行。
    """
    cell.text = ""
    first = True
    for label, content in sections:
        # 加粗标签段（复用首个空段落）
        if first:
            p = cell.paragraphs[0]
            clear_para(p)
            first = False
        else:
            p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_run_font(p.add_run((label or "") + "："), size=SIZE_BODY, bold=True)
        # 常规内容段
        cp = cell.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_run_font(cp.add_run(content or ""), size=SIZE_BODY, bold=False)
    set_cell_vcenter(cell)


def fill_points(cell, points):
    """
    本案可参考观点单元格：points = [要点, ...]，每条单独成段并以 1、2、3、 编号。
    """
    cell.text = ""
    first = True
    n = 0
    for pt in points:
        pt = (pt or "").strip()
        # 去掉要点自身可能带的前导编号，避免重复编号
        pt = re.sub(r"^\s*[(（]?\d+[)、.．]\s*", "", pt)
        if not pt:
            continue
        n += 1
        text = f"{n}、{pt}"
        if first:
            p = cell.paragraphs[0]
            clear_para(p)
            first = False
        else:
            p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_run_font(p.add_run(text), size=SIZE_BODY, bold=False)
    if first:  # 没有任何要点
        p = cell.paragraphs[0]
        clear_para(p)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_run_font(p.add_run("（请补充本案可参考观点）"), size=SIZE_BODY, bold=False)
    set_cell_vcenter(cell)


# ---------- 文书解析 ----------
def read_text(path):
    """自动识别编码读取文本（UTF-8 → GB18030/GBK → 容错）。"""
    for enc in ("utf-8", "gb18030", "gbk"):
        try:
            return open(path, encoding=enc).read()
        except (UnicodeDecodeError, UnicodeError):
            continue
    return open(path, encoding="utf-8", errors="ignore").read()


def load_doc_text(path):
    """读取裁判文书文本：.docx 提取段落文本，其余按纯文本读取。"""
    if path.lower().endswith(".docx"):
        try:
            d = Document(path)
            return "\n".join(p.text for p in d.paragraphs)
        except Exception:
            pass
    return read_text(path)


def case_level(case_no):
    """由案号后缀推断审理层级，作为兜底裁判观点标签。"""
    if "民初" in case_no:
        return "一审法院裁判观点"
    if "民终" in case_no:
        return "二审法院裁判观点"
    if "民申" in case_no:
        return "再审审查裁判观点"
    if "民再" in case_no:
        return "再审法院裁判观点"
    return "裁判观点"


def clean_chunk(chunk, limit=500):
    """清洗裁判观点片段：去首尾空白、截断程序性尾部、限长。"""
    chunk = chunk.strip()
    # 截断到程序性尾部之前（审判长/本院/如不服本判决等）
    stop = re.search(r"\n{0,1}(?:审判长|审判员|代理审判员|二〇|二○|本(?:判决|裁定|调解书)|如不服本判决|（此页无正文）|附录|附：)", chunk)
    if stop:
        chunk = chunk[:stop.start()]
    chunk = re.sub(r"\s+", " ", chunk).strip()
    if len(chunk) > limit:
        cut = chunk.rfind("。", 0, limit)
        chunk = (chunk[:cut + 1] if cut > limit * 0.5 else chunk[:limit] + "…")
    return chunk


# 裁判文书层级标记（按出现顺序抓取；「认为/查明」后兼容逗号或冒号）
SEG_RE = re.compile(
    r"(再审申请人[^\n：:。]{0,30}?称[:：]|"
    r"(?:一审|原审)[^\n：:。]{0,12}?法院查明[:：，,]|"
    r"(?:一审|原审)[^\n：:。]{0,12}?法院认为[:：，,]|"
    r"二审[^\n：:。]{0,12}?法院认为[:：，,]|"
    r"(?:再审法院|本院再审|本院经审查|再审)[^\n：:。]{0,12}?法院认为[:：，,]|"
    r"本院认为[:：，,])"
)


def _label_for(marker):
    if "再审申请人" in marker and "称" in marker:
        return "再审申请人意见"
    if "查明" in marker:
        return "二审法院查明事实" if "二审" in marker else "一审法院查明事实"
    if "认为" in marker:
        if "二审" in marker:
            return "二审法院裁判观点"
        if "一审" in marker or "原审" in marker:
            return "一审法院裁判观点"
        # 仅当明确含「再审」或「本院再审/本院经审查」才归再审；裸「本院认为」返回 None 兜底
        if "再审" in marker or "本院再审" in marker or "本院经审查" in marker:
            return "再审审查裁判观点"
        return None  # 裸「本院认为」，稍后按审级兜底
    return None


def extract_opinion_sections(text, case_no):
    """
    将裁判文书正文拆分为裁判观点子段，返回 [(标签, 内容), ...]。
    顺序天然为：再审申请人意见 → 查明事实 → 一审观点 → 二审观点 → 再审观点。
    """
    matches = list(SEG_RE.finditer(text))
    if not matches:
        return []
    sections = []
    for i, m in enumerate(matches):
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        label = _label_for(m.group(0))
        content = clean_chunk(text[start:end])
        if not content:
            continue
        # 一审法院认为常先列「本案争议焦点：1、…」，该部分已在争议焦点列呈现，剔除避免重复
        if label and label.startswith("一审法院") and re.match(r"^(本案)?争议焦点", content):
            cut = re.search(r"[。！]\s*(?:当事人|关于|对于|本院|双方|经查|根据)", content)
            if cut:
                content = content[cut.start() + 1:].strip()
        if not content:
            continue
        sections.append([label, content])
    # 为裸「本院认为」补标签（按审级兜底）
    for s in sections:
        if s[0] is None:
            s[0] = case_level(case_no)
    # 去除连续重复标签
    seen = []
    for s in sections:
        if seen and seen[-1][0] == s[0]:
            seen[-1][1] += " " + s[1]
        else:
            seen.append(s)
    return [(a, b) for a, b in seen]


def extract_case(text):
    """从裁判文书文本中规则提取：案号、裁判法院、争议焦点（精简）。"""
    # 案号（兼容全/半角括号）
    m = re.search(r"[\(（]\d{4}[\)）][^\n]*?(?:民|刑|行|商|终|初|再|申)[^\n]*?号", text)
    case_no = m.group(0).strip() if m else ""

    # 裁判法院
    m = re.search(r"([一-龥]{2,}?(?:高级|中级)?人民法院)", text)
    court = m.group(1).strip() if m else ""

    # 争议焦点（精简：取首个短句，去掉编号前缀）
    focus = ""
    mf = re.search(r"争议焦点[：:为]?\s*([^\n。；]{4,90})", text)
    if mf:
        focus = mf.group(1).strip()
        focus = re.sub(r"^[(（]?\d+[)、.]\s*", "", focus)
        focus = re.sub(r"^[一二三四五六七八九十]+、\s*", "", focus)
        if len(focus) > 70:
            focus = focus[:70] + "…"

    return dict(case_no=case_no, court=court, focus=focus)


def parse_notes(path):
    """解析同名笔记（.md/.txt）里的「争议焦点:」「可参考观点:」（支持多行要点）。"""
    d = {}
    if not path or not os.path.exists(path):
        return d
    txt = read_text(path)
    m = re.search(r"争议焦点[：:]\s*([^\n]+)", txt)
    d["focus"] = m.group(1).strip() if m else ""
    # 可参考观点：从「可参考观点:」取到文件末或下一个一级字段，按行拆要点
    m = re.search(r"可参考观点[：:]\s*\n([\s\S]+?)(?:\n\s*\n\s*[一-龥]{2,}[：:]|\Z)", txt)
    if not m:
        m = re.search(r"可参考观点[：:]\s*([^\n]+(?:\n[^\n]+)*)", txt)
    if m:
        lines = [ln.strip(" ·•-—\t").strip() for ln in m.group(1).split("\n")]
        d["insight"] = [ln for ln in lines if ln]
    else:
        d["insight"] = []
    return d


# ---------- 主流程 ----------
def main():
    ap = argparse.ArgumentParser(description="按案例检索报告模板格式生成 Word")
    ap.add_argument("docs", nargs="+", help="裁判文书 .txt/.docx 文件（可多个）")
    ap.add_argument("--req", help="检索要求 / 分析要点（填入关键词行）")
    ap.add_argument("--req-file", help="检索要求文件（.txt/.md）")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE, help="模板 docx 路径")
    ap.add_argument("--out", default="案例检索报告_生成.docx", help="输出 docx 路径")
    args = ap.parse_args()

    req_text = args.req or ""
    if args.req_file and os.path.exists(args.req_file):
        req_text = read_text(args.req_file).strip()

    # 1) 打开模板，继承页面（横向A4）、页脚、默认字体
    doc = Document(args.template)

    # 2) 删除模板中原有表格（保留非表格段落如标题）
    body = doc.element.body
    for t in body.findall(qn("w:tbl")):
        body.remove(t)

    # 3) 新建表格：行数 = 关键词行 + 列名行 + 每案一行
    n = len(args.docs)
    table = doc.add_table(rows=2 + n, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    set_cell_margins(table)

    # 4) 列宽（横向 A4 内容宽约 24.6cm）
    widths = [Cm(0.9), Cm(2.3), Cm(2.1), Cm(3.6), Cm(4.6), Cm(4.6)]
    for i, w in enumerate(widths):
        for r in table.rows:
            r.cells[i].width = w

    # 第1行：关键词（合并 6 列）
    r0 = table.rows[0]
    merged = r0.cells[0].merge(r0.cells[5])
    kw = ("关键词：" + req_text) if req_text else "关键词：（未提供检索要求）"
    fill_single(merged, kw, SIZE_HEAD, True, WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 第2行：列名
    headers = ["序号", "裁判法院", "案号", "争议焦点", "裁判观点", "本案可参考观点"]
    for i, h in enumerate(headers):
        fill_single(table.rows[1].cells[i], h, SIZE_HEAD, True, WD_ALIGN_PARAGRAPH.CENTER)

    # 数据行
    for idx, docpath in enumerate(args.docs):
        text = load_doc_text(docpath)
        base = os.path.splitext(docpath)[0]
        notes = parse_notes(base + ".md") or parse_notes(base + ".txt")
        info = extract_case(text)

        # 争议焦点：笔记优先（精简手写），否则自动精简
        focus = notes.get("focus") or info["focus"] or "（未能自动识别）"
        if len(focus) > 70:
            focus = focus[:70] + "…"

        # 裁判观点：拆子段
        sections = extract_opinion_sections(text, info["case_no"])
        if not sections:
            mv = re.search(r"本院认为[，,：:.\s]*([\s\S]{20,600})", text)
            content = clean_chunk(mv.group(1)) if mv else "（未能自动识别裁判观点）"
            sections = [(case_level(info["case_no"]), content)]

        # 本案可参考观点：笔记多行要点优先，否则空（提示补充）
        insight = notes.get("insight") or []

        row = table.rows[2 + idx].cells
        fill_single(row[0], str(idx + 1), SIZE_BODY, False, WD_ALIGN_PARAGRAPH.CENTER)
        fill_single(row[1], info["court"], SIZE_BODY, False, WD_ALIGN_PARAGRAPH.CENTER)
        fill_single(row[2], info["case_no"], SIZE_BODY, False, WD_ALIGN_PARAGRAPH.LEFT)
        fill_single(row[3], focus, SIZE_BODY, False, WD_ALIGN_PARAGRAPH.LEFT)
        fill_opinion(row[4], sections)
        fill_points(row[5], insight)

    doc.save(args.out)
    print(f"已生成：{args.out}  | 案例数：{n}  | 模板：{os.path.basename(args.template)}")


if __name__ == "__main__":
    main()
