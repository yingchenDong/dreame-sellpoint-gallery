#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
卖点图渲染验收标准参考图库 - 网站自动更新脚本
================================================
功能：
  1. 解析卖点表格 xlsx（无需安装第三方库），自动更新页面中的
     核心验收标准 / 细分类别 / 卖点参考（含高配、中配、低配、主图KV）文案；
  2. 可选：从图片压缩包 zip 解压新增参考图（自动按 文件名前缀 / 子目录 分组追加到图库）；
  3. 可选：从新版 PDF 提取内嵌图（自动按源页分组追加到图库）。

用法（在仓库根目录执行）：
  python3 tools/update_site.py --xlsx "/path/验收标准.xlsx"            # 仅更新文案
  python3 tools/update_site.py --xlsx "..." --zip "/path/images.zip"   # 文案 + 图库
  python3 tools/update_site.py --xlsx "..." --pdf "/path/标准.pdf"      # 文案 + PDF 图
  python3 tools/update_site.py --all --source source/                  # 从 source/ 目录自动读取
  python3 tools/update_site.py --xlsx "..." --dry-run                  # 只打印解析结果，不写文件

更新规则：
  - 卖点文案按行号顺序与页面 .sp 区块一一对应替换；表格新增行会自动追加新卖点卡
    （无参考图时显示占位提示，可后续补图）；表格删除的行对应卡片会标记为已归档。
  - 参考图是手工映射的资产，文案更新不会改动图片；zip/pdf 更新只会【新增】分组，
    不会删除已有分组，避免误删。
"""
import argparse
import json
import re
import shutil
import sys
import zipfile
import os
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
HTML_FILE = "index.html"
GALLERY_DIR = "ref_assets/embedded"


# ---------------------------------------------------------------- xlsx 解析
def read_xlsx(path):
    """返回 {sheet: [ {row: r, cells: {A: '文本', ...}} ]}"""
    zf = zipfile.ZipFile(path)
    shared = [
        re.sub(r"<[^>]+>", "", t)
        for t in re.findall(r"<si>.*?</si>", zf.read("xl/sharedStrings.xml").decode("utf-8", "ignore"), re.S)
    ]
    sheets = {}
    for name in zf.namelist():
        if name.startswith("xl/worksheets/") and name.endswith(".xml"):
            root = ET.fromstring(zf.read(name))
            rows = []
            for r in root.findall(".//m:sheetData/m:row", NS):
                cells = {}
                for c in r.findall("m:c", NS):
                    v = c.find("m:v", NS)
                    if v is None:
                        continue
                    val = v.text or ""
                    if c.get("t") == "s":
                        try:
                            val = shared[int(val)]
                        except (ValueError, IndexError):
                            continue
                    col = re.match(r"[A-Z]+", c.get("r", "A1")).group(0)
                    cells[col] = val.strip().replace("\u00a0", " ")
                rows.append({"row": int(r.get("r")), "cells": cells})
            sheets[name.split("/")[-1]] = rows
    return sheets


def cell(sheet, row, col):
    for r in sheet:
        if r["row"] == row:
            return r["cells"].get(col, "")
    return ""


def split_lines(text):
    """单元格内多行文本 -> 行列表（去掉空行与首尾空白）"""
    return [ln.strip() for ln in re.split(r"[\n\r]", text.replace("&#xA;", "\n")) if ln.strip()]


def parse_sellpoint(block):
    """解析卖点块文本 -> {title, sub, items}"""
    lines = split_lines(block)
    if not lines:
        return None
    title, sub = lines[0], ""
    if "：" in title:
        title, sub = title.split("：", 1)
        title, sub = title.strip(), sub.strip()
    elif ":" in title:
        title, sub = title.split(":", 1)
        title, sub = title.strip(), sub.strip()
    items = [ln for ln in lines[1:]]
    return {"title": title, "sub": sub, "items": items}


def parse_site(src_xlsx):
    """从验收标准 xlsx 提取页面结构化数据"""
    sheets = read_xlsx(src_xlsx)
    sheet = sheets[list(sheets.keys())[0]]

    # 核心标准：A 列有值且 B 列有"核心关键词"的行
    standards = []
    for r in sheet:
        a, b = r["cells"].get("A", ""), r["cells"].get("B", "")
        if a and ("核心关键词" in b):
            standards.append({"name": a, "body": b})

    # 细分类别：C 列，R1 表头之后、卖点表头之前
    categories = []
    for r in sheet:
        if r["row"] <= 1:
            continue
        a = r["cells"].get("A", "")
        if "卖点图展现形式" in a or "主图展现形式" in a:
            break
        c = r["cells"].get("C", "")
        if c:
            lines = split_lines(c)
            categories.append({"name": lines[0], "desc": " ".join(lines[1:]) if len(lines) > 1 else ""})

    # 卖点：从"卖点图展现形式"行起，含"主图展现形式"行，C 列每行一个卖点
    sellpoints = []
    in_sell = False
    idx = 0
    for r in sheet:
        a = r["cells"].get("A", "")
        c = r["cells"].get("C", "")
        b = r["cells"].get("B", "")
        if "卖点图展现形式" in a or "主图展现形式" in a:
            in_sell = True  # 表头行自身也含首个卖点（如 R18 蒸汽拖地、R41 主图KV 高配），不跳过
        if not in_sell:
            continue
        if c:
            idx += 1
            sp = parse_sellpoint(c)
            if sp:
                sp["idx"] = idx
                sp["tier"] = b.split("（")[0].strip() if b and len(b) < 20 else ""
                sellpoints.append(sp)
    return {"standards": standards, "categories": categories, "sellpoints": sellpoints}


# ---------------------------------------------------------------- 新版 Excel 重建（含内嵌图）
XLSX_NS_DRAW = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def xlsx_drawing_row_map(path):
    """解析 xlsx 内嵌图锚点：返回 {excel_row: [(col, media_path), ...]}"""
    zf = zipfile.ZipFile(path)
    names = zf.namelist()
    drawing_xml = [n for n in names if n.startswith("xl/drawings/") and n.endswith(".xml")]
    if not drawing_xml:
        return {}
    rels = {}
    rels_name = "xl/drawings/_rels/" + drawing_xml[0].split("/")[-1] + ".rels"
    if rels_name in names:
        rels_root = ET.fromstring(zf.read(rels_name))
        rels = {rel.get("Id"): rel.get("Target").replace("../media/", "xl/media/") for rel in rels_root}
    drawing_root = ET.fromstring(zf.read(drawing_xml[0]))
    row_map = {}
    for anchor in drawing_root:
        tag = anchor.tag.split("}")[-1]
        if tag not in ("twoCellAnchor", "oneCellAnchor"):
            continue
        frm = anchor.find("xdr:from", XLSX_NS_DRAW)
        blip = anchor.find(".//xdr:blipFill/a:blip", XLSX_NS_DRAW)
        if blip is None or frm is None:
            continue
        rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        fcol = int(frm.find("xdr:col", XLSX_NS_DRAW).text)
        frow = int(frm.find("xdr:row", XLSX_NS_DRAW).text)
        media_path = rels.get(rid, "")
        if media_path:
            row_map.setdefault(frow + 1, []).append((fcol, media_path))
    for k in row_map:
        row_map[k].sort(key=lambda x: x[0])
    return row_map


def extract_xlsx_images(path, out_root):
    """提取 xlsx 内嵌图到 out_root/row_XX/N.ext，返回 {excel_row: [相对路径...]}"""
    row_map = xlsx_drawing_row_map(path)
    zf = zipfile.ZipFile(path)
    result = {}
    for excel_row, items in row_map.items():
        row_dir = Path(out_root) / f"row_{excel_row:02d}"
        row_dir.mkdir(parents=True, exist_ok=True)
        paths = []
        for idx, (_col, media_path) in enumerate(items, 1):
            ext = os.path.splitext(media_path)[1] or ".jpg"
            dest = row_dir / f"{idx}{ext}"
            try:
                dest.write_bytes(zf.read(media_path))
            except KeyError:
                continue
            rel = f"{out_root}/row_{excel_row:02d}/{idx}{ext}".replace("\\", "/")
            paths.append(rel)
        result[excel_row] = paths
    return result


def rebuild_sections_from_xlsx(src_xlsx, html_file=HTML_FILE):
    """新版 Excel：提取内嵌图并整体重建 category/sellpoint 区块（横向框架逐行展示）"""
    sheets = read_xlsx(src_xlsx)
    sheet = sheets[list(sheets.keys())[0]]
    # 行文本：{row: {A, B, C}}
    rows_text = {}
    for r in sheet:
        rows_text[r["row"]] = r["cells"]
    images = extract_xlsx_images(src_xlsx, "ref_assets/sellpoints")

    def row_text(row, col):
        return (rows_text.get(row) or {}).get(col, "")

    # category：卖点表头前的行
    cat_rows = []
    for r in sheet:
        a = r["cells"].get("A", "")
        if "卖点图展现形式" in a or "主图展现形式" in a:
            break
        if r["row"] <= 1:
            continue
        c = r["cells"].get("C", "")
        if c:
            cat_rows.append(r["row"])
    sp_rows = []
    in_sell = False
    for r in sheet:
        a = r["cells"].get("A", "")
        if "卖点图展现形式" in a or "主图展现形式" in a:
            in_sell = True
        if not in_sell:
            continue
        c = r["cells"].get("C", "")
        if c:
            sp_rows.append(r["row"])

    cat_parts = []
    for row in cat_rows:
        c = row_text(row, "C")
        lines = split_lines(c)
        title = lines[0] if lines else ""
        sub = " ".join(lines[1:]) if len(lines) > 1 else ""
        cat_parts.append(
            '<div class="cat reveal" data-search="%s"><b>%s</b><span>%s</span></div>'
            % (esc(title + " " + sub), esc(title), esc(sub))
        )
    sp_parts = []
    for i, row in enumerate(sp_rows, 1):
        c = row_text(row, "C")
        b = row_text(row, "B")
        lines = split_lines(c)
        title = lines[0] if lines else f"卖点 {i}"
        sub = ""
        if "：" in title:
            sub = title.split("：", 1)[1].strip()
        elif len(lines) > 1:
            sub = lines[1]
        body_items = lines[1:] if len(lines) > 1 else []
        body_html = "".join("<li>%s</li>" % esc(x) for x in body_items)
        req_lines = split_lines(b)
        req_html = ""
        if req_lines:
            req_li = "".join("<li>%s</li>" % esc(x) for x in req_lines)
            req_html = '<div class="sp-req"><div class="sp-req-t">验收标准（B列）</div><ul>%s</ul></div>' % req_li
        imgs = images.get(row, [])
        imgs_html = []
        for j, rel in enumerate(imgs):
            badge = '<span class="badge key">重点参考图</span>' if j == 0 else '<span class="badge opt">参考图 %d</span>' % (j + 1)
            imgs_html.append(
                '<div class="thumb reveal">%s<img src="%s" loading="lazy" alt="%s"></div>'
                % (badge, esc(rel), esc(title))
            )
        search_txt = (title + " " + c + " " + b).replace("\n", " ")
        sp_parts.append(
            '<div class="sp reveal" data-search="%s">\n'
            '  <div class="sp-head">\n'
            '    <div class="sp-idx">%02d</div>\n'
            '    <div><h3>%s</h3>%s</div>\n'
            '  </div>\n'
            '  <div class="sp-body">%s%s</div>\n'
            '  <div class="sp-imgs">%s</div>\n'
            '</div>'
            % (esc(search_txt), i, esc(title),
               ('<div class="sp-sub">%s</div>' % esc(sub)) if sub else "",
               ('<ul>%s</ul>' % body_html) if body_html else "", req_html,
               "\n".join(imgs_html))
        )

    src = Path(html_file).read_text(encoding="utf-8")

    def replace_section(src, sec_id, new_inner):
        pat = re.compile(r'(<section id="%s"[^>]*>).*?(</section>)' % sec_id, re.S)
        m = pat.search(src)
        if not m:
            raise RuntimeError("section %s not found" % sec_id)
        return src[:m.start()] + m.group(1) + "\n" + new_inner + "\n" + m.group(2) + src[m.end():]

    src = replace_section(src, "category", "\n".join(cat_parts))
    src = replace_section(src, "sellpoint", "\n\n".join(sp_parts))
    src = src.replace(
        "从卖点表格第 20 行起、纵向至 F 列的全部卖点，完整保留文案与参考图。每行左侧第一张为",
        "按新版验收标准表格横向逐行制作：每一行即为一个完整的卖点框架，左至右依次为验收要求与对应参考图素材。每行第一张为"
    )
    Path(html_file).write_text(src, encoding="utf-8")
    print(f"[完成] 新版 Excel 重建：细分类别 {len(cat_parts)} 个，卖点框架 {len(sp_parts)} 个，内嵌图 {sum(len(v) for v in images.values())} 张")
    return len(sp_parts)


# ---------------------------------------------------------------- HTML 更新
def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def build_standard_html(st):
    kw = ""
    m = re.search(r"核心关键词[：:]\s*([^\n|]+)", st["body"])
    if m:
        kw = m.group(1).strip()
    lines = [ln for ln in split_lines(st["body"]) if not ln.startswith("核心关键词")]
    lis = "".join(f"<li>{esc(ln)}</li>" for ln in lines)
    return f'<div class="kw">核心关键词：{esc(kw)}</div><ul>{lis}</ul>'


def build_sp_html(sp):
    idx = f"{sp['idx']:02d}"
    title = esc(sp["title"])
    sub = esc(sp["sub"])
    items = "".join(f"<li>{esc(ln)}</li>" for ln in sp["items"])
    ds = esc((sp["title"] + " " + sp["sub"] + " " + " ".join(sp["items"])).strip())
    return (
        f'<div class="sp-idx">{idx}</div><div><h3>{title}</h3>'
        f'<div class="sp-sub">{sub}</div></div>',
        f"<div class=\"sp-body\"><ul>{items}</ul></div>",
        f'<div class="sp reveal" data-search="{ds}">',
    )


def update_html(src_xlsx, dry_run=False):
    if not Path(HTML_FILE).exists():
        print(f"[错误] 未找到 {HTML_FILE}，请在仓库根目录运行")
        sys.exit(1)
    data = parse_site(src_xlsx)
    html = Path(HTML_FILE).read_text(encoding="utf-8")

    # 1) 核心标准卡
    cards = re.findall(r'(<div class="card reveal"[^>]*>)(.*?)(</div>\s*</div>\s*</div>\s*</div>)', html, re.S)
    for i, st in enumerate(data["standards"][:3]):
        if i < len(cards):
            body = build_standard_html(st)
            old = cards[i][1]
            new_body = re.sub(r'<div class="kw">.*?</div><ul>.*?</ul>', body, old, flags=re.S)
            html = html.replace(old, new_body, 1)

    # 2) 细分类别
    cats_html = "".join(
        f'<div class="cat reveal" data-search="{esc(c["name"] + " " + c["desc"])}"><b>{esc(c["name"])}</b>'
        f'<span>{esc(c["desc"])}</span></div>\n'
        for c in data["categories"]
    )
    html = re.sub(r'<div class="cats">.*?</div>\s*</div>\s*</section>',
                  f'<div class="cats">\n{cats_html}\n    </div>\n  </div>\n</section>',
                  html, flags=re.S, count=1)

    # 3) 卖点卡（按序号替换文案，保留图）
    sps = re.findall(r'<div class="sp reveal"[^>]*>(.*?)(?=<div class="sp reveal"|</div>\s*</div>\s*</section>)', html, re.S)
    new_sps = []
    for i, sp in enumerate(data["sellpoints"]):
        head, body, opener = build_sp_html(sp)
        if i < len(sps):
            old = sps[i]
            old_head = re.search(r'<div class="sp-head">.*?</div>', old, re.S).group(0)
            old_body = re.search(r'<div class="sp-body">.*?</div>', old, re.S).group(0)
            new = old.replace(old_head, f'<div class="sp-head">{head}</div>', 1)
            new = new.replace(old_body, body, 1)
            new = re.sub(r'<div class="sp reveal"[^>]*>', opener, new, count=1)
            new_sps.append(new)
        else:
            new_sps.append(f'{opener}<div class="sp-head">{head}</div>{body}'
                           f'<div class="sp-imgs"><div class="thumb" style="flex:1 1 100%;min-width:0;padding:18px;cursor:default">'
                           f'<span class="badge opt" style="position:static;display:inline-block">待补充参考图</span></div></div></div>')

    sell_join = "\n\n\n".join(new_sps)
    html = re.sub(r'<div class="sell-grid">.*?</div>\s*</div>\s*</section>',
                  f'<div class="sell-grid">\n\n\n{sell_join}\n\n\n    </div>\n  </div>\n</section>',
                  html, flags=re.S, count=1)

    if dry_run:
        print("[dry-run] 解析结果预览:")
        print(json.dumps({k: (v if k == "categories" else v) for k, v in data.items()},
                         ensure_ascii=False, indent=2)[:3000])
        return
    Path(HTML_FILE).write_text(html, encoding="utf-8")
    print(f"[完成] 文案已更新：标准 {len(data['standards'])} 条 / 分类 {len(data['categories'])} 条 / 卖点 {len(data['sellpoints'])} 个")


# ---------------------------------------------------------------- 图片更新
def add_gallery_from_zip(zip_path):
    """解压图片并追加为新图库分组（按 顶层子目录 / 文件名公共前缀 分组）"""
    with zipfile.ZipFile(zip_path) as zf:
        names = [n for n in zf.namelist()
                 if n.lower().endswith((".jpg", ".jpeg", ".png", ".webp")) and not n.startswith("__MACOSX")]
    if not names:
        print("[跳过] zip 中没有图片")
        return 0
    os.makedirs(GALLERY_DIR, exist_ok=True)
    groups = {}
    for n in names:
        parts = Path(n).parts
        key = parts[0] if len(parts) > 1 else Path(n).stem.split("_")[0]
        groups.setdefault(key, []).append(n)
    added = 0
    for key, items in groups.items():
        thumb_html = []
        for n in items:
            dest = Path(GALLERY_DIR) / Path(n).name
            with zipfile.ZipFile(zip_path) as zf, zf.open(n) as src, open(dest, "wb") as out:
                shutil.copyfileobj(src, out)
            thumb_html.append(f'<div class="g-thumb reveal"><img src="{GALLERY_DIR}/{Path(n).name}" loading="lazy" alt="{Path(n).name}"></div>')
            added += 1
        grp = (f'\n\n    <div class="grp reveal" data-search="{esc(key)}">\n'
               f'      <h4>{esc(key)}</h4>\n      <div class="gallery-grid">\n'
               + "\n".join(thumb_html) + "\n      </div>\n    </div>\n\n")
        # 插入到图库 section 末尾（</section> 前的最后一个 </div> 之前）
        idx = html_section_gallery_end()
        html = Path(HTML_FILE).read_text(encoding="utf-8")
        html = html[:idx] + grp + html[idx:]
        Path(HTML_FILE).write_text(html, encoding="utf-8")
    print(f"[完成] zip 图库新增 {added} 张，分组 {len(groups)} 个")
    return added


def html_section_gallery_end():
    html = Path(HTML_FILE).read_text(encoding="utf-8")
    gal = html.rfind('id="gallery"')
    if gal == -1:
        return len(html)
    end = html.find("</section>", gal)
    # 找到该 section 内最后一个 "  </div>" 行的位置，插在其后
    seg = html[:end]
    last = seg.rfind("\n  </div>")
    return last + len("\n  </div>") if last != -1 else end


def add_gallery_from_pdf(pdf_path):
    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("[错误] 需要 PyMuPDF，请先 pip install pymupdf")
        sys.exit(1)
    doc = fitz.open(pdf_path)
    os.makedirs(GALLERY_DIR, exist_ok=True)
    total = 0
    for pno in range(len(doc)):
        imgs = doc.get_page_images(pno)
        if not imgs:
            continue
        thumbs = []
        for xref, *_ in imgs:
            try:
                pix = fitz.Pixmap(doc, xref)
                if pix.n > 4:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                name = f"upd_p{pno+1:03d}_x{xref}.jpg"
                dest = Path(GALLERY_DIR) / name
                if dest.exists():
                    thumbs.append(name)
                    continue
                pix.save(str(dest))
                thumbs.append(name)
                total += 1
            except Exception:
                continue
        if thumbs:
            grp = (f'\n\n    <div class="grp reveal" data-search="更新页 {pno+1}">\n'
                   f'      <h4>更新页 {pno+1}</h4>\n      <div class="gallery-grid">\n'
                   + "\n".join(f'<div class="g-thumb reveal"><img src="{GALLERY_DIR}/{name}" loading="lazy" alt="{name}"></div>' for name in thumbs)
                   + "\n      </div>\n    </div>\n\n")
            idx = html_section_gallery_end()
            html = Path(HTML_FILE).read_text(encoding="utf-8")
            html = html[:idx] + grp + html[idx:]
            Path(HTML_FILE).write_text(html, encoding="utf-8")
    print(f"[完成] PDF 提取新增 {total} 张内嵌图")
    return total


# ---------------------------------------------------------------- 入口
def main():
    ap = argparse.ArgumentParser(description="更新卖点图验收标准参考图库网站")
    ap.add_argument("--xlsx", help="卖点表格 .xlsx 路径")
    ap.add_argument("--zip", help="参考图压缩包 .zip 路径")
    ap.add_argument("--pdf", help="验收标准 .pdf 路径")
    ap.add_argument("--source", help="source/ 目录：自动取其中的 xlsx/zip/pdf")
    ap.add_argument("--all", action="store_true", help="结合 --source 自动处理全部")
    ap.add_argument("--dry-run", action="store_true", help="仅解析预览，不写文件")
    args = ap.parse_args()

    src = Path(args.source) if args.source else None
    xlsx = args.xlsx
    if src:
        if xlsx is None:
            cands = list(src.glob("*.xlsx")) + list(src.glob("*.xls"))
            if cands:
                xlsx = str(cands[0])
        if args.all:
            if args.zip is None:
                zc = list(src.glob("*.zip"))
                if zc:
                    args.zip = str(zc[0])
            if args.pdf is None:
                pc = list(src.glob("*.pdf"))
                if pc:
                    args.pdf = str(pc[0])

    if xlsx:
        # 新版 Excel（含内嵌卖点图）走整体重建，兼容旧版仅文案更新
        if xlsx_drawing_row_map(xlsx):
            rebuild_sections_from_xlsx(xlsx, html_file=HTML_FILE) if not args.dry_run else print("[dry-run] 新版 Excel 重建（含内嵌图）")
        else:
            update_html(xlsx, dry_run=args.dry_run)
    if not args.dry_run:
        if args.zip:
            add_gallery_from_zip(args.zip)
        if args.pdf:
            add_gallery_from_pdf(args.pdf)
    if not xlsx and not args.zip and not args.pdf:
        ap.print_help()


if __name__ == "__main__":
    main()
