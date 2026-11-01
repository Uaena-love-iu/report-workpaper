#!/usr/bin/env python3
"""
build_workpaper.py — Assemble the 底稿 (work paper) from a manifest.

Layout per subsection sheet (mirrors the analyst's existing 底稿):
  LEFT block   = report-paragraph screenshots, stacked top-to-bottom.
  RIGHT block  = source screenshots (red-boxed), beside the paragraph they
                 support; multiple sources sit side-by-side.
  Each source gets a clickable TITLE cell (hyperlink to the URL) above its
  screenshot. One row of claim ←→ one row of evidence.
  source type "sheet" -> a cross-sheet link to a table sheet (no image).

Emits an .xlsx and (optionally) a PNG preview per sheet from the SAME layout
pass, so you can eyeball the result without opening Excel.

Manifest JSON — see SKILL.md. Usage:
  python build_workpaper.py manifest.json --out 底稿.xlsx [--preview-dir prev/]
"""
import argparse, json, os
from PIL import Image as PILImage, ImageDraw, ImageFont
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side

# ---- pixel grid ----
# NOTE: images are anchored to REAL cells (twoCellAnchor), like the analyst's
# original 底稿. Anchoring to A1 + a big offset gets clamped to the top-left by
# Excel/Numbers on macOS (all images pile up) — so we pin every image to the
# grid cell its pixel position falls in. A normal (not tiny) grid keeps the sheet
# looking like a regular spreadsheet (gridlines on); sub-cell EMU offsets in the
# twoCellAnchor carry the exact pixel position, so the grid size doesn't affect
# image alignment. COL_PX/ROW_PX map back to exact pixel sizes: width=(COL_PX-5)/7,
# height_pt = ROW_PX*0.75.
ROW_PX = 20       # normal row height (~15pt)
COL_PX = 64       # normal column width (~8.43 chars)
MARGIN_X = 16
MARGIN_Y = 14
LW = 640          # left (report) display width
GAP = 96          # gap between left block and right block (holds the divider)
SW = 620          # source display width (wider = bigger, more readable source text)
SGAP = 30
TITLE_H = 44      # title band height above each source
ROW_GAP = 60      # vertical gap between paragraph rows
HEADER_H = 60     # sheet title band

BLUE = Font(color="0563C1", underline="single", size=10, name="PingFang SC")
TITLE_FONT = Font(bold=True, size=14, color="C00000", name="PingFang SC")
SHEET_LINK = Font(color="7030A0", underline="single", size=10, bold=True,
                  name="PingFang SC")
WRAP = Alignment(wrap_text=True, vertical="top")
# colored separator line between paragraph bands (matches the report's red)
SEP_SIDE = Side(style="medium", color="C00000")


def px_to_col_width(px):
    # inverse of Excel's pixel = round(width*7)+5 for the default font
    return max((px - 5) / 7.0, 0.4)


def img_disp_size(path, target_w):
    with PILImage.open(path) as im:
        w, h = im.size
    return target_w, int(round(h * target_w / w))


def layout_sheet(sheet):
    """Return (placements, width_px, height_px). placements are pixel-positioned
    dicts shared by the xlsx writer and the preview renderer."""
    pl = []
    y = MARGIN_Y
    pl.append({"kind": "header", "x": MARGIN_X, "y": y,
               "text": f"{sheet['label']}  {sheet['title']}"})
    y += HEADER_H
    right_x = MARGIN_X + LW + GAP
    max_x = right_x
    for row in sheet["rows"]:
        row_top = y
        left_h = 0
        li = row.get("left_img")
        if li and os.path.exists(li):
            w, h = img_disp_size(li, LW)
            pl.append({"kind": "image", "x": MARGIN_X, "y": row_top,
                       "w": w, "h": h, "path": li})
            left_h = h
        src_block_h = 0
        sx = right_x
        for src in row.get("sources", []):
            pl.append({"kind": "title", "x": sx, "y": row_top, "src": src})
            simg = src.get("img")
            if src.get("type") == "sheet" or not simg or not os.path.exists(simg):
                src_block_h = max(src_block_h, TITLE_H + 28)
                sx += int(SW * 0.62) + SGAP
                continue
            w, h = img_disp_size(simg, SW)
            pl.append({"kind": "image", "x": sx, "y": row_top + TITLE_H,
                       "w": w, "h": h, "path": simg})
            src_block_h = max(src_block_h, TITLE_H + h)
            sx += SW + SGAP
        max_x = max(max_x, sx)
        row_h = max(left_h, src_block_h) or 60
        y += row_h + ROW_GAP
        pl.append({"kind": "sep", "y": int(y - ROW_GAP / 2)})
    return pl, int(max_x + MARGIN_X), int(y + MARGIN_Y)


# ---------- xlsx ----------
def cell_marker(x_px, y_px):
    """Pixel -> AnchorMarker pinned to the real grid cell + sub-cell offset."""
    c = int(x_px // COL_PX)
    r = int(y_px // ROW_PX)
    return AnchorMarker(col=c, colOff=pixels_to_EMU(x_px - c * COL_PX),
                        row=r, rowOff=pixels_to_EMU(y_px - r * ROW_PX))


def cell_rc(x_px, y_px):
    return int(y_px // ROW_PX) + 1, int(x_px // COL_PX) + 1


def _add_side(cell, left=None, bottom=None):
    """Add a border side without clobbering the others already set."""
    b = cell.border
    cell.border = Border(left=left if left is not None else b.left,
                         right=b.right, top=b.top,
                         bottom=bottom if bottom is not None else b.bottom)


def write_xlsx_sheet(ws, sheet):
    pl, w_px, h_px = layout_sheet(sheet)
    total_cols = int(w_px // COL_PX) + 4
    for p in pl:
        if p["kind"] == "header":
            r, c = cell_rc(p["x"], p["y"])
            cell = ws.cell(row=r, column=c, value=p["text"])
            cell.font = TITLE_FONT
            cell.alignment = Alignment(vertical="center")
            # give the title room (merge end_* are INCLUSIVE -> span-1)
            try:
                ws.merge_cells(start_row=r, start_column=c, end_row=r + 2,
                               end_column=c + max(1, int(LW // COL_PX)) - 1)
            except Exception as e:
                print(f"[warn] header merge failed @r{r}c{c}: {e}")
        elif p["kind"] == "image":
            # anchor to REAL cells (twoCellAnchor) like the original 底稿,
            # so macOS Excel/Numbers place it correctly instead of clamping.
            img = XLImage(p["path"])
            img.anchor = TwoCellAnchor(
                editAs="oneCell",
                _from=cell_marker(p["x"], p["y"]),
                to=cell_marker(p["x"] + p["w"], p["y"] + p["h"]))
            ws.add_image(img)
        elif p["kind"] == "title":
            src = p["src"]
            r, c = cell_rc(p["x"], p["y"])
            cell = ws.cell(row=r, column=c)
            if src.get("type") == "sheet":
                cell.value = src.get("title") or f"见表: {src['target']}"
                cell.hyperlink = f"#'{src['target']}'!A1"
                cell.font = SHEET_LINK
            else:
                cell.value = src.get("title") or src.get("url")
                if src.get("url"):
                    cell.hyperlink = src["url"]
                cell.font = BLUE
            cell.alignment = WRAP
            span = max(1, int(SW // COL_PX))   # merge end_* INCLUSIVE -> span-1
            try:
                ws.merge_cells(start_row=r, start_column=c, end_row=r + 1,
                               end_column=c + span - 1)
            except Exception as e:
                print(f"[warn] title merge failed @r{r}c{c}: {e}")
        elif p["kind"] == "sep":
            r = int(p["y"] // ROW_PX) + 1
            for c in range(1, total_cols + 1):
                _add_side(ws.cell(row=r, column=c), bottom=SEP_SIDE)
    total_rows = int(h_px // ROW_PX) + 4
    # vertical red divider between the report column and the source column
    sep_col = int((MARGIN_X + LW + GAP // 2) // COL_PX) + 1
    for r in range(1, total_rows + 1):
        _add_side(ws.cell(row=r, column=sep_col), left=SEP_SIDE)
    for c in range(1, total_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = px_to_col_width(COL_PX)
    for r in range(1, total_rows + 1):
        ws.row_dimensions[r].height = ROW_PX * 0.75
    ws.sheet_view.showGridLines = True


# ---------- preview PNG ----------
def _font(size, bold=False):
    cands = ["/System/Library/Fonts/PingFang.ttc",
             "/System/Library/Fonts/STHeiti Medium.ttc",
             "/System/Library/Fonts/Hiragino Sans GB.ttc"]
    for c in cands:
        if os.path.exists(c):
            try:
                return ImageFont.truetype(c, size)
            except Exception:
                pass
    return ImageFont.load_default()


def render_preview(sheet, out_png):
    pl, w_px, h_px = layout_sheet(sheet)
    scale = 1.0
    canvas = PILImage.new("RGB", (int(w_px * scale), int(h_px * scale)), "white")
    draw = ImageDraw.Draw(canvas)
    for p in pl:
        if p["kind"] == "sep":
            draw.line([(0, p["y"]), (w_px, p["y"])], fill="#C00000", width=2)
        elif p["kind"] == "image":
            im = PILImage.open(p["path"]).convert("RGB")
            im = im.resize((int(p["w"] * scale), int(p["h"] * scale)))
            canvas.paste(im, (int(p["x"] * scale), int(p["y"] * scale)))
            draw.rectangle([p["x"], p["y"], p["x"] + p["w"], p["y"] + p["h"]],
                           outline="#cfcfcf", width=1)
        elif p["kind"] == "header":
            draw.text((p["x"], p["y"]), p["text"], fill="#C00000",
                      font=_font(20, True))
        elif p["kind"] == "title":
            src = p["src"]
            txt = src.get("title") or src.get("url") or src.get("target", "")
            color = "#7030A0" if src.get("type") == "sheet" else "#0563C1"
            # wrap title to source width
            f = _font(13)
            words = txt
            line = ""
            yy = p["y"]
            maxw = SW
            for ch in words:
                if draw.textlength(line + ch, font=f) > maxw:
                    draw.text((p["x"], yy), line, fill=color, font=f)
                    yy += 17
                    line = ch
                else:
                    line += ch
            draw.text((p["x"], yy), line, fill=color, font=f)
    # vertical red divider between report column and source column
    sep_x = MARGIN_X + LW + GAP // 2
    draw.line([(sep_x, MARGIN_Y), (sep_x, h_px)], fill="#C00000", width=2)
    canvas.save(out_png)
    return out_png


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("manifest")
    ap.add_argument("--out", required=True)
    ap.add_argument("--preview-dir")
    a = ap.parse_args()
    man = json.load(open(a.manifest))
    wb = Workbook()
    wb.remove(wb.active)
    previews = []
    for sheet in man["sheets"]:
        ws = wb.create_sheet(title=sheet["label"][:31])
        write_xlsx_sheet(ws, sheet)
        if a.preview_dir:
            os.makedirs(a.preview_dir, exist_ok=True)
            pv = os.path.join(a.preview_dir, f"preview_{sheet['label']}.png")
            render_preview(sheet, pv)
            previews.append(pv)
    wb.save(a.out)
    nsh = len(man["sheets"])
    nrows = sum(len(s["rows"]) for s in man["sheets"])
    print(f"写出 {a.out}: {nsh} 个 sheet, {nrows} 段")
    for pv in previews:
        print("  预览:", pv)


if __name__ == "__main__":
    main()
