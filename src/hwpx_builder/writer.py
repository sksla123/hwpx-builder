"""The only module that maps normalized model objects to python-hwpx/XML."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import replace
from pathlib import Path

from hwpx import HwpxDocument
from lxml import etree as ET

from .model import Composition, Document, PageBreak, Paragraph, ParagraphStyle, Picture, Table
from .units import mm_to_hwpunit as unit

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
HC = "{http://www.hancom.co.kr/hwpml/2011/core}"


class Writer:
    def __init__(self, document: Document, debug: bool = False):
        self.model = document
        self.debug = debug
        self.doc = HwpxDocument.new()
        self.header = self.doc.oxml.headers[0]
        self.initial = self.doc.sections[0].paragraphs[0]
        self.pending_break = False
        self.styles = {}
        self.paragraph_styles = {}
        self.fills = {}
        self.fonts = set()
        self.images = {}

    def font(self, name):
        if name in self.fonts:
            return
        for face in self.header.element.findall(".//" + HH + "fontface"):
            if any(f.get("face") == name for f in face):
                continue
            next_id = max((int(f.get("id", "0")) for f in face), default=-1) + 1
            ET.SubElement(face, HH + "font", id=str(next_id), face=name, type="TTF", isEmbedded="0")
            face.set("fontCnt", str(len(face)))
        self.fonts.add(name)
        self.header.mark_dirty()

    def char_style(self, style):
        if style not in self.styles:
            self.font(style.font_family)
            self.styles[style] = self.doc.oxml.ensure_run_style(
                font=style.font_family,
                size=float(style.font_size),
                bold=style.bold,
                italic=style.italic,
                underline=style.underline,
                strike=style.strike,
                color=style.color,
                highlight=style.highlight or "none",
                base_char_pr_id="0",
            )
        return self.styles[style]

    def para_style(self, style):
        if style in self.paragraph_styles:
            return self.paragraph_styles[style]
        result = self.header.ensure_paragraph_format(
            alignment=style.align.upper(),
            line_spacing_percent=style.line_spacing,
            margins={
                "left": unit(style.indent),
                "right": unit(style.right_indent),
                "prev": unit(style.space_before),
                "next": unit(style.space_after),
                "intent": unit(style.first_line_indent),
            },
            break_setting={"keepWithNext": style.keep_with_next, "keepLines": style.keep_together},
        )
        self.paragraph_styles[style] = result
        return result

    def runs(self, target, paragraph):
        target.element.set("paraPrIDRef", self.para_style(paragraph.style))
        for run in paragraph.runs:
            style = (
                replace(run.style, highlight="#FFFF00")
                if self.debug and run.reference is not None
                else run.style
            )
            created = target.add_run(run.text, char_pr_id_ref=self.char_style(style))
            if run.linebreak:
                t = created.element.find(HP + "t")
                t.text = None
                ET.SubElement(t, HP + "lineBreak")
        if not paragraph.runs:
            target.add_run("")

    def host(self, paragraph):
        """Move initial section controls to the first real host; avoid a leading blank line."""
        if self.initial is not None:
            if paragraph.element is not self.initial.element:
                controls = [
                    r
                    for r in self.initial.element
                    if r.tag == HP + "run" and any(e.tag in (HP + "secPr", HP + "ctrl") for e in r)
                ]
                for i, control in enumerate(controls):
                    paragraph.element.insert(i, control)
                self.doc.sections[0].element.remove(self.initial.element)
            else:
                for child in list(paragraph.element):
                    if child.tag == HP + "linesegarray" or (
                        child.tag == HP + "run" and all(c.tag == HP + "t" for c in child)
                    ):
                        paragraph.element.remove(child)
            self.initial = None
        if self.pending_break:
            paragraph.element.set("pageBreak", "1")
            self.pending_break = False
        paragraph.section.mark_dirty()

    def paragraph(self, model):
        target = (
            self.initial
            if self.initial is not None
            else self.doc.add_paragraph("", include_run=False, inherit_style=False)
        )
        self.host(target)
        self.runs(target, model)

    def picture(self, model):
        pic = self.doc.add_picture(
            model.data,
            model.format,
            width=unit(model.width),
            height=unit(model.height),
            align=model.align.upper(),
        )
        self.host(pic.paragraph)
        return pic

    def binary(self, model):
        key = (model.data, model.format)
        if key not in self.images:
            # Reuse the public package/image registration API, then remove the temporary host.
            pic = self.doc.add_picture(
                model.data, model.format, width=unit(model.width), height=unit(model.height)
            )
            image = pic.element.find(HC + "img")
            self.images[key] = image.get("binaryItemIDRef")
            pic.paragraph.element.getparent().remove(pic.paragraph.element)
            self.doc.sections[0].mark_dirty()
        return self.images[key]

    def border_fill(self, cell):
        image_id = self.binary(cell.picture) if cell.picture_as_background else None
        key = (cell.borders, cell.background, image_id)
        if key in self.fills:
            return self.fills[key]
        fills = self.header.element.find(".//" + HH + "borderFills")
        base = deepcopy(fills[0])
        fid = str(max(int(n.get("id")) for n in fills) + 1)
        base.set("id", fid)
        for tag, border in zip(("left", "right", "top", "bottom"), cell.borders):
            n = base.find(HH + tag + "Border")
            if n is None:
                n = ET.SubElement(base, HH + tag + "Border")
            n.attrib.update({"type": border.style, "width": border.width, "color": border.color})
        diagonal = base.find(HH + "diagonal")
        if diagonal is not None:
            diagonal.set("type", "NONE")
        for fill in list(base.findall(HC + "fillBrush")):
            base.remove(fill)
        if image_id or cell.background:
            brush = ET.SubElement(base, HC + "fillBrush")
            if image_id:
                imgbrush = ET.SubElement(brush, HC + "imgBrush", mode="TOTAL")
                ET.SubElement(
                    imgbrush,
                    HC + "img",
                    binaryItemIDRef=image_id,
                    bright="0",
                    contrast="0",
                    effect="REAL_PIC",
                    alpha="0",
                )
            else:
                ET.SubElement(
                    brush, HC + "winBrush", faceColor=cell.background, hatchColor="#FF000000", alpha="0"
                )
        fills.append(base)
        fills.set("itemCnt", str(len(fills)))
        self.header.mark_dirty()
        self.fills[key] = fid
        return fid

    def table(self, model):
        data = model.data
        widths = [unit(w) for w in data.widths]
        heights = [unit(h) for h in data.heights]
        table = self.doc.add_table(len(heights), len(widths), width=sum(widths), height=sum(heights))
        self.host(table.paragraph)
        table.paragraph.element.set("paraPrIDRef", self.para_style(data.host_style))
        table.element.set("pageBreak", data.page_break)
        table.element.set("repeatHeader", str(int(data.repeat_header)))
        if data.size is not None:
            table.element.find(HP + "sz").attrib.update(
                {"width": str(unit(data.size[0])), "height": str(unit(data.size[1]))}
            )
        pos = table.element.find(HP + "pos")
        pos.set("treatAsChar", str(int(data.inline)))
        pos.set("horzAlign", data.host_style.align.upper())
        for tag, values in (("inMargin", data.padding), ("outMargin", data.margin)):
            element = table.element.find(HP + tag)
            for name, value in zip(("left", "right", "top", "bottom"), values):
                element.set(name, str(unit(value)))
        for r, h in enumerate(heights):
            for c, w in enumerate(widths):
                size = table.cell(r, c).element.find(HP + "cellSz")
                size.set("width", str(w))
                size.set("height", str(h))
        for cell in data.cells:
            if cell.row_span > 1 or cell.col_span > 1:
                table.merge_cells(
                    cell.row, cell.col, cell.row + cell.row_span - 1, cell.col + cell.col_span - 1
                )
        for cell in data.cells:
            target = table.cell(cell.row, cell.col)
            size = target.element.find(HP + "cellSz")
            size.set("width", str(sum(widths[cell.col : cell.col + cell.col_span])))
            size.set("height", str(sum(heights[cell.row : cell.row + cell.row_span])))
            if cell.size is not None:
                size.set("width", str(unit(cell.size[0])))
                size.set("height", str(unit(cell.size[1])))
            target.element.set("borderFillIDRef", self.border_fill(cell))
            target.element.set("hasMargin", "1")
            target.element.set("header", str(int(cell.header)))
            margin = target.element.find(HP + "cellMargin")
            for name, value in zip(("left", "right", "top", "bottom"), cell.padding):
                margin.set(name, str(unit(value)))
            sub = target.element.find(HP + "subList")
            sub.set("vertAlign", cell.vertical_align.upper())
            for p in list(sub.findall(HP + "p")):
                sub.remove(p)
            if cell.picture is not None and not cell.picture_as_background:
                pic = self.doc.add_picture(
                    cell.picture.data,
                    cell.picture.format,
                    width=unit(cell.picture.width),
                    height=unit(cell.picture.height),
                    align="CENTER",
                )
                sub.append(pic.paragraph.element)
            if cell.picture is None or any(r.text for r in cell.paragraph.runs) or cell.picture_as_background:
                for paragraph in cell.paragraphs or (cell.paragraph,):
                    p = target.add_paragraph("")
                    for r in list(p.element.findall(HP + "run")):
                        p.element.remove(r)
                    self.runs(p, paragraph)
        table.mark_dirty()

    def composition(self, model):
        """An inline native HWPX group; text remains editable over bitmap artwork."""
        group = self.doc.add_rectangle(unit(model.width), unit(model.height))
        self.host(group.paragraph)
        group.paragraph.element.set("paraPrIDRef", self.para_style(ParagraphStyle(align="justify")))
        root = group.element
        root.tag = HP + "container"
        root.attrib.pop("ratio", None)
        root.set("numberingType", "PICTURE")
        for child in list(root):
            if child.tag in (HP + "lineShape", HP + "shadow") or child.tag.startswith(HP + "pt"):
                root.remove(child)
        insertion = list(root).index(root.find(HP + "sz"))
        for index, child in enumerate(model.contents):
            if isinstance(child, Picture):
                obj = self.doc.add_picture(
                    child.data, child.format, width=unit(child.width), height=unit(child.height)
                )
                x = y = 0
            else:
                obj = self.doc.add_rectangle(unit(child.width), unit(child.height))
                x, y = unit(child.x), unit(child.y)
                obj.element.find(HP + "lineShape").set("style", "NONE")
                draw = ET.Element(HP + "drawText", name="", editable="0")
                sub = ET.SubElement(
                    draw,
                    HP + "subList",
                    id="",
                    textDirection="HORIZONTAL",
                    lineWrap="BREAK",
                    vertAlign="CENTER",
                    linkListIDRef="0",
                    linkListNextIDRef="0",
                    textWidth="0",
                    textHeight="0",
                    hasTextRef="0",
                    hasNumRef="0",
                )
                for paragraph in child.paragraphs:
                    p = self.doc.add_paragraph("", include_run=False, inherit_style=False)
                    self.runs(p, paragraph)
                    sub.append(p.element)
                ET.SubElement(
                    draw,
                    HP + "textMargin",
                    **{k: str(unit(v)) for k, v in zip(("left", "right", "top", "bottom"), child.padding)},
                )
                obj.element.insert(list(obj.element).index(obj.element.find(HP + "pt0")), draw)
                for point in list(obj.element):
                    if point.tag.startswith(HP + "pt"):
                        point.tag = HC + point.tag[len(HP) :]
            element = obj.element
            obj.paragraph.element.getparent().remove(obj.paragraph.element)
            element.set("groupLevel", "1")
            element.set("zOrder", str(index))
            element.find(HP + "offset").attrib.update({"x": str(x), "y": str(y)})
            transform = element.find(HP + "renderingInfo/" + HC + "transMatrix")
            transform.set("e3", str(x))
            transform.set("e6", str(y))
            for tag in ("sz", "pos", "outMargin"):
                for node in element.findall(HP + tag):
                    element.remove(node)
            root.insert(insertion + index, element)
        group.paragraph.section.mark_dirty()

    def save(self, path: Path):
        page = self.model.sections[0].page
        section = self.doc.sections[0]
        pagepr = section.element.find(".//" + HP + "pagePr")
        pagepr.set("width", str(unit(page.width)))
        pagepr.set("height", str(unit(page.height)))
        pagepr.set("landscape", "WIDELY")
        pagepr.set(
            "gutterType",
            "LEFT_ONLY",
        )
        margin = pagepr.find(HP + "margin")
        for name in ("left", "right", "top", "bottom", "header", "footer", "gutter"):
            margin.set(name, str(unit(getattr(page, name))))
        section.mark_dirty()
        for logical in self.model.sections:
            for block in logical.contents:
                if isinstance(block, PageBreak):
                    self.pending_break = True
                    self.paragraph(Paragraph(style=ParagraphStyle(line_spacing=1)))
                elif isinstance(block, Paragraph):
                    self.paragraph(block)
                elif isinstance(block, Picture):
                    self.picture(block)
                elif isinstance(block, Table):
                    self.table(block)
                elif isinstance(block, Composition):
                    self.composition(block)
        if self.pending_break:
            self.paragraph(Paragraph())
        self.doc.save_to_path(path)
