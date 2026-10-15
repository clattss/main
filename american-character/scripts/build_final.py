"""American Character deck: hook, one character per group member, collective conclusion.

Layout rules (agreed with the user):
- every slide states one main idea as big type (3-6 words), sits on a regular grid
- the four character slides share one identical layout: claim on top, two contrasting
  stills side by side with two-word labels and an arrow between them
- images must carry the logic (a before/after pair), not just decorate
- on-screen English stays under ~15 words per slide; the Chinese explanation, the
  "would it work in China" prompt and the spoken line live in the speaker notes
"""

from __future__ import annotations

import copy
from pathlib import Path

from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageOps
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(r"D:\AI-agent\codexworksection\american-character")
SOURCE = ROOT / "assets" / "source"
TREATED = ROOT / "assets" / "treated"
BUILD = ROOT / "build"
OUT_FILE = BUILD / "American-Character-Forrest-Gump.pptx"

INK = RGBColor(0x0E, 0x11, 0x16)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT = RGBColor(0xFF, 0xC5, 0x3D)
MUTED = RGBColor(0xA8, 0xAF, 0xB8)
FONT = "Arial"

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
SLIDE_W, SLIDE_H = 13.3333, 7.5
MARGIN = 0.9


# ------------------------------------------------------------------ helpers


def treat(name: str) -> Path:
    """Unify the stills: warm duotone, contrast lift, soft vignette."""
    target = TREATED / f"{Path(name).stem}.png"
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        return target
    with Image.open(SOURCE / name) as im:
        src = im.convert("RGB")
    grey = ImageEnhance.Contrast(ImageOps.autocontrast(ImageOps.grayscale(src), cutoff=1)).enhance(1.25)
    dark, light = (12, 15, 22), (246, 239, 226)
    lut: list[int] = []
    for ch in range(3):
        lut.extend(int(dark[ch] + (light[ch] - dark[ch]) * (v / 255)) for v in range(256))
    out = Image.blend(grey.convert("RGB").point(lut), src, 0.32)
    out = ImageEnhance.Color(out).enhance(1.12)

    w, h = out.size
    mask = Image.new("L", (w, h), 0)
    draw = ImageDraw.Draw(mask)
    for i in range(60):
        ix, iy = int(w * 0.5 * (i / 60) ** 1.6), int(h * 0.5 * (i / 60) ** 1.6)
        draw.rectangle([ix, iy, w - ix, h - iy], fill=int(255 * (i / 60) ** 0.55))
    mask = mask.filter(ImageFilter.GaussianBlur(w / 22))
    out = Image.composite(out, Image.new("RGB", (w, h), (6, 8, 12)), mask)
    out.save(target)
    return target


def cover(slide, image_name, box, keep=(0.0, 0.0, 1.0, 1.0), anchor=(0.5, 0.5)):
    path = treat(image_name)
    with Image.open(path) as im:
        src_w, src_h = im.size
    left, top, width, height = box
    kl, kt, kr, kb = keep
    vis_w, vis_h = (kr - kl) * src_w, (kb - kt) * src_h
    target_ar = width / height
    ax, ay = anchor
    if vis_w / vis_h > target_ar:
        new_w = vis_h * target_ar
        x0 = kl * src_w + (vis_w - new_w) * ax
        x1 = x0 + new_w
        y0, y1 = kt * src_h, kb * src_h
    else:
        new_h = vis_w / target_ar
        y0 = kt * src_h + (vis_h - new_h) * ay
        y1 = y0 + new_h
        x0, x1 = kl * src_w, kr * src_w
    pic = slide.shapes.add_picture(str(path), Inches(left), Inches(top), Inches(width), Inches(height))
    pic.crop_left, pic.crop_right = x0 / src_w, 1 - x1 / src_w
    pic.crop_top, pic.crop_bottom = y0 / src_h, 1 - y1 / src_h
    return pic


def text(slide, content, box, size, *, color=WHITE, bold=True, align=PP_ALIGN.LEFT,
         spacing=1.0, tracking=None, anchor=MSO_ANCHOR.TOP, name=None):
    left, top, width, height = box
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    if name:
        shape.name = name
    frame = shape.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    for i, line in enumerate(content.split("\n")):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.alignment = align
        para.line_spacing = spacing
        run = para.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = FONT
        run.font.color.rgb = color
        if tracking:
            run.font._rPr.set("spc", str(int(tracking * 100)))
    return shape


def background(slide, color=INK):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def scrim(slide, direction="bottom", box=(0, 0, SLIDE_W, SLIDE_H), strength=235):
    path = ROOT / "assets" / "scrim" / f"{direction}-{strength}.png"
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        w, h = 1920, 1080
        img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        if direction == "bottom":
            for i in range(h):
                a = int(strength * max(0.0, ((i / h) + 0.1) ** 2.6))
                draw.line([(0, i), (w, i)], fill=(6, 8, 12, min(a, 255)))
        else:  # left
            for i in range(w):
                a = int(strength * max(0.0, 1 - (i / (w * 0.78)) ** 1.1))
                draw.line([(i, 0), (i, h)], fill=(6, 8, 12, min(a, 255)))
        img.save(path)
    left, top, width, height = box
    return slide.shapes.add_picture(str(path), Inches(left), Inches(top), Inches(width), Inches(height))


def credit(slide):
    return text(slide, "Stills: Forrest Gump (1994), Paramount Pictures", (8.2, 7.12, 4.1, 0.26),
                10, color=MUTED, bold=False, align=PP_ALIGN.RIGHT, name="credit")


def notes(slide, body: str):
    if isinstance(body, (list, tuple)):  # tolerate a trailing comma in the call sites
        body = "".join(str(part) for part in body)
    slide.notes_slide.notes_text_frame.text = body


# ------------------------------------------------------------ animation XML


def _fade_pair(shape_id: int, node_id: int):
    set_el = etree.fromstring(
        f'<p:set xmlns:p="{NS_P}" xmlns:a="{NS_A}"><p:cBhvr>'
        f'<p:cTn id="{node_id}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        f'<p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl>'
        f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
        f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
    )
    fade_el = etree.fromstring(
        f'<p:animEffect xmlns:p="{NS_P}" xmlns:a="{NS_A}" transition="in" filter="fade">'
        f'<p:cBhvr><p:cTn id="{node_id + 1}" dur="400"/>'
        f'<p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl></p:cBhvr></p:animEffect>'
    )
    return [set_el, fade_el]


def add_steps(slide, groups):
    timing = etree.SubElement(etree.Element("dummy"), f"{{{NS_P}}}timing")
    tn_lst = etree.SubElement(timing, f"{{{NS_P}}}tnLst")
    root_ctn = etree.SubElement(etree.SubElement(tn_lst, f"{{{NS_P}}}par"), f"{{{NS_P}}}cTn",
                                id="1", dur="indefinite", restart="never", nodeType="tmRoot")
    seq = etree.SubElement(etree.SubElement(root_ctn, f"{{{NS_P}}}childTnLst"),
                           f"{{{NS_P}}}seq", concurrent="1", nextAc="seek")
    seq_ctn = etree.SubElement(seq, f"{{{NS_P}}}cTn", id="2", dur="indefinite", nodeType="mainSeq")
    children = etree.SubElement(seq_ctn, f"{{{NS_P}}}childTnLst")

    def flatten(items):
        flat = []
        for item in items:
            if isinstance(item, (list, tuple)):
                flat.extend(flatten(item))
            else:
                flat.append(item)
        return flat

    node_id = 3
    for group in groups:
        par = etree.SubElement(children, f"{{{NS_P}}}par")
        ctn = etree.SubElement(par, f"{{{NS_P}}}cTn", id=str(node_id), fill="hold", nodeType="clickEffect")
        node_id += 1
        etree.SubElement(etree.SubElement(ctn, f"{{{NS_P}}}stCondLst"), f"{{{NS_P}}}cond", delay="0")
        grand = etree.SubElement(ctn, f"{{{NS_P}}}childTnLst")
        for shape in flatten(group):
            for effect in _fade_pair(shape.shape_id, node_id):
                grand.append(copy.deepcopy(effect))
                node_id += 2

    for evt, tag in (("onPrev", "prevCondLst"), ("onNext", "nextCondLst")):
        cond = etree.SubElement(etree.SubElement(seq, f"{{{NS_P}}}{tag}"), f"{{{NS_P}}}cond",
                                evt=evt, delay="0")
        etree.SubElement(etree.SubElement(cond, f"{{{NS_P}}}tgtEl"), f"{{{NS_P}}}sldTgt")
    slide._element.append(timing)


def kicker(slide, label):
    return text(slide, label, (MARGIN, 0.6, 8.0, 0.45), 20, color=ACCENT, bold=True,
                tracking=1.6, name="kicker")


# ------------------------------------------------------------------- slides


def slide_1(prs):
    """Hook: the audience's own life first, then the map of the whole deck."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background(slide)
    question = text(slide, "WHAT DO THEY ALL\nHAVE IN COMMON?", (MARGIN, 1.0, 11.6, 2.4), 60,
                    spacing=0.98, name="question")
    china = text(slide, "And would it work in China?", (MARGIN, 3.5, 11.6, 0.7), 28,
                 color=ACCENT, bold=False, name="china")

    faces = [("image6.png", "FORREST"), ("image2.png", "JENNY"),
             ("image12.jpeg", "HIS MOTHER"), ("image18.jpeg", "LIEUTENANT DAN")]
    row = []
    for i, (img, label) in enumerate(faces):
        x = MARGIN + i * 3.01
        row.append(cover(slide, img, (x, 4.62, 2.76, 1.72), keep=(0.04, 0.04, 0.96, 0.96), anchor=(0.5, 0.4)))
        row.append(text(slide, label, (x, 6.5, 2.76, 0.4), 20, color=WHITE, bold=True,
                        tracking=0.6, name=f"who{i}"))
    src = credit(slide)
    add_steps(slide, [[row, src], [question], [china]])
    notes(
        slide,
        "开场（由主管理人开场，或本页负责人）：\n"
        "1) 先问观众问题、把他们的生活拉进来：这四个人有什么共同点？（先别给答案）\n"
        "2) 再问一句俏皮的：如果这些性格长在你自己身上、或者发生在中国，会怎样？会改变你的选择吗？\n"
        "3) 点题：American character 不是阿甘一个人的性格，是这四个人的集合——这是本场的主张。\n"
        "口播：Before we talk about America, let's talk about you. Would these traits survive here?",
    )


def character_slide(prs, *, kicker_label, claim_lines, img_a, label_a, img_b, label_b, note):
    """One layout reused for all four group members, so the four sections stay parallel."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background(slide)
    k = kicker(slide, kicker_label)
    claim = text(slide, claim_lines, (MARGIN, 1.12, 11.6, 1.9), 58, spacing=0.98, name="claim")

    a = cover(slide, img_a, (MARGIN, 3.25, 5.35, 3.0), keep=(0.03, 0.03, 0.97, 0.97), anchor=(0.5, 0.42))
    la = text(slide, label_a, (MARGIN, 6.38, 5.35, 0.45), 20, color=ACCENT, bold=True,
              tracking=1.2, name="label_a")
    arrow = text(slide, "→", (6.25, 4.3, 0.85, 0.9), 44, color=ACCENT, align=PP_ALIGN.CENTER, name="arrow")
    b = cover(slide, img_b, (7.1, 3.25, 5.35, 3.0), keep=(0.03, 0.03, 0.97, 0.97), anchor=(0.5, 0.42))
    lb = text(slide, label_b, (7.1, 6.38, 5.35, 0.45), 20, color=ACCENT, bold=True,
              tracking=1.2, name="label_b")
    src = credit(slide)
    add_steps(slide, [[a, la, arrow, b, lb, src, k], [claim]])
    notes(slide, note)
    return slide


def main() -> None:
    BUILD.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    slide_1(prs)

    character_slide(
        prs,
        kicker_label="01  FORREST",
        claim_lines="HE NEVER STOPPED.",
        img_a="image1.png", label_a="CARRIES OTHERS",
        img_b="image21.jpeg", label_b="KEEPS RUNNING",
        note=(
            "阿甘（Baby 1）：\n"
            "主张：他从不算计，也从不回头——不是聪明，是一种不含算计的坚持。\n"
            "左图：越南战场上他把 Bubba 背出来，他自己受了伤也没把朋友丢下。\n"
            "右图：没有人追他，他跑了三年多，身后跟了一群人。\n"
            "点给观众：美国的乐观与坚持，落在具体行为上就是'先做，再想'。\n"
            "口播：He is not smart. He is unstoppable - and that is a choice, not a talent.",
        ),
    )

    character_slide(
        prs,
        kicker_label="02  JENNY",
        claim_lines="SHE CHOSE FREEDOM.",
        img_a="image19.jpeg", label_a="RUNS AWAY",
        img_b="image2.png", label_b="COMES BACK",
        note=(
            "珍妮（Baby 2）：\n"
            "主张：她代表美国性格里的另一面——个人自由，以及自由的代价。\n"
            "左图：她第一个喊出 Run, Forrest, Run!——她自己一辈子也在跑，离开家、离开小镇。\n"
            "右图：她最后还是回来了。\n"
            "点给观众：自由不是免费的，她付了钱；这一点让'美国性格'不再只是好话。\n"
            "口播：She wanted to be free more than she wanted to be safe. That costs something.",
        ),
    )

    character_slide(
        prs,
        kicker_label="03  HIS MOTHER",
        claim_lines="SHE SAID: NO DIFFERENT.",
        img_a="image9.jpeg", label_a="SENDS HIM OUT",
        img_b="image12.jpeg", label_b="STAYS TO THE END",
        note=(
            "阿甘妈妈（Baby 3）：\n"
            "主张：她给了他一生最有用的一句话——你和别人没有不同。\n"
            "左图：她坚持把戴腿支架的他送进普通学校，不让他待在角落。\n"
            "右图：临终她还在跟他讲道理（Death is part of life）。\n"
            "点给观众：美国性格里的平等信念，是从家庭里这样一句话开始的。\n"
            "口播：She never asked him to become someone else. That is where his confidence comes from.",
        ),
    )

    character_slide(
        prs,
        kicker_label="04  LIEUTENANT DAN",
        claim_lines="HIS PLAN DIED. HE LIVED.",
        img_a="image7.png", label_a="WANTED TO DIE",
        img_b="image18.jpeg", label_b="CHOSE TO LIVE",
        note=(
            "丹中尉（Baby 4）：\n"
            "主张：他把人生写成一张计划表，计划失败了，他先是恨，最后选择活。\n"
            "左图：他相信自己的命运是像祖辈一样战死，被救下来＝耻辱。\n"
            "右图：风暴那一夜他对着天吼，活下来了，之后对阿甘说了谢谢。\n"
            "点给观众：美国性格里还有一种勇气——承认计划失败之后，重新开始。\n"
            "口播：He wanted to die a hero. He ended up alive, thanking the man he blamed.",
        ),
    )

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background(slide)
    k = kicker(slide, "THE POINT")
    line1 = text(slide, "NOT ONE HERO.", (MARGIN, 1.15, 11.6, 1.3), 66, name="line1")
    line2 = text(slide, "ALL OF THEM.", (MARGIN, 2.5, 11.6, 1.3), 66, color=ACCENT, name="line2")
    support = text(slide, "Together they make American character.",
                   (MARGIN, 4.15, 11.6, 0.7), 26, color=WHITE, bold=False, name="support")
    strip = []
    for i, img in enumerate(["image1.png", "image19.jpeg", "image12.jpeg", "image18.jpeg"]):
        strip.append(cover(slide, img, (MARGIN + i * 3.01, 5.2, 2.76, 1.6),
                           keep=(0.04, 0.06, 0.96, 0.94), anchor=(0.5, 0.4)))
    src = credit(slide)
    add_steps(slide, [[strip, src, k], [line1], [line2], [support]])
    notes(
        slide,
        "收束（建议主管理人来讲）：\n"
        "回到开场的问题：四个人有什么共同点？答案是——他们合起来才构成 American character。\n"
        "四个人各代表一面：坚持（阿甘）、自由（珍妮）、平等（妈妈）、重新开始（丹中尉）。\n"
        "带走的一句话：判断一个国家的性格，看它的人怎么对待失败和离别。\n"
        "口播：American character is not one man's story. It is four of them, living together.",
    )

    prs.save(OUT_FILE)
    print(f"saved {OUT_FILE} ({OUT_FILE.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
