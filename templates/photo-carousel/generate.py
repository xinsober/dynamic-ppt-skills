"""多图轮播 · Photo Carousel

生成 1 页 PPT：一排照片在上下两条弧形色带之间从右往左慢慢滚动，
放映时自动开始（直线路径动画，与上一动画同时）。适合活动照片很多的展示页。

所有文字都会转成形状，换一台没装字体的电脑打开，效果也不会变。

用法：
  python generate.py --images 1.jpg 2.jpg 3.jpg 4.jpg 5.jpg 6.jpg --title "并肩向前 共赴新程"
  python generate.py --help      查看全部参数
"""
import argparse
import base64
import io
from pathlib import Path

import pathops
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.util import Emu

HERE = Path(__file__).resolve().parent
FONTS = HERE.parent.parent / "fonts"

SW, SH = 12192000, 6858000      # 16:9 页面（EMU）
PX = SW / 1920                  # 以 1920×1080 为坐标，1 像素对应的 EMU
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]

# 版式（单位：像素，页面 1920×1080）
PHOTO_TOP, PHOTO_H, GAP = 170, 816, 38          # 照片一排：顶边、统一高度、间距
ARC_RX = 1900                                   # 弧线所在椭圆的横半径
TOP_ARC = (-373, 635)                           # 上色带：椭圆圆心 y、竖半径（下沿中间 262、两边约 175）
BOT_ARC = (1545, 650)                           # 下色带：椭圆圆心 y、竖半径（上沿中间 895、两边约 984）
WHITE = 22                                      # 白边宽度
SPEED_REF = 272                                 # 默认 20 秒、6 张图时的滚动速度（像素/秒）


# ---------------- 文字转形状 ----------------
_font_cache = {}


def _font(path):
    path = str(path)
    if path not in _font_cache:
        f = TTFont(path)
        _font_cache[path] = (f, f.getGlyphSet(), f.getBestCmap(), f["head"].unitsPerEm)
    return _font_cache[path]


def line_path(font_path, text, em, cx, cy, tracking=0.0):
    """一行字：字号 em 像素，水平居中于 cx，字身（汉字方框）垂直居中于 cy。"""
    f, gs, cmap, upm = _font(font_path)
    s = em / upm
    pen_x, placed = 0.0, []
    for ch in text:
        if ch == " ":
            pen_x += upm * 0.3
            continue
        g = cmap.get(ord(ch))
        if g is None:
            raise SystemExit(f"字体 {Path(font_path).name} 里没有「{ch}」这个字，请换一个字体")
        placed.append((g, pen_x))
        pen_x += gs[g].width + tracking * upm
    width = (pen_x - tracking * upm) * s
    x0 = cx - width / 2
    base = cy + 0.38 * em       # 汉字字身大约在基线上 -0.12em 到 0.88em
    p = pathops.Path()
    pen = p.getPen(glyphSet=gs)
    for g, ox in placed:
        gs[g].draw(TransformPen(pen, (s * PX, 0, 0, -s * PX, (x0 + ox * s) * PX, base * PX)))
    return pathops.simplify(p, fix_winding=True, keep_starting_points=False)


def text_width(font_path, text, em):
    f, gs, cmap, upm = _font(font_path)
    w = sum(upm * 0.3 if ch == " " else gs[cmap[ord(ch)]].width for ch in text if ch == " " or ord(ch) in cmap)
    return w * em / upm


def wrap_lines(font_path, text, em, max_w):
    """按宽度自动换行；文案里自己写的换行会保留。"""
    out = []
    for para in text.split("\n"):
        line = ""
        for ch in para:
            if text_width(font_path, line + ch, em) > max_w and line:
                out.append(line)
                line = ch
            else:
                line += ch
        if line:
            out.append(line)
    return out


# ---------------- 几何 ----------------
def rect_path(x0, y0, x1, y1):
    p = pathops.Path()
    pen = p.getPen()
    pen.moveTo((x0 * PX, y0 * PX)); pen.lineTo((x1 * PX, y0 * PX))
    pen.lineTo((x1 * PX, y1 * PX)); pen.lineTo((x0 * PX, y1 * PX)); pen.closePath()
    return p


def ellipse_path(cx, cy, rx, ry):
    k = 0.5522847498
    cx, cy, rx, ry = cx * PX, cy * PX, rx * PX, ry * PX
    p = pathops.Path()
    pen = p.getPen()
    pen.moveTo((cx + rx, cy))
    pen.curveTo((cx + rx, cy + k * ry), (cx + k * rx, cy + ry), (cx, cy + ry))
    pen.curveTo((cx - k * rx, cy + ry), (cx - rx, cy + k * ry), (cx - rx, cy))
    pen.curveTo((cx - rx, cy - k * ry), (cx - k * rx, cy - ry), (cx, cy - ry))
    pen.curveTo((cx + k * rx, cy - ry), (cx + rx, cy - k * ry), (cx + rx, cy))
    pen.closePath()
    return p


def band(top, dy=0.0):
    """上/下弧形色带：整页矩形和椭圆相交后剩下的一条（合并形状 → 拆分的结果）。"""
    bleed = 4
    if top:
        cy, ry = TOP_ARC
        box = rect_path(-bleed, -bleed + dy, 1920 + bleed, cy + ry + dy)
    else:
        cy, ry = BOT_ARC
        box = rect_path(-bleed, cy - ry + dy, 1920 + bleed, 1080 + bleed + dy)
    return pathops.op(box, ellipse_path(960, cy + dy, ARC_RX, ry), pathops.PathOp.INTERSECTION, fix_winding=True)


def _segments(path):
    for verb, pts in path.segments:
        if verb == "qCurveTo":
            offs, end = list(pts[:-1]), pts[-1]
            if end is None:  # 整圈都是控制点的闭合曲线
                offs = list(pts[:-1])
                start = ((offs[-1][0] + offs[0][0]) / 2, (offs[-1][1] + offs[0][1]) / 2)
                yield "m", (start,)
                for i, c in enumerate(offs):
                    n = offs[(i + 1) % len(offs)]
                    yield "q", (c, ((c[0] + n[0]) / 2, (c[1] + n[1]) / 2))
                continue
            for i, c in enumerate(offs):
                nxt = ((c[0] + offs[i + 1][0]) / 2, (c[1] + offs[i + 1][1]) / 2) if i < len(offs) - 1 else end
                yield "q", (c, nxt)
        elif verb == "moveTo":
            yield "m", pts
        elif verb == "lineTo":
            yield "l", pts
        elif verb == "curveTo":
            yield "c", pts
        else:
            yield "z", ()


# ---------------- 写入 PPT ----------------
_next_id = [100]


def new_id():
    _next_id[0] += 1
    return _next_id[0]


def add_shape(slide, name, path, fill_xml, effect_xml=""):
    x0, y0, x1, y1 = path.bounds
    w, h = int(x1 - x0), int(y1 - y0)
    geom = etree.Element(A + "custGeom")
    for t in ("avLst", "gdLst", "ahLst", "cxnLst"):
        etree.SubElement(geom, A + t)
    etree.SubElement(geom, A + "rect", l="l", t="t", r="r", b="b")
    pe = etree.SubElement(etree.SubElement(geom, A + "pathLst"), A + "path", w=str(w), h=str(h))
    tag = {"m": "moveTo", "l": "lnTo", "q": "quadBezTo", "c": "cubicBezTo"}
    for kind, pts in _segments(path):
        if kind == "z":
            etree.SubElement(pe, A + "close")
            continue
        e = etree.SubElement(pe, A + tag[kind])
        for q in pts:
            etree.SubElement(e, A + "pt", x=str(round(q[0] - x0)), y=str(round(q[1] - y0)))
    sp = etree.fromstring(f"""
<p:sp xmlns:p="{NS['p']}" xmlns:a="{NS['a']}">
  <p:nvSpPr><p:cNvPr id="{new_id()}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{int(x0)}" y="{int(y0)}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
  {fill_xml}<a:ln><a:noFill/></a:ln>{effect_xml}</p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>
</p:sp>""")
    sp.find("p:spPr", NS).insert(1, geom)
    slide.shapes._spTree.append(sp)


def solid(color):
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


SHADOW = ('<a:effectLst><a:outerShdw blurRad="127000" dist="0" dir="5400000" algn="ctr" rotWithShape="0">'
          '<a:srgbClr val="000000"><a:alpha val="55000"/></a:srgbClr></a:outerShdw></a:effectLst>')
CAPSULE = ('<a:gradFill rotWithShape="1"><a:gsLst>'
           '<a:gs pos="0"><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:gs>'
           '<a:gs pos="30000"><a:srgbClr val="FFFFFF"/></a:gs>'
           '<a:gs pos="70000"><a:srgbClr val="FFFFFF"/></a:gs>'
           '<a:gs pos="100000"><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:gs>'
           '</a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>')


def add_motion(slide, spid, dx, ms):
    """直线路径动画：放映开始即运行（与上一动画同时），匀速。dx 以页宽为单位，负数向左。"""
    timing = etree.fromstring(f"""
<p:timing xmlns:p="{NS['p']}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="5" presetID="0" presetClass="path" presetSubtype="0" accel="0" decel="0" fill="hold" nodeType="withEffect">
<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:animMotion origin="layout" path="M 0 0 L {dx:.6f} 0 E" pathEditMode="relative" rAng="0" ptsTypes="AA">
<p:cBhvr><p:cTn id="6" dur="{ms}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>
<p:rCtr x="{dx * 50000:.0f}" y="0"/></p:animMotion>
</p:childTnLst></p:cTn></p:par>
</p:childTnLst></p:cTn></p:par>
</p:childTnLst></p:cTn></p:par>
</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>
</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>""")
    slide._element.append(timing)


def load_photos(paths):
    """每张照片统一高度，宽度按原比例。返回 [(路径, 宽像素)]。"""
    out = []
    for p in paths:
        w, h = Image.open(p).size
        out.append((Path(p), PHOTO_H * w / h))
    return out


# 文字版式（像素）
TITLE_EM, TITLE_Y = 84, 92          # 大标题字号、中线位置
SUB_EM, SUB_Y, SUB_PAD = 36, 186, 110   # 小字条字号、中线位置、左右留白
TEXT_EM, TEXT_Y, TEXT_STEP, TEXT_W = 34, 988, 50, 1300   # 说明文字字号、中线位置、行距、最大行宽


def text_shapes(a):
    """所有文字和小字条底，返回 [(名字, 路径, 颜色类型)]。"""
    out = []
    if a.title:
        out.append(("标题", line_path(a.title_font, a.title, TITLE_EM, 960, TITLE_Y), "white"))
    if a.subtitle:
        w = text_width(a.subtitle_font, a.subtitle, SUB_EM) + SUB_PAD * 2
        h = SUB_EM * 1.55
        out.append(("小字条底", rect_path(960 - w / 2, SUB_Y - h / 2, 960 + w / 2, SUB_Y + h / 2), "cap"))
        out.append(("小字条", line_path(a.subtitle_font, a.subtitle, SUB_EM, 960, SUB_Y), "color"))
    if a.text:
        lines = wrap_lines(a.text_font, a.text, TEXT_EM, TEXT_W)[:2]
        for i, ln in enumerate(lines):
            cy = TEXT_Y + (i - (len(lines) - 1) / 2) * TEXT_STEP
            out.append((f"说明{i + 1}", line_path(a.text_font, ln, TEXT_EM, 960, cy), "white"))
    return out


def build(a):
    photos = load_photos(a.images)
    strip_w = sum(w for _, w in photos) + GAP * (len(photos) - 1)
    travel = max(strip_w - 1920, 0)             # 从第一张靠左，滚到最后一张靠右
    ms = int(a.duration * 1000) if a.duration else int(max(travel, 1) / SPEED_REF * 1000)

    prs = Presentation()
    prs.slide_width, prs.slide_height = SW, SH
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 1. 照片排成一排并组合
    grp = slide.shapes.add_group_shape()
    grp._element.nvGrpSpPr.cNvPr.set("name", "照片组")
    x = 0.0
    for i, (p, w) in enumerate(photos):
        pic = grp.shapes.add_picture(str(p), Emu(int(x * PX)), Emu(int(PHOTO_TOP * PX)),
                                     Emu(int(w * PX)), Emu(int(PHOTO_H * PX)))
        pic._element.nvPicPr.cNvPr.set("name", f"照片{i + 1}")
        x += w + GAP
    gid = grp.shape_id

    # 2. 白边（同形状白色，往内错开）在下，红色带加阴影在上
    add_shape(slide, "上白边", band(True, WHITE), solid("FFFFFF"))
    add_shape(slide, "下白边", band(False, -WHITE), solid("FFFFFF"))
    add_shape(slide, "上色带", band(True), solid(a.color), SHADOW)
    add_shape(slide, "下色带", band(False), solid(a.color), SHADOW)

    # 3. 文字
    for name, path, kind in text_shapes(a):
        add_shape(slide, name, path, CAPSULE if kind == "cap" else solid("FFFFFF" if kind == "white" else a.color))

    add_motion(slide, gid, -travel / 1920, ms)
    prs.save(a.out)
    print("已生成：", a.out)
    return photos, travel, ms


# ---------------- 网页预览 ----------------
def svg_d(path):
    s = 1 / PX
    out = []
    for kind, pts in _segments(path):
        if kind == "z":
            out.append("Z")
            continue
        out.append({"m": "M", "l": "L", "q": "Q", "c": "C"}[kind] + " ".join(f"{q[0] * s:.1f} {q[1] * s:.1f}" for q in pts))
    return "".join(out)


PREVIEW = """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>多图轮播预览</title>
<style>
:root{--bg:#0b0b0c;--fg:#e9e9ea;--muted:#8a8a90;--btn:#1d1d20;--line:#2a2a2e}
@media (prefers-color-scheme: light){:root:not([data-theme="dark"]){--bg:#f4f4f5;--fg:#18181b;--muted:#6b6b72;--btn:#fff;--line:#d9d9de}}
:root[data-theme="dark"]{--bg:#0b0b0c;--fg:#e9e9ea;--muted:#8a8a90;--btn:#1d1d20;--line:#2a2a2e}
html,body{margin:0;background:var(--bg);color:var(--fg);font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif}
.wrap{max-width:1200px;margin:0 auto;padding:16px}
.stage{position:relative;width:100%;aspect-ratio:16/9;background:#fff;overflow:hidden;border-radius:6px}
.stage svg{position:absolute;inset:0;width:100%;height:100%;display:block}
.bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:12px}
button{background:var(--btn);color:var(--fg);border:1px solid var(--line);border-radius:8px;padding:8px 14px;font-size:14px;cursor:pointer}
.hint{color:var(--muted);font-size:13px}
</style></head><body><div class="wrap">
<div class="stage"><svg viewBox="0 0 1920 1080">
<defs><filter id="sh" x="-5%" y="-30%" width="110%" height="160%"><feDropShadow dx="0" dy="0" stdDeviation="7" flood-color="#000" flood-opacity=".55"/></filter>
<linearGradient id="cap" x1="0" x2="1"><stop offset="0" stop-color="#fff" stop-opacity="0"/><stop offset=".3" stop-color="#fff"/><stop offset=".7" stop-color="#fff"/><stop offset="1" stop-color="#fff" stop-opacity="0"/></linearGradient></defs>
<rect width="1920" height="1080" fill="#fff"/>
<g id="strip">__PHOTOS__</g>
__SHAPES__
</svg></div>
<div class="bar"><button id="play">重新播放</button><span class="hint">和 PPT 放映一样，进入这一页就开始滚动</span></div>
</div>
<script>
const TRAVEL=__TRAVEL__,DUR=__DUR__,strip=document.getElementById("strip");let t0=performance.now(),raf;
function f(now){const t=Math.min(1,(now-t0)/DUR);strip.setAttribute("transform",`translate(${-TRAVEL*t} 0)`);if(t<1)raf=requestAnimationFrame(f)}
function play(){cancelAnimationFrame(raf);t0=performance.now();raf=requestAnimationFrame(f)}
document.getElementById("play").onclick=play;play();
window.preview={at:s=>{cancelAnimationFrame(raf);strip.setAttribute("transform",`translate(${-TRAVEL*Math.min(1,s*1000/DUR)} 0)`)}};
</script></body></html>"""


def write_preview(a, photos, travel, ms):
    imgs, x = [], 0.0
    for p, w in photos:
        im = Image.open(p).convert("RGB")
        im.thumbnail((1400, 1400))
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=82)
        uri = "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
        imgs.append(f'<image href="{uri}" x="{x:.1f}" y="{PHOTO_TOP}" width="{w:.1f}" height="{PHOTO_H}" preserveAspectRatio="none"/>')
        x += w + GAP
    shapes = [f'<path d="{svg_d(band(True, WHITE))}" fill="#fff"/>',
              f'<path d="{svg_d(band(False, -WHITE))}" fill="#fff"/>',
              f'<path d="{svg_d(band(True))}" fill="#{a.color}" filter="url(#sh)"/>',
              f'<path d="{svg_d(band(False))}" fill="#{a.color}" filter="url(#sh)"/>']
    for name, path, kind in text_shapes(a):
        fill = {"cap": "url(#cap)", "white": "#fff"}.get(kind, "#" + a.color)
        shapes.append(f'<path d="{svg_d(path)}" fill="{fill}"/>')
    html = (PREVIEW.replace("__PHOTOS__", "".join(imgs)).replace("__SHAPES__", "\n".join(shapes))
            .replace("__TRAVEL__", f"{travel:.1f}").replace("__DUR__", str(ms)))
    Path(a.preview).write_text(html, encoding="utf-8")
    print("已生成网页预览：", a.preview)


def main():
    ap = argparse.ArgumentParser(description="多图轮播：一排照片在弧形色带之间自动滚动")
    default_imgs = sorted(str(p) for p in HERE.glob("example*.jpg"))
    ap.add_argument("--images", nargs="+", default=default_imgs, help="照片，按顺序排成一排，建议 4～10 张横图")
    ap.add_argument("--title", default="并肩向前　共赴新程", help="上方大标题，留空则不显示")
    ap.add_argument("--subtitle", default="每一次全力以赴，都是团队的高光", help="标题下的一行小字，留空则不显示")
    ap.add_argument("--text", default="感谢每一位并肩作战的伙伴，新的一年继续一起出发", help="下方说明文字，一句话最好，最多 2 行，会自动换行")
    ap.add_argument("--color", default="C00000", help="色带颜色（十六进制，不带 #）")
    ap.add_argument("--duration", type=float, default=20, help="滚完一遍的秒数；填 0 则按照片数量自动算")
    ap.add_argument("--title-font", default=str(FONTS / "NotoSerifSC-Bold.ttf"), help="标题字体文件")
    ap.add_argument("--subtitle-font", default=str(FONTS / "NotoSansSC-Black.ttf"), help="小标题字体文件")
    ap.add_argument("--text-font", default=str(FONTS / "NotoSansSC-Medium.ttf"), help="说明文字字体文件")
    ap.add_argument("--out", default="多图轮播.pptx", help="输出的 PPT 文件名")
    ap.add_argument("--preview", default=None, help="同时输出网页预览，例如 --preview 预览.html")
    a = ap.parse_args()
    a.color = a.color.lstrip("#").upper()
    if len(a.images) < 2:
        raise SystemExit("至少需要 2 张照片")
    photos, travel, ms = build(a)
    if a.preview:
        write_preview(a, photos, travel, ms)


if __name__ == "__main__":
    main()
