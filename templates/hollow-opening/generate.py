"""镂空开场 · Hollow Opening

生成一份 3 页的 PPT，页与页之间用「平滑」切换：
  第 1 页：全黑
  第 2 页：镂空大字从中间切成上下两半，分别从右、左滑入拼合，字里透出底图
  第 3 页：上下两半黑色移出画面，底图放大，标题、副标题、署名淡入

所有文字都会转成形状，换一台没装字体的电脑打开，效果也不会变。
PowerPoint 2019 及以上 / Microsoft 365 / WPS 都能播放平滑切换。

用法：
  python generate.py --word "MY FUTURE" --title "规划未来 职引人生" --image 我的图.jpg
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

SW, SH = 12192000, 6858000  # 16:9 页面尺寸（EMU）
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
A = "{%s}" % NS["a"]


# ---------------- 文字转形状 ----------------
def _font(path):
    f = TTFont(str(path))
    return f, f.getGlyphSet(), f.getBestCmap(), f["head"].unitsPerEm


def text_path(font_path, text, box, tracking=0.0):
    """把一行字等比缩放后居中放进 box=(x0, y0, x1, y1)，返回路径。"""
    f, gs, cmap, upm = _font(font_path)
    placed, x = [], 0.0
    for ch in text:
        if ch == " ":
            x += upm * 0.25
            continue
        g = cmap.get(ord(ch))
        if g is None:
            raise SystemExit(f"字体 {Path(font_path).name} 里没有「{ch}」这个字，请换一个字体")
        placed.append((g, x))
        x += gs[g].width + tracking * upm
    bp = BoundsPen(gs)
    for g, ox in placed:
        gs[g].draw(TransformPen(bp, (1, 0, 0, 1, ox, 0)))
    bx0, by0, bx1, by1 = bp.bounds
    x0, y0, x1, y1 = box
    s = min((x1 - x0) / (bx1 - bx0), (y1 - y0) / (by1 - by0))
    ox0 = x0 + ((x1 - x0) - (bx1 - bx0) * s) / 2
    oy0 = y0 + ((y1 - y0) - (by1 - by0) * s) / 2
    p = pathops.Path()
    pen = p.getPen()
    for g, ox in placed:
        gs[g].draw(TransformPen(pen, (s, 0, 0, -s, ox0 + (ox - bx0) * s, oy0 + by1 * s)))
    return pathops.simplify(p, fix_winding=True, keep_starting_points=False)


def spaced_path(font_path, text, cy, max_width, char_size):
    """字距随字数自动调整的一行字（副标题用），以页面中线居中。"""
    f, gs, cmap, upm = _font(font_path)
    chars = [c for c in text if c != " "]
    units = len(chars) + text.count(" ") * 0.6
    step = min(max_width / max(units - 1, 1), char_size * 1.6)
    cw = min(char_size, step * 0.82)
    x = 0.5 - step * (units - 1) / 2
    p = pathops.Path()
    pen = p.getPen()
    for ch in text:
        if ch == " ":
            x += step * 0.6
            continue
        g = cmap.get(ord(ch))
        if g is None:
            raise SystemExit(f"字体 {Path(font_path).name} 里没有「{ch}」这个字，请换一个字体")
        bp = BoundsPen(gs)
        gs[g].draw(bp)
        if bp.bounds:
            gx0, gy0, gx1, gy1 = bp.bounds
            half = cw * SW / 2
            s = min(2 * half / (gx1 - gx0), 2 * half / (gy1 - gy0))
            cx = x * SW
            gs[g].draw(TransformPen(pen, (s, 0, 0, -s, cx - (gx0 + gx1) / 2 * s, cy * SH + (gy0 + gy1) / 2 * s)))
        x += step
    return pathops.simplify(p, fix_winding=True, keep_starting_points=False)


def rect_path(x0, y0, x1, y1):
    p = pathops.Path()
    pen = p.getPen()
    pen.moveTo((x0, y0)); pen.lineTo((x1, y0)); pen.lineTo((x1, y1)); pen.lineTo((x0, y1)); pen.closePath()
    return p


def moved(p, dx, dy):
    q = pathops.Path()
    p.draw(TransformPen(q.getPen(), (1, 0, 0, 1, dx, dy)))
    return q


def _segments(path):
    """把路径拆成 moveTo / lineTo / 二次 / 三次贝塞尔 / close。"""
    for verb, pts in path.segments:
        if verb == "qCurveTo":
            offs, end = list(pts[:-1]), pts[-1]
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
    _next_id[0] += 1
    sp = etree.fromstring(f"""
<p:sp xmlns:p="{NS['p']}" xmlns:a="{NS['a']}">
  <p:nvSpPr><p:cNvPr id="{_next_id[0]}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{int(x0)}" y="{int(y0)}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
  {fill_xml}<a:ln><a:noFill/></a:ln>{effect_xml}</p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>
</p:sp>""")
    sp.find("p:spPr", NS).insert(1, geom)
    slide.shapes._spTree.append(sp)


def solid(color):
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


def add_picture(slide, name, img, frame, crop):
    x, y, w, h = frame
    pic = slide.shapes.add_picture(str(img), Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = crop
    pic._element.nvPicPr.cNvPr.set("name", name)


def black_background(slide):
    bg = etree.fromstring(f'<p:bg xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:bgPr>'
                          '<a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>')
    slide._element.find("p:cSld", NS).insert(0, bg)


def add_morph(slide, ms):
    el = etree.fromstring(f"""<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">
    <p:transition xmlns:p="{NS['p']}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="slow" p14:dur="{ms}"><p159:morph option="byObject"/></p:transition>
  </mc:Choice>
  <mc:Fallback><p:transition xmlns:p="{NS['p']}" spd="slow"><p:fade/></p:transition></mc:Fallback>
</mc:AlternateContent>""")
    slide._element.find("p:clrMapOvr", NS).addnext(el)


TITLE_FILL = ('<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>'
              '<a:gs pos="55000"><a:srgbClr val="F4F4F4"/></a:gs><a:gs pos="100000"><a:srgbClr val="A9A9A9"/></a:gs>'
              '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>')
TITLE_SHADOW = ('<a:effectLst><a:outerShdw blurRad="190500" dist="63500" dir="2700000" algn="tl" rotWithShape="0">'
                '<a:srgbClr val="000000"><a:alpha val="45000"/></a:srgbClr></a:outerShdw></a:effectLst>')
SOFT_SHADOW = ('<a:effectLst><a:outerShdw blurRad="76200" dist="0" dir="0" algn="ctr" rotWithShape="0">'
               '<a:srgbClr val="000000"><a:alpha val="70000"/></a:srgbClr></a:outerShdw></a:effectLst>')
BAR_FILL = ('<a:gradFill rotWithShape="1"><a:gsLst>'
            '<a:gs pos="0"><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:gs>'
            '<a:gs pos="50000"><a:srgbClr val="FFFFFF"><a:alpha val="78000"/></a:srgbClr></a:gs>'
            '<a:gs pos="100000"><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:gs>'
            '</a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>')


def build(a):
    X = lambda v: v * SW
    Y = lambda v: v * SH
    shift_x, shift_y, overlap, mid = 0.95, 0.55, Y(0.002), Y(0.5)

    # 镂空大字，上下两条黑条（两条在中线重叠一点点，放映时不会露缝）
    word = text_path(a.word_font, a.word, (X(0.06), Y(0.25), X(0.94), Y(0.75)))
    top = pathops.op(rect_path(-shift_x * SW, -overlap, SW + overlap, mid + overlap), word,
                     pathops.PathOp.DIFFERENCE, fix_winding=True)
    bot = pathops.op(rect_path(-overlap, mid - overlap, SW + shift_x * SW, SH + overlap), word,
                     pathops.PathOp.DIFFERENCE, fix_winding=True)

    # 底图：图片框等于整页，按 zoom / 偏移裁剪；第 3 页以中心放大 1.3 倍
    iw, ih = Image.open(a.image).size
    r = (iw / ih) / (SW / SH)            # 图片比 16:9 宽多少
    cw = a.zoom * max(1, r)
    ch = a.zoom * max(1, 1 / r)
    ox = -(cw - 1) / 2 * (1 + a.image_x)
    oy = -(ch - 1) / 2 * (1 + a.image_y)
    crop = (-ox / cw, -oy / ch, (cw + ox - 1) / cw, (ch + oy - 1) / ch)
    k = 1.3
    frame3 = (X((1 - k) / 2), Y((1 - k) / 2), X(k), Y(k))

    prs = Presentation()
    prs.slide_width, prs.slide_height = SW, SH
    slides = [prs.slides.add_slide(prs.slide_layouts[6]) for _ in range(3)]
    for s in slides:
        black_background(s)
    black = solid("000000")

    add_picture(slides[0], "!!bg", a.image, (0, 0, SW, SH), crop)
    add_shape(slides[0], "!!top", moved(top, shift_x * SW, 0), black)
    add_shape(slides[0], "!!bot", moved(bot, -shift_x * SW, 0), black)

    add_picture(slides[1], "!!bg", a.image, (0, 0, SW, SH), crop)
    add_shape(slides[1], "!!top", top, black)
    add_shape(slides[1], "!!bot", bot, black)
    add_morph(slides[1], a.duration)

    add_picture(slides[2], "!!bg", a.image, frame3, crop)
    add_shape(slides[2], "!!top", moved(top, 0, -shift_y * SH), black)
    add_shape(slides[2], "!!bot", moved(bot, 0, shift_y * SH), black)
    parts = {"top": top, "bot": bot}
    if a.title:
        parts["title"] = text_path(a.title_font, a.title, (X(0.107), Y(0.265), X(0.856), Y(0.5)), tracking=0.02)
        add_shape(slides[2], "标题", parts["title"], TITLE_FILL, TITLE_SHADOW)
    if a.subtitle:
        parts["sub"] = spaced_path(a.subtitle_font, a.subtitle, 0.577, 0.5, 0.026)
        add_shape(slides[2], "副标题", parts["sub"], solid(a.subtitle_color.lstrip("#")), SOFT_SHADOW)
    if a.reporter:
        add_shape(slides[2], "署名底条", rect_path(X(0.29), Y(0.694), X(0.71), Y(0.745)), BAR_FILL)
        parts["rep"] = text_path(a.reporter_font, a.reporter, (X(0.30), Y(0.705), X(0.70), Y(0.733)), tracking=0.18)
        add_shape(slides[2], "署名", parts["rep"], solid("2B2B2B"))
    add_morph(slides[2], a.duration)

    prs.save(a.out)
    print("已生成：", a.out)
    if a.preview:
        img2 = (ox, oy, cw, ch)
        img3 = ((1 - k) / 2 + k * ox, (1 - k) / 2 + k * oy, cw * k, ch * k)
        write_preview(a, parts, img2, img3)


# ---------------- 网页预览（可选） ----------------
def svg_d(path, k=1920 / SW):
    out = []
    for kind, pts in _segments(path):
        f = lambda q: f"{q[0] * k:.1f} {q[1] * k:.1f}"
        out.append({"m": "M", "l": "L", "q": "Q", "c": "C", "z": "Z"}[kind] + " ".join(f(q) for q in pts))
    return "".join(out)


def write_preview(a, parts, img2, img3):
    im = Image.open(a.image).convert("RGB")
    im.thumbnail((2400, 2400))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=82)
    box = lambda r: "%.1f %.1f %.1f %.1f" % (r[0] * 1920, r[1] * 1080, r[2] * 1920, r[3] * 1080)
    html = PREVIEW
    for key, val in {
        "__IMG__": "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode(),
        "__TOP__": svg_d(parts["top"]), "__BOT__": svg_d(parts["bot"]),
        "__TITLE__": svg_d(parts["title"]) if "title" in parts else "",
        "__SUB__": svg_d(parts["sub"]) if "sub" in parts else "",
        "__REP__": svg_d(parts["rep"]) if "rep" in parts else "",
        "__BAR__": box((0.29, 0.694, 0.42, 0.051)) if "rep" in parts else "0 0 0 0",
        "__SUBC__": a.subtitle_color, "__IMG2__": box(img2), "__IMG3__": box(img3),
    }.items():
        html = html.replace(key, val)
    Path(a.preview).write_text(html, encoding="utf-8")
    print("已生成网页预览：", a.preview)


PREVIEW = r"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>镂空开场预览</title>
<style>html,body{margin:0;background:#111;color:#ddd;font-family:sans-serif}
.stage{position:relative;max-width:1200px;margin:16px auto;aspect-ratio:16/9;background:#000;overflow:hidden;cursor:pointer}
svg{position:absolute;inset:0;width:100%;height:100%}p{text-align:center;font-size:13px;color:#888}</style></head><body>
<div class="stage" id="stage"><svg viewBox="0 0 1920 1080" preserveAspectRatio="xMidYMid slice"><defs>
<linearGradient id="tg" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#fff"/><stop offset=".55" stop-color="#f4f4f4"/><stop offset="1" stop-color="#a9a9a9"/></linearGradient>
<linearGradient id="cg" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#fff" stop-opacity="0"/><stop offset=".5" stop-color="#fff" stop-opacity=".78"/><stop offset="1" stop-color="#fff" stop-opacity="0"/></linearGradient>
<filter id="sh" x="-10%" y="-20%" width="120%" height="150%"><feDropShadow dx="3.5" dy="3.5" stdDeviation="7" flood-color="#000" flood-opacity=".45"/></filter>
<filter id="ss" x="-5%" y="-50%" width="110%" height="200%"><feDropShadow dx="0" dy="0" stdDeviation="3" flood-color="#000" flood-opacity=".7"/></filter></defs>
<rect width="1920" height="1080" fill="#000"/><image id="bg" href="__IMG__" preserveAspectRatio="none"/>
<path id="top" d="__TOP__"/><path id="bot" d="__BOT__"/>
<g id="ttl" opacity="0"><path d="__TITLE__" fill="url(#tg)" filter="url(#sh)"/><path d="__SUB__" fill="__SUBC__" filter="url(#ss)"/>
<rect id="bar" fill="url(#cg)"/><path d="__REP__" fill="#2B2B2B"/></g></svg></div>
<p>点画面或按 → 翻页（网页模拟效果，最终以 PPT 放映为准）</p>
<script>
const W=1920,H=1080,I2="__IMG2__".split(" ").map(Number),I3="__IMG3__".split(" ").map(Number),B="__BAR__".split(" ").map(Number);
const bar=document.getElementById("bar");["x","y","width","height"].forEach((k,i)=>bar.setAttribute(k,B[i]));
const S=[{tx:.95*W,ty:0,bx:-.95*W,by:0,img:I2,t:0},{tx:0,ty:0,bx:0,by:0,img:I2,t:0},{tx:0,ty:-.55*H,bx:0,by:.55*H,img:I3,t:1}];
const $=id=>document.getElementById(id),L=(a,b,t)=>a+(b-a)*t;
function apply(s){$("top").setAttribute("transform",`translate(${s.tx} ${s.ty})`);$("bot").setAttribute("transform",`translate(${s.bx} ${s.by})`);
["x","y","width","height"].forEach((k,i)=>$("bg").setAttribute(k,s.img[i]));$("ttl").setAttribute("opacity",s.t)}
function mix(a,b,t){const last=b.t>a.t,e=last?(1-Math.cos(Math.PI*t))/2:1-Math.pow(1-t,2.4),f=last?Math.max(0,(t-.6)/.4):0;
return{tx:L(a.tx,b.tx,e),ty:L(a.ty,b.ty,e),bx:L(a.bx,b.bx,e),by:L(a.by,b.by,e),img:a.img.map((v,i)=>L(v,b.img[i],e)),t:L(a.t,b.t,f)}}
let cur=0,raf;function go(n){n=Math.max(0,Math.min(2,n));if(n===cur)return;cancelAnimationFrame(raf);const A=S[cur],Bs=S[n],back=n<cur;cur=n;
if(back){apply(Bs);return}const d=n===2?1450:1700,t0=performance.now();(function f(now){const t=Math.min(1,(now-t0)/d);apply(mix(A,Bs,t));if(t<1)raf=requestAnimationFrame(f)})(t0)}
apply(S[0]);$("stage").onclick=()=>go(cur>=2?0:cur+1);
addEventListener("keydown",e=>{if(["ArrowRight"," ","Enter","PageDown"].includes(e.key)){e.preventDefault();go(cur+1)}if(["ArrowLeft","PageUp"].includes(e.key))go(cur-1)});
window.preview={S,mix,apply};
</script></body></html>"""


def main():
    ap = argparse.ArgumentParser(description="镂空开场：生成带平滑切换的 3 页 PPT")
    ap.add_argument("--word", default="MY FUTURE", help="镂空大字（英文、数字、中文都可以）")
    ap.add_argument("--title", default="规划未来 职引人生", help="第 3 页的大标题，留空则不显示")
    ap.add_argument("--subtitle", default="大学生职业生涯规划大赛", help="副标题，留空则不显示")
    ap.add_argument("--reporter", default="汇报人：XXX", help="署名，留空则不显示")
    ap.add_argument("--image", default=str(HERE / "example.jpg"), help="底图，横图最佳")
    ap.add_argument("--zoom", type=float, default=1.14, help="第 1、2 页底图的放大倍数（≥1）")
    ap.add_argument("--image-x", type=float, default=0.0, help="底图左右取景，-1 露最左边，0 居中，1 露最右边")
    ap.add_argument("--image-y", type=float, default=-0.15, help="底图上下取景，-1 露最上面，0 居中，1 露最下面")
    ap.add_argument("--word-font", default=str(FONTS / "Anton-Regular.ttf"), help="镂空大字的字体文件")
    ap.add_argument("--title-font", default=str(FONTS / "MaShanZheng-Regular.ttf"), help="标题字体文件")
    ap.add_argument("--subtitle-font", default=str(FONTS / "NotoSerifSC-Bold.ttf"), help="副标题字体文件")
    ap.add_argument("--reporter-font", default=str(FONTS / "NotoSansSC-Medium.ttf"), help="署名字体文件")
    ap.add_argument("--subtitle-color", default="#F2CF5B", help="副标题颜色")
    ap.add_argument("--duration", type=int, default=2000, help="平滑切换时长（毫秒）")
    ap.add_argument("--out", default="镂空开场.pptx", help="输出的 PPT 文件名")
    ap.add_argument("--preview", default=None, help="同时输出网页预览，例如 --preview 预览.html")
    build(ap.parse_args())


if __name__ == "__main__":
    main()
