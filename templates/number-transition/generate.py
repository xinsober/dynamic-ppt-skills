"""数字过渡页 · Number Transition

生成一组章节过渡页（每节一页）：整页铺图，超大斜体数字把画面切开，
数字右边是一整块半透明色块和标题。放映时底图从放大状态缓缓缩回，
数字、色块和文字一起从右边飞入。

所有文字都会转成形状，换一台没装字体的电脑打开，效果也不会变。

用法：
  python generate.py --images 1.jpg 2.jpg 3.jpg --titles "KNOW|YOURSELF" "EXPLORE|OPTIONS" "MAKE|A PLAN"
  python generate.py --help      查看全部参数
"""
import argparse
import base64
import io
from pathlib import Path

import pathops
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.util import Emu

HERE = Path(__file__).resolve().parent
FONTS = HERE.parent.parent / "fonts"

SW, SH = 12192000, 6858000      # 16:9 页面（EMU）
PX = SW / 1920                  # 以 1920×1080 为坐标
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
A = "{%s}" % NS["a"]

# 版式（像素，页面 1920×1080）
DIGIT_H = 1080          # 数字高度：上下正好贴着画面边，不出画面
DIGIT_SLANT = 0.3375    # 倾斜（水平偏移 / 高度）
DIGIT_CX = 620          # 数字中心的 x
SPLIT_X = 700           # 色块左边斜线在页面中线处的 x，斜度和数字一样
TITLE_EM = 124          # 标题字号（像素）
TITLE_X, TITLE_STEP = 1440, 150
BODY_EM, BODY_GAP, BODY_STEP = 38, 112, 56
START_ZOOM = 1.7        # 底图开场放大倍数


# ---------------- 文字转形状 ----------------
_cache = {}


def _font(path):
    path = str(path)
    if path not in _cache:
        f = TTFont(path)
        _cache[path] = (f, f.getGlyphSet(), f.getBestCmap(), f["head"].unitsPerEm)
    return _cache[path]


def _glyph(gs, cmap, ch, path):
    g = cmap.get(ord(ch))
    if g is None:
        raise SystemExit(f"字体 {Path(path).name} 里没有「{ch}」这个字，请换一个字体")
    return g


def line_path(font_path, text, em, x, base, align="center"):
    f, gs, cmap, upm = _font(font_path)
    s = em / upm
    placed, pen_x = [], 0.0
    for ch in text:
        if ch == " ":
            pen_x += gs[cmap[32]].width if 32 in cmap else upm * 0.25
            continue
        g = _glyph(gs, cmap, ch, font_path)
        placed.append((g, pen_x))
        pen_x += gs[g].width
    x0 = x - pen_x * s / 2 if align == "center" else x
    p = pathops.Path()
    pen = p.getPen(glyphSet=gs)
    for g, ox in placed:
        gs[g].draw(TransformPen(pen, (s * PX, 0, 0, -s * PX, (x0 + ox * s) * PX, base * PX)))
    return pathops.simplify(p, fix_winding=True, keep_starting_points=False)


def digit_path(font_path, ch):
    """超大斜体数字：上下正好占满画面高度，居中放在画面左半边。"""
    f, gs, cmap, upm = _font(font_path)
    g = pathops.Path()
    gs[_glyph(gs, cmap, ch, font_path)].draw(g.getPen(glyphSet=gs))
    s = DIGIT_H / (g.bounds[3] - g.bounds[1])
    k = DIGIT_SLANT
    p = pathops.Path()
    gs[_glyph(gs, cmap, ch, font_path)].draw(TransformPen(p.getPen(glyphSet=gs), (s * PX, 0, k * s * PX, -s * PX, 0, 0)))
    x0, y0, x1, y1 = p.bounds
    dx, dy = DIGIT_CX * PX - (x0 + x1) / 2, 540 * PX - (y0 + y1) / 2
    q = pathops.Path()
    p.draw(TransformPen(q.getPen(), (1, 0, 0, 1, dx, dy)))
    return pathops.simplify(q, fix_winding=True, keep_starting_points=False)


def _rect_units(x0, y0, x1, y1):
    p = pathops.Path()
    pen = p.getPen()
    pen.moveTo((x0, y0)); pen.lineTo((x1, y0)); pen.lineTo((x1, y1)); pen.lineTo((x0, y1)); pen.closePath()
    return p


def rect_path(x0, y0, x1, y1):
    return _rect_units(x0 * PX, y0 * PX, x1 * PX, y1 * PX)


def split_by_digit(digit):
    """色块是页面右边的一块斜边四边形，斜边和数字同样倾斜；数字的位置从色块里挖掉，两块不重叠。"""
    page = rect_path(0, 0, 1920, 1080)
    num = pathops.op(page, digit, pathops.PathOp.INTERSECTION, fix_winding=True)
    k = DIGIT_SLANT
    quad = pathops.Path()
    pen = quad.getPen()
    for x, y in ((SPLIT_X + k * 540, 0), (1920, 0), (1920, 1080), (SPLIT_X - k * 540, 1080)):
        (pen.moveTo if y == 0 and x != 1920 else pen.lineTo)((x * PX, y * PX))
    pen.closePath()
    panel = pathops.op(quad, digit, pathops.PathOp.DIFFERENCE, fix_winding=True)
    return num, panel


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


def shape_xml(name, path, fill_xml):
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
  {fill_xml}<a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>
</p:sp>""")
    sp.find("p:spPr", NS).insert(1, geom)
    return sp, (x0, y0, x1, y1)


def fill(color, alpha=100):
    a = "" if alpha >= 100 else f'<a:alpha val="{int(alpha * 1000)}"/>'
    return f'<a:solidFill><a:srgbClr val="{color}">{a}</a:srgbClr></a:solidFill>'


def add_group(slide, name, items):
    """把若干形状组合成一组，返回组的 id。"""
    x0 = min(b[0] for _, b in items); y0 = min(b[1] for _, b in items)
    x1 = max(b[2] for _, b in items); y1 = max(b[3] for _, b in items)
    gid = new_id()
    grp = etree.fromstring(f"""
<p:grpSp xmlns:p="{NS['p']}" xmlns:a="{NS['a']}">
  <p:nvGrpSpPr><p:cNvPr id="{gid}" name="{name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="{int(x0)}" y="{int(y0)}"/><a:ext cx="{int(x1 - x0)}" cy="{int(y1 - y0)}"/>
  <a:chOff x="{int(x0)}" y="{int(y0)}"/><a:chExt cx="{int(x1 - x0)}" cy="{int(y1 - y0)}"/></a:xfrm></p:grpSpPr>
</p:grpSp>""")
    for sp, _ in items:
        grp.append(sp)
    slide.shapes._spTree.append(grp)
    return gid


def cover_crop(img_path):
    """按 16:9 铺满裁剪。"""
    w, h = Image.open(img_path).size
    r = (w / h) / (16 / 9)
    if r > 1:
        c = (1 - 1 / r) / 2
        return (c, 0, c, 0)
    c = (1 - r) / 2
    return (0, c, 0, c)


def add_timing(slide, pic_id, grp_id, zoom_ms, fly_ms, fly_decel):
    """进入这一页即开始：底图从放大缩回原大（强调-放大/缩小），同时整组从右侧飞入。"""
    z = int(START_ZOOM * 100000)
    xml = f"""
<p:timing xmlns:p="{NS['p']}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>

<p:par><p:cTn id="5" presetID="6" presetClass="emph" presetSubtype="0" decel="100000" fill="hold" nodeType="withEffect">
<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:animScale><p:cBhvr><p:cTn id="6" dur="{zoom_ms}" fill="hold"/><p:tgtEl><p:spTgt spid="{pic_id}"/></p:tgtEl></p:cBhvr>
<p:from x="{z}" y="{z}"/><p:to x="100000" y="100000"/></p:animScale>
</p:childTnLst></p:cTn></p:par>

<p:par><p:cTn id="7" presetID="2" presetClass="entr" presetSubtype="2" decel="{int(fly_decel * 100000)}" fill="hold" nodeType="withEffect">
<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:set><p:cBhvr><p:cTn id="8" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
<p:tgtEl><p:spTgt spid="{grp_id}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>
<p:to><p:strVal val="visible"/></p:to></p:set>
<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="9" dur="{fly_ms}" fill="hold"/>
<p:tgtEl><p:spTgt spid="{grp_id}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr>
<p:tavLst><p:tav tm="0"><p:val><p:strVal val="1+#ppt_w/2"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav></p:tavLst></p:anim>
<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="10" dur="{fly_ms}" fill="hold"/>
<p:tgtEl><p:spTgt spid="{grp_id}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>
<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>
</p:childTnLst></p:cTn></p:par>

</p:childTnLst></p:cTn></p:par>
</p:childTnLst></p:cTn></p:par>
</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>
</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"""
    slide._element.append(etree.fromstring(xml))


def section_paths(a, i):
    """第 i 节的所有形状路径。"""
    num, panel = split_by_digit(digit_path(a.digit_font, a.numbers[i]))
    titles = [ln for ln in a.titles[i].split("|") if ln]
    bodies = [ln for ln in a.bodies[i].split("|") if ln] if a.bodies[i] else []
    # 标题和小字整体上下居中
    cap = TITLE_EM * 0.7
    h = cap + (len(titles) - 1) * TITLE_STEP + (BODY_GAP + (len(bodies) - 1) * BODY_STEP + BODY_EM * 0.4 if bodies else 0)
    top = 540 - h / 2
    texts = []
    for j, ln in enumerate(titles):
        texts.append(line_path(a.title_font, ln, TITLE_EM, TITLE_X, top + cap + j * TITLE_STEP))
    y = top + cap + (len(titles) - 1) * TITLE_STEP + BODY_GAP
    for j, ln in enumerate(bodies):
        texts.append(line_path(a.body_font, ln, BODY_EM, TITLE_X, y + j * BODY_STEP))
    return num, panel, texts


def build(a):
    prs = Presentation()
    prs.slide_width, prs.slide_height = SW, SH
    info = []
    for i, img in enumerate(a.images):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        pic = slide.shapes.add_picture(str(img), 0, 0, Emu(SW), Emu(SH))
        pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = cover_crop(img)
        pic._element.nvPicPr.cNvPr.set("name", "底图")
        num, panel, texts = section_paths(a, i)
        color = a.colors[i]
        items = [shape_xml("数字", num, fill("FFFFFF", 80)), shape_xml("色块", panel, fill(color, 80))]
        items += [shape_xml(f"文字{j + 1}", t, fill("FFFFFF")) for j, t in enumerate(texts)]
        gid = add_group(slide, "数字和色块", items)
        add_timing(slide, pic.shape_id, gid, a.zoom_ms, a.fly_ms, a.fly_decel)
        info.append((num, panel, texts, color))
    prs.save(a.out)
    print("已生成：", a.out)
    return info


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
<title>数字过渡页预览</title>
<style>
:root{--bg:#0b0b0c;--fg:#e9e9ea;--muted:#8a8a90;--btn:#1d1d20;--line:#2a2a2e}
@media (prefers-color-scheme: light){:root:not([data-theme="dark"]){--bg:#f4f4f5;--fg:#18181b;--muted:#6b6b72;--btn:#fff;--line:#d9d9de}}
:root[data-theme="dark"]{--bg:#0b0b0c;--fg:#e9e9ea;--muted:#8a8a90;--btn:#1d1d20;--line:#2a2a2e}
html,body{margin:0;background:var(--bg);color:var(--fg);font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif}
.wrap{max-width:1200px;margin:0 auto;padding:16px}
.stage{position:relative;width:100%;aspect-ratio:16/9;background:#000;overflow:hidden;border-radius:6px;cursor:pointer}
.stage svg{position:absolute;inset:0;width:100%;height:100%;display:block}
.bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:12px}
button{background:var(--btn);color:var(--fg);border:1px solid var(--line);border-radius:8px;padding:8px 14px;font-size:14px;cursor:pointer}
.hint{color:var(--muted);font-size:13px}.dots{margin-left:auto;font-size:13px;color:var(--muted)}
</style></head><body><div class="wrap">
<div class="stage" id="stage"><svg viewBox="0 0 1920 1080">__SLIDES__</svg></div>
<div class="bar"><button id="prev">上一页</button><button id="next">下一页</button>
<span class="hint">点画面或按 → 翻页，和 PPT 放映一样</span><span class="dots" id="pg"></span></div>
</div>
<script>
const N=__N__,ZD=__ZD__,FD=__FD__,DEC=__DEC__,Z=__Z__,GL=__GL__;
const $=id=>document.getElementById(id);let cur=0,t0=0,raf;
function prog(u,a){u=Math.min(1,Math.max(0,u));if(a<=0)return u;const v=1/(1-a/2);return u<1-a?v*u:v*(1-a)+v*(u-(1-a))-v*(u-(1-a))**2/(2*a)}
function draw(i,ms){for(let k=0;k<N;k++)$("s"+k).style.display=k===i?"":"none";
 const z=Z+(1-Z)*prog(ms/ZD,1);$("p"+i).setAttribute("transform",`translate(960 540) scale(${z}) translate(-960 -540)`);
 const f=prog(ms/FD,DEC);$("g"+i).setAttribute("transform",`translate(${(1920-GL[i])*(1-f)} 0)`);$("pg").textContent=(i+1)+" / "+N}
function go(i){i=Math.max(0,Math.min(N-1,i));cur=i;cancelAnimationFrame(raf);t0=performance.now();
 (function f(now){const ms=now-t0;draw(cur,ms);if(ms<Math.max(ZD,FD))raf=requestAnimationFrame(f)})(t0)}
$("stage").onclick=()=>go(cur>=N-1?0:cur+1);$("next").onclick=()=>go(cur+1);$("prev").onclick=()=>go(cur-1);
addEventListener("keydown",e=>{if(["ArrowRight"," ","Enter","PageDown"].includes(e.key)){e.preventDefault();go(cur+1)}if(["ArrowLeft","PageUp"].includes(e.key))go(cur-1)});
go(0);window.preview={at:(i,s)=>{cancelAnimationFrame(raf);cur=i;draw(i,s*1000)}};
</script></body></html>"""


def write_preview(a, info):
    parts, lefts = [], []
    for i, (img, (num, panel, texts, color)) in enumerate(zip(a.images, info)):
        im = Image.open(img).convert("RGB")
        l, t, r, b = cover_crop(img)
        W, H = im.size
        im = im.crop((int(l * W), int(t * H), int(W - r * W), int(H - b * H))).resize((1920, 1080))
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=85)
        uri = "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
        g = [f'<path d="{svg_d(num)}" fill="#fff" fill-opacity=".8"/>',
             f'<path d="{svg_d(panel)}" fill="#{color}" fill-opacity=".8"/>']
        g += [f'<path d="{svg_d(t)}" fill="#fff"/>' for t in texts]
        lefts.append(min(p.bounds[0] for p in [num, panel] + texts) / PX)
        parts.append(f'<g id="s{i}"><image id="p{i}" href="{uri}" width="1920" height="1080" preserveAspectRatio="none"/>'
                     f'<g id="g{i}">{"".join(g)}</g></g>')
    html = (PREVIEW.replace("__SLIDES__", "".join(parts)).replace("__N__", str(len(info)))
            .replace("__ZD__", str(a.zoom_ms)).replace("__FD__", str(a.fly_ms)).replace("__DEC__", str(a.fly_decel))
            .replace("__Z__", str(START_ZOOM)).replace("__GL__", "[" + ",".join(f"{v:.1f}" for v in lefts) + "]"))
    Path(a.preview).write_text(html, encoding="utf-8")
    print("已生成网页预览：", a.preview)


def average_color(img):
    """从图片右半边取一个色块颜色（深一点，文字更清楚）。"""
    im = Image.open(img).convert("RGB").resize((64, 36))
    px = [im.getpixel((x, y)) for x in range(32, 64) for y in range(36)]
    r, g, b = (sum(c[k] for c in px) / len(px) for k in range(3))
    return "%02X%02X%02X" % tuple(int(v * 0.85) for v in (r, g, b))


DEFAULT_TITLES = ["KNOW|YOURSELF", "EXPLORE|OPTIONS", "MAKE|A PLAN"]
DEFAULT_BODIES = ["认识自己，找准方向", "走进行业，看清选择", "拆解目标，持续行动"]


def main():
    ap = argparse.ArgumentParser(description="数字过渡页：超大数字切开画面，色块和标题从右飞入")
    default_imgs = sorted(str(p) for p in HERE.glob("example*.jpg"))
    ap.add_argument("--images", nargs="+", default=default_imgs, help="每节一张底图，按顺序；有几张就生成几页")
    ap.add_argument("--numbers", nargs="+", default=None, help="每页的数字，默认 1、2、3……")
    ap.add_argument("--titles", nargs="+", default=None, help="每页的标题，用 | 分行，例如 \"KNOW|YOURSELF\"")
    ap.add_argument("--bodies", nargs="+", default=None, help="每页标题下的小字，用 | 分行；不要就写 \"\"")
    ap.add_argument("--colors", nargs="+", default=None, help="每页色块颜色（十六进制），默认从图里取")
    ap.add_argument("--digit-font", default=str(FONTS / "NotoSerifDisplay-Medium.ttf"), help="数字字体文件")
    ap.add_argument("--title-font", default=str(FONTS / "Cinzel-Regular.ttf"), help="标题字体文件")
    ap.add_argument("--body-font", default=str(FONTS / "NotoSansSC-Medium.ttf"), help="小字字体文件（写中文要用中文字体）")
    ap.add_argument("--zoom-ms", type=int, default=2000, help="底图缩回的时长（毫秒）")
    ap.add_argument("--fly-ms", type=int, default=1700, help="色块飞入的时长（毫秒）")
    ap.add_argument("--fly-decel", type=float, default=0.9, help="飞入减速比例 0～1")
    ap.add_argument("--out", default="数字过渡页.pptx", help="输出的 PPT 文件名")
    ap.add_argument("--preview", default=None, help="同时输出网页预览，例如 --preview 预览.html")
    a = ap.parse_args()
    n = len(a.images)
    if n == 0:
        raise SystemExit("至少需要 1 张底图")
    a.numbers = a.numbers or [str(i + 1) for i in range(n)]
    a.titles = a.titles or [DEFAULT_TITLES[i % 3] for i in range(n)]
    a.bodies = a.bodies if a.bodies is not None else [DEFAULT_BODIES[i % 3] for i in range(n)]
    a.colors = [c.lstrip("#").upper() for c in a.colors] if a.colors else [average_color(p) for p in a.images]
    for name in ("numbers", "titles", "bodies", "colors"):
        v = getattr(a, name)
        if len(v) < n:
            setattr(a, name, v + [v[-1]] * (n - len(v)))
    info = build(a)
    if a.preview:
        write_preview(a, info)


if __name__ == "__main__":
    main()
