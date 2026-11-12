"""唱片机结束页 · Record Ending

生成 1 页结束页：
  整页铺一张底图并稍微压暗，左边一个圆环透出没压暗的原图，像一张唱片，一直匀速转；
  右上角一支唱臂搭在唱片上；右边的结尾标题和几段小字像片尾字幕一样，从下往上慢慢滚动。

动画用的是 PowerPoint / WPS 自带的「陀螺旋」和「直线路径」，进入这一页就自动播放、一直循环。
所有文字都会转成形状，换一台没装字体的电脑打开，效果也不会变。

用法：
  python generate.py --title "谢谢聆听" --image 我的图.jpg
  python generate.py --help      查看全部参数
"""
import argparse
import base64
import copy
import io
import json
import math
from pathlib import Path

import pathops
from lxml import etree
from pptx import Presentation
from pptx.util import Emu

HERE = Path(__file__).resolve().parent
FONTS = HERE.parent.parent / "fonts"

# ---------------- 通用：场景 → .pptx / 网页预览 ----------------
# 坐标统一用 1920×1080 像素，写 pptx 时换算成 EMU。
W, H = 1920, 1080
EMU = 6350
NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
A = '{%s}' % NS['a']
P = '{%s}' % NS['p']
K = 0.5522847498


def E(v):
    return int(round(v * EMU))


# ---------------- 几何 ----------------
def circle_path(cx, cy, r, ccw=False):
    p = pathops.Path(); pen = p.getPen()
    pts = [(cx + r, cy), (cx, cy + r), (cx - r, cy), (cx, cy - r)]
    if ccw: pts = pts[::-1]
    pen.moveTo(pts[0])
    for i in range(4):
        a, b = pts[i], pts[(i + 1) % 4]
        # 控制点：沿切线方向
        ax, ay = a[0] - cx, a[1] - cy; bx, by = b[0] - cx, b[1] - cy
        sgn = 1 if not ccw else -1
        c1 = (a[0] - ay * K * sgn, a[1] + ax * K * sgn)
        c2 = (b[0] + by * K * sgn, b[1] - bx * K * sgn)
        pen.curveTo(c1, c2, b)
    pen.closePath()
    return p


def rect_path(x0, y0, x1, y1):
    p = pathops.Path(); pen = p.getPen()
    pen.moveTo((x0, y0)); pen.lineTo((x1, y0)); pen.lineTo((x1, y1)); pen.lineTo((x0, y1)); pen.closePath()
    return p


def ring_path(size, inner_ratio=0.5, gap_ratio=0.045):
    """size×size 方框里的四段圆环（圆环减十字缝），坐标 0..size。"""
    R = size / 2; c = R
    ring = pathops.op(circle_path(c, c, R), circle_path(c, c, R * inner_ratio), pathops.PathOp.DIFFERENCE)
    g = R * gap_ratio / 2
    bars = pathops.op(rect_path(c - g, -1, c + g, size + 1), rect_path(-1, c - g, size + 1, c + g), pathops.PathOp.UNION)
    return pathops.op(ring, bars, pathops.PathOp.DIFFERENCE)


def path_segments(path, size_w, size_h):
    """返回归一化 (0..1) 的段列表，供 pptx 与网页共用。"""
    segs = []; start = None
    for verb, pts in path.segments:
        if verb == 'qCurveTo' and pts[-1] is None:
            offs = pts[:-1]
            m = ((offs[-1][0] + offs[0][0]) / 2, (offs[-1][1] + offs[0][1]) / 2)
            segs.append(['M', [m[0] / size_w, m[1] / size_h]]); start = m
            pts = tuple(offs) + (m,)
        if verb == 'moveTo': start = pts[0]; segs.append(['M', [pts[0][0] / size_w, pts[0][1] / size_h]])
        elif verb == 'lineTo': segs.append(['L', [pts[0][0] / size_w, pts[0][1] / size_h]])
        elif verb == 'curveTo': segs.append(['C'] + [[q[0] / size_w, q[1] / size_h] for q in pts])
        elif verb == 'qCurveTo':
            q = list(pts); offs, end = q[:-1], q[-1]
            for i, c in enumerate(offs):
                nxt = ((c[0] + offs[i + 1][0]) / 2, (c[1] + offs[i + 1][1]) / 2) if i < len(offs) - 1 else end
                segs.append(['Q', [c[0] / size_w, c[1] / size_h], [nxt[0] / size_w, nxt[1] / size_h]])
        elif verb in ('closePath', 'endPath'): segs.append(['Z'])
    return segs


def custgeom(segs):
    S = 100000
    geom = etree.Element(A + 'custGeom')
    for t in ('avLst', 'gdLst', 'ahLst', 'cxnLst'):
        etree.SubElement(geom, A + t)
    etree.SubElement(geom, A + 'rect', l='l', t='t', r='r', b='b')
    pl = etree.SubElement(geom, A + 'pathLst')
    pe = etree.SubElement(pl, A + 'path', w=str(S), h=str(S))
    def pt(parent, q):
        etree.SubElement(parent, A + 'pt', x=str(int(round(q[0] * S))), y=str(int(round(q[1] * S))))
    for s in segs:
        if s[0] == 'M': e = etree.SubElement(pe, A + 'moveTo'); pt(e, s[1])
        elif s[0] == 'L': e = etree.SubElement(pe, A + 'lnTo'); pt(e, s[1])
        elif s[0] == 'C':
            e = etree.SubElement(pe, A + 'cubicBezTo'); [pt(e, q) for q in s[1:]]
        elif s[0] == 'Q':
            e = etree.SubElement(pe, A + 'quadBezTo'); [pt(e, q) for q in s[1:]]
        elif s[0] == 'Z': etree.SubElement(pe, A + 'close')
    return geom


# ---------------- 填充 / 效果 XML ----------------
def alpha_xml(op):
    return '' if op is None or op >= 1 else f'<a:alpha val="{int(op * 100000)}"/>'


def fill_xml(f):
    if f is None: return '<a:noFill/>'
    t = f['type']
    if t == 'solid':
        return f'<a:solidFill><a:srgbClr val="{f["color"]}">{alpha_xml(f.get("alpha"))}</a:srgbClr></a:solidFill>'
    if t == 'grad':
        gs = ''.join(f'<a:gs pos="{int(p * 100000)}"><a:srgbClr val="{c}">{alpha_xml(al)}</a:srgbClr></a:gs>' for p, c, al in f['stops'])
        return f'<a:gradFill rotWithShape="1"><a:gsLst>{gs}</a:gsLst><a:lin ang="{int(f.get("angle", 90) * 60000)}" scaled="0"/></a:gradFill>'
    if t == 'bg':
        return ''  # useBgFill
    raise ValueError(t)


def shadow_xml(s):
    if not s: return '<a:effectLst/>'
    return (f'<a:effectLst><a:outerShdw blurRad="{int(s["blur"] * 12700)}" dist="{int(s.get("dist", 0) * 12700)}" '
            f'dir="{int(s.get("dir", 0) * 60000)}" sx="{int(s.get("size", 1) * 100000)}" sy="{int(s.get("size", 1) * 100000)}" '
            f'algn="ctr" rotWithShape="0"><a:srgbClr val="{s.get("color", "000000")}"><a:alpha val="{int(s["alpha"] * 100000)}"/></a:srgbClr></a:outerShdw></a:effectLst>')


def line_xml(l):
    if not l: return '<a:ln><a:noFill/></a:ln>'
    return f'<a:ln w="{int(l["w"] * 12700)}"><a:solidFill><a:srgbClr val="{l["color"]}">{alpha_xml(l.get("alpha"))}</a:srgbClr></a:solidFill></a:ln>'


# ---------------- 写 pptx ----------------
_id = [10]


def nid():
    _id[0] += 1; return _id[0]


def xfrm_xml(o, tag='a:xfrm', group_child=None):
    rot = int(round((o.get('rot', 0) % 360) * 60000))
    ra = f' rot="{rot}"' if rot else ''
    return f'<{tag}{ra}><a:off x="{E(o["x"])}" y="{E(o["y"])}"/><a:ext cx="{E(o["w"])}" cy="{E(o["h"])}"/></{tag}>'


def shape_el(o):
    kind = o['kind']
    ub = ' useBgFill="1"' if o.get('fill') and o['fill']['type'] == 'bg' else ''
    if kind in ('rect', 'ellipse', 'flowChartSummingJunction', 'roundRect', 'path'):
        prst = {'rect': 'rect', 'ellipse': 'ellipse', 'flowChartSummingJunction': 'flowChartSummingJunction', 'roundRect': 'roundRect'}.get(kind)
        geom = f'<a:prstGeom prst="{prst}"><a:avLst>{o.get("av", "")}</a:avLst></a:prstGeom>' if prst else '<GEOM/>'
        txt = ''
        if o.get('text') is not None:
            txt = text_body(o)
        else:
            txt = '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>'
        xml = (f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"{ub}><p:nvSpPr><p:cNvPr id="{nid()}" name="{o["name"]}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
               f'<p:spPr>{xfrm_xml(o)}{geom}{fill_xml(o.get("fill"))}{line_xml(o.get("line"))}{shadow_xml(o.get("shadow"))}</p:spPr>{txt}</p:sp>')
        el = etree.fromstring(xml)
        if kind == 'path':
            g = el.find('.//GEOM'); g.getparent().replace(g, custgeom(o['segs']))
        return el
    if kind == 'text':
        xml = (f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvSpPr><p:cNvPr id="{nid()}" name="{o["name"]}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
               f'<p:spPr>{xfrm_xml(o)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>{shadow_xml(o.get("shadow"))}</p:spPr>{text_body(o)}</p:sp>')
        return etree.fromstring(xml)
    if kind == 'group':
        kids = o['children']
        xml = (f'<p:grpSp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvGrpSpPr><p:cNvPr id="{nid()}" name="{o["name"]}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
               f'<p:grpSpPr><a:xfrm{(" rot=\"%d\"" % int(round((o.get("rot", 0) % 360) * 60000))) if o.get("rot") else ""}><a:off x="{E(o["x"])}" y="{E(o["y"])}"/><a:ext cx="{E(o["w"])}" cy="{E(o["h"])}"/>'
               f'<a:chOff x="{E(o["x"])}" y="{E(o["y"])}"/><a:chExt cx="{E(o["w"])}" cy="{E(o["h"])}"/></a:xfrm></p:grpSpPr></p:grpSp>')
        el = etree.fromstring(xml)
        for k in kids: el.append(shape_el(k))
        return el
    raise ValueError(kind)


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def text_body(o):
    t = o['textstyle']
    anchor = {'top': 't', 'middle': 'ctr', 'bottom': 'b'}[t.get('valign', 'middle')]
    algn = {'left': 'l', 'center': 'ctr', 'right': 'r', 'justify': 'just'}[t.get('align', 'left')]
    sz = int(round(t['size'] * 100))
    spc = f' spc="{int(round(t.get("spacing", 0) * 100))}"' if t.get('spacing') else ''
    b = ' b="1"' if t.get('bold') else ''
    font = t['font']; ea = t.get('ea', font)
    if t.get('fill'):
        tf = fill_xml(t['fill'])
    else:
        tf = f'<a:solidFill><a:srgbClr val="{t.get("color", "FFFFFF")}">{alpha_xml(t.get("alpha"))}</a:srgbClr></a:solidFill>'
    eff = shadow_xml(t.get('shadow')) if t.get('shadow') else ''
    if t.get('outline'):
        tf = f'<a:ln w="{int(t["outline"]["w"] * 12700)}"><a:solidFill><a:srgbClr val="{t["outline"]["color"]}"/></a:solidFill><a:round/></a:ln>' + tf
    ln = t.get('line_spacing')
    lnx = f'<a:lnSpc><a:spcPct val="{int(ln * 100000)}"/></a:lnSpc>' if ln else ''
    paras = ''
    for line in o['text'].split('\n'):
        bullet = '<a:buFont typeface="Arial"/><a:buChar char="•"/>' if t.get('bullet') else '<a:buNone/>'
        ind = f' marL="{E(t["bullet_indent"])}" indent="-{E(t["bullet_indent"])}"' if t.get('bullet') else ''
        paras += (f'<a:p><a:pPr algn="{algn}"{ind}>{lnx}{bullet}</a:pPr><a:r><a:rPr lang="zh-CN" sz="{sz}"{b}{spc} dirty="0">{tf}{eff}'
                  f'<a:latin typeface="{font}"/><a:ea typeface="{ea}"/><a:cs typeface="{font}"/></a:rPr><a:t>{esc(line)}</a:t></a:r></a:p>')
    return (f'<p:txBody><a:bodyPr wrap="{"square" if t.get("wrap") else "none"}" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}" rtlCol="0"><a:noAutofit/></a:bodyPr>'
            f'<a:lstStyle/>{paras}</p:txBody>')


def set_bg_picture(slide, rId):
    cSld = slide._element.find('p:cSld', NS)
    bg = etree.fromstring(f'<p:bg xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}" xmlns:r="{NS["r"]}"><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="{rId}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>')
    cSld.insert(0, bg)


def add_morph(slide, dur_ms):
    xml = f'''<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">
    <p:transition xmlns:p="{NS['p']}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="slow" p14:dur="{dur_ms}"><p159:morph option="byObject"/></p:transition>
  </mc:Choice>
  <mc:Fallback><p:transition xmlns:p="{NS['p']}" spd="slow"><p:fade/></p:transition></mc:Fallback>
</mc:AlternateContent>'''
    el = etree.fromstring(xml)
    root = slide._element
    clr = root.find('p:clrMapOvr', NS)
    (clr if clr is not None else root.find('p:cSld', NS)).addnext(el)


def image_part_rid(slide, path):
    _, rId = slide.part.get_or_add_image_part(path)
    return rId


def build_pptx(scene, out):
    prs = Presentation(); prs.slide_width, prs.slide_height = E(W), E(H)
    blank = prs.slide_layouts[6]
    for sd in scene['slides']:
        s = prs.slides.add_slide(blank)
        rId = image_part_rid(s, scene['images'][sd['bg']])
        set_bg_picture(s, rId)
        tree = s.shapes._spTree
        objs = sd['objects']
        if scene.get('vector_text'):
            objs = [v for o in objs for v in vectorize(o)]
        for o in objs:
            if o['kind'] == 'picture':
                pic = s.shapes.add_picture(scene['images'][o['image']], Emu(E(o['x'])), Emu(E(o['y'])), Emu(E(o['w'])), Emu(E(o['h'])))
                pic._element.nvPicPr.cNvPr.set('name', o['name'])
                if o.get('alpha') is not None and o['alpha'] < 1:
                    blip = pic._element.find('.//a:blip', NS)
                    etree.SubElement(blip, A + 'alphaModFix', amt=str(int(o['alpha'] * 100000)))
            else:
                tree.append(shape_el(o))
        if sd.get('morph'):
            add_morph(s, sd['morph'])
    prs.save(out)


# ---------------- 网页预览 ----------------
def subset_font(path, text):
    from fontTools import subset
    opts = subset.Options(); opts.layout_features = ['*']
    f = subset.load_font(path, opts)
    s = subset.Subsetter(opts); s.populate(text=text); s.subset(f)
    buf = io.BytesIO(); subset.save_font(f, buf, opts)
    return base64.b64encode(buf.getvalue()).decode()


def img_b64(path, maxw=1920):
    from PIL import Image
    im = Image.open(path).convert('RGB'); im.thumbnail((maxw, maxw))
    buf = io.BytesIO(); im.save(buf, 'JPEG', quality=86)
    return 'data:image/jpeg;base64,' + base64.b64encode(buf.getvalue()).decode()


def all_text(scene):
    out = []
    def walk(o):
        if o.get('text'): out.append(o['text'])
        for k in o.get('children', []): walk(k)
    for sd in scene['slides']:
        for o in sd['objects']: walk(o)
    return ''.join(out) + ' 0123456789'


def build_html(scene, fonts, out, title):
    """fonts: [(字体名, 字体文件, 字重)]，网页里用同样的字体排版。"""
    tpl = PLAYER
    txt = all_text(scene)
    css = ''
    for fam, path, weight in fonts:
        css += f'@font-face{{font-family:"{fam}";src:url(data:font/ttf;base64,{subset_font(path, txt)}) format("truetype");font-weight:{weight};}}\n'
    sc = copy.deepcopy(scene)
    sc['images'] = {k: img_b64(v) for k, v in scene['images'].items()}
    html = tpl.replace('/*FONTS*/', css).replace('__SCENE__', json.dumps(sc, ensure_ascii=False)).replace('__TITLE__', title)
    open(out, 'w', encoding='utf-8').write(html)


# ---------------- 文字转矢量（pptx 里不依赖字体） ----------------
FONT_FILES = {}   # (family, bold) -> ttf 路径
_fcache = {}


def _font(path):
    if path not in _fcache:
        from fontTools.ttLib import TTFont
        f = TTFont(path); _fcache[path] = (f, f.getGlyphSet(), f.getBestCmap(), f['head'].unitsPerEm, f['hmtx'])
    return _fcache[path]


def _pick(fam, bold):
    return FONT_FILES.get((fam, bold)) or FONT_FILES.get((fam, not bold))


def text_outline(o):
    """按网页预览同样的排版规则，把文字排成路径（绝对像素坐标）。"""
    from fontTools.pens.transformPen import TransformPen
    s = o['textstyle']; px = s['size'] * 2; sp = (s.get('spacing') or 0) * 2
    lat = _pick(s['font'], bool(s.get('bold'))); ea = _pick(s.get('ea', s['font']), bool(s.get('bold')))
    lines = o['text'].split('\n'); lh = px * 1.2 * (s.get('line_spacing') or 1); total = lh * len(lines)
    va = s.get('valign', 'middle')
    y0 = o['y'] + (o['h'] - total) / 2 if va == 'middle' else (o['y'] + o['h'] - total if va == 'bottom' else o['y'])
    path = pathops.Path(); pen = path.getPen()
    ind = s.get('bullet_indent', 0) if s.get('bullet') else 0
    for i, ln in enumerate(lines):
        glyphs = []; x = 0
        for ch in ln:
            fp = lat if ord(ch) in _font(lat)[2] else ea
            f, gs, cmap, upm, hmtx = _font(fp)
            g = cmap.get(ord(ch))
            if g is None: x += px * 0.3 + sp; continue
            k = px / upm
            glyphs.append((gs, g, x, k)); x += hmtx[g][0] * k + sp
        w = x - sp
        al = s.get('align', 'left')
        ox = o['x'] + (o['w'] - w) / 2 if al == 'center' else (o['x'] + o['w'] - w if al == 'right' else o['x'] + ind)
        by = y0 + i * lh + lh * 0.5 + px * 0.36
        from fontTools.pens.recordingPen import DecomposingRecordingPen
        for gs, g, gx, k in glyphs:
            rp = DecomposingRecordingPen(gs); gs[g].draw(rp)
            rp.replay(TransformPen(pen, (k, 0, 0, -k, ox + gx, by)))
        if s.get('bullet'):
            c = circle_path(o['x'] + px * 0.3, by - px * 0.33, px * 0.2)
            path = pathops.op(path, c, pathops.PathOp.UNION); pen = path.getPen()
    path = pathops.simplify(path, fix_winding=True)
    rot = o.get('rot', 0) % 360
    if rot:
        from fontTools.pens.transformPen import TransformPen as TP
        cx, cy = o['x'] + o['w'] / 2, o['y'] + o['h'] / 2
        a = math.radians(rot); ca, sa = math.cos(a), math.sin(a)
        q = pathops.Path(); path.draw(TP(q.getPen(), (ca, sa, -sa, ca, cx - ca * cx + sa * cy, cy - sa * cx - ca * cy)))
        path = q
    return path


def vectorize(o):
    """把文字对象换成路径形状；带文字的形状拆成形状 + 文字路径。返回对象列表。"""
    if o['kind'] == 'group':
        g = dict(o); g['children'] = [v for k in o['children'] for v in vectorize(k)]; return [g]
    if o.get('text') is None: return [o]
    out = []
    if o['kind'] != 'text':
        sh = dict(o); sh.pop('text'); sh.pop('textstyle'); out.append(sh)
    s = o['textstyle']
    p = text_outline(o)
    x0, y0, x1, y1 = p.bounds
    fill = s.get('fill') or {'type': 'solid', 'color': s.get('color', 'FFFFFF'), 'alpha': s.get('alpha')}
    v = {'name': o['name'] + ('' if o['kind'] == 'text' else '字'), 'kind': 'path', 'segs': path_segments_off(p, x0, y0, x1 - x0, y1 - y0),
         'x': x0, 'y': y0, 'w': x1 - x0, 'h': y1 - y0, 'fill': fill, 'shadow': s.get('shadow'),
         'line': {'w': s['outline']['w'], 'color': s['outline']['color']} if s.get('outline') else None}
    out.append(v)
    return out


def path_segments_off(path, x0, y0, w, h):
    segs = path_segments(path, 1, 1)
    def n(q): return [(q[0] - x0) / w, (q[1] - y0) / h]
    return [[s[0]] + [n(q) for q in s[1:]] for s in segs]


PLAYER = r'''<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITLE__</title>
<style>
/*FONTS*/
:root{--bg:#0b0b0c;--fg:#e9e9ea;--muted:#8a8a90;--btn:#1d1d20;--line:#2a2a2e}
@media (prefers-color-scheme: light){:root:not([data-theme="dark"]){--bg:#f4f4f5;--fg:#18181b;--muted:#6b6b72;--btn:#fff;--line:#d9d9de}}
:root[data-theme="dark"]{--bg:#0b0b0c;--fg:#e9e9ea;--muted:#8a8a90;--btn:#1d1d20;--line:#2a2a2e}
html,body{margin:0;background:var(--bg);color:var(--fg);font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif}
.wrap{max-width:1200px;margin:0 auto;padding:16px}
.stage{position:relative;width:100%;aspect-ratio:16/9;background:#000;overflow:hidden;border-radius:6px;cursor:pointer}
canvas{position:absolute;inset:0;width:100%;height:100%;display:block}
.bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:12px}
button{background:var(--btn);color:var(--fg);border:1px solid var(--line);border-radius:8px;padding:8px 14px;font-size:14px;cursor:pointer}
.hint{color:var(--muted);font-size:13px}
.dots{margin-left:auto;font-size:13px;color:var(--muted)}
body.cap .wrap{max-width:none;padding:0}body.cap .stage{border-radius:0;width:1920px;height:1080px}body.cap .bar{display:none}
</style></head><body><div class="wrap">
<div class="stage" id="stage"><canvas id="cv" width="1920" height="1080"></canvas></div>
<div class="bar"><button id="prev">上一页</button><button id="next">下一页</button><button id="play">自动播放</button>
<span class="hint">点画面或按 → 翻页，效果和 PPT 放映的平滑切换一致</span><span class="dots" id="pg"></span></div>
</div>
<script>
const SC=__SCENE__;
const W=1920,H=1080,PT=2; // 1pt = 2px
const cv=document.getElementById('cv'),ctx=cv.getContext('2d');
const IM={},VID={};let pending=0;
for(const k in SC.images){pending++;const im=new Image();im.onload=()=>{if(--pending===0)ready()};im.src=SC.images[k];IM[k]=im}
for(const k in (SC.videos||{})){pending++;const v=document.createElement('video');v.muted=true;v.loop=true;v.playsInline=true;v.preload='auto';v.oncanplaythrough=()=>{v.oncanplaythrough=null;if(--pending===0)ready()};v.src=SC.videos[k];VID[k]=v}
const N=SC.slides.length;
function lerp(a,b,t){return a+(b-a)*t}
function ease(t){return (1-Math.cos(Math.PI*t))/2}
function segPath(segs,x,y,w,h){const p=new Path2D();for(const s of segs){const q=v=>[x+v[0]*w,y+v[1]*h];
 if(s[0]==='M')p.moveTo(...q(s[1]));else if(s[0]==='L')p.lineTo(...q(s[1]));else if(s[0]==='C')p.bezierCurveTo(...q(s[1]),...q(s[2]),...q(s[3]));else if(s[0]==='Q')p.quadraticCurveTo(...q(s[1]),...q(s[2]));else p.closePath()}return p}
function shapePath(o){const {x,y,w,h}=o;const p=new Path2D();
 if(o.kind==='path')return segPath(o.segs,x,y,w,h);
 if(o.kind==='ellipse'||o.kind==='flowChartSummingJunction'){p.ellipse(x+w/2,y+h/2,w/2,h/2,0,0,Math.PI*2);return p}
 if(o.kind==='roundRect'){const r=Math.min(w,h)/2;p.roundRect(x,y,w,h,r);return p}
 p.rect(x,y,w,h);return p}
let BG=[]; // [{img,alpha}]
function drawBg(alphaMul){for(const b of BG){ctx.globalAlpha=b.alpha*(alphaMul??1);ctx.drawImage(VID[b.img]||IM[b.img],0,0,W,H)}ctx.globalAlpha=1}
function fillStyle(f,o){if(!f)return null;if(f.type==='solid')return rgba(f.color,f.alpha??1);
 if(f.type==='grad'){const a=(f.angle??90)*Math.PI/180;const cx=o.x+o.w/2,cy=o.y+o.h/2;const L=Math.abs(o.w*Math.cos(a))+Math.abs(o.h*Math.sin(a));
  const g=ctx.createLinearGradient(cx-Math.cos(a)*L/2,cy-Math.sin(a)*L/2,cx+Math.cos(a)*L/2,cy+Math.sin(a)*L/2);for(const [p,c,al] of f.stops)g.addColorStop(p,rgba(c,al??1));return g}return null}
function rgba(hex,a){const n=parseInt(hex,16);return `rgba(${n>>16&255},${n>>8&255},${n&255},${a})`}
function withXf(o,fn){ctx.save();if(o.rot){const cx=o.x+o.w/2,cy=o.y+o.h/2;ctx.translate(cx,cy);ctx.rotate(o.rot*Math.PI/180);ctx.translate(-cx,-cy)}fn();ctx.restore()}
function drawShadow(path,o,s,op){ // 阴影不随形状旋转：在未旋转坐标中偏移
 const d=(s.dist||0)*PT,dir=(s.dir||0)*Math.PI/180;const sz=s.size||1;
 ctx.save();ctx.setTransform(1,0,0,1,0,0);
 const m=ctx.__m;ctx.shadowColor=rgba(s.color||'000000',s.alpha*op);ctx.shadowBlur=s.blur*PT*0.9;
 ctx.shadowOffsetX=d*Math.cos(dir)+20000;ctx.shadowOffsetY=d*Math.sin(dir);
 ctx.translate(-20000,0);
 const cx=o.__cx,cy=o.__cy;ctx.translate(cx,cy);ctx.scale(sz,sz);ctx.translate(-cx,-cy);
 ctx.transform(m.a,m.b,m.c,m.d,m.e,m.f);ctx.fillStyle='#000';ctx.fill(path);ctx.restore()}
let CLK=0;
function animOf(o){const a=o.anim;if(!a)return o;const r=Object.assign({},o);
 if(a.spin){r.rot=(o.rot||0)+((CLK/a.spin)%1)*360;r.__spin=true}
 if(a.move){const p=a.loop?((CLK/a.dur)%1):Math.min(1,CLK/a.dur);const dy=a.move[0]+(a.move[1]-a.move[0])*p;r.__dy=dy}
 if(a.fade){const p=Math.max(0,Math.min(1,(CLK-(a.delay||0))/a.fade));r.opacity=(o.opacity??1)*p}
 return r}
function drawObj(o,op){o=animOf(o);if(o.__dy){ctx.save();ctx.translate(0,o.__dy);const q=Object.assign({},o);delete q.__dy;delete q.anim;drawObj(q,op);ctx.restore();return}
 op*=(o.opacity??1);if(op<=0.002)return;
 if(o.kind==='group'){withXf(o,()=>{for(const k of o.children)drawObj(k,op)});return}
 if(o.kind==='picture'){withXf(o,()=>{ctx.globalAlpha=op*(o.alpha??1);ctx.drawImage(IM[o.image],o.x,o.y,o.w,o.h);ctx.globalAlpha=1});return}
 if(o.kind==='text'){withXf(o,()=>{if(o.__xfade){const [ta,tb,t]=o.__xfade;drawText(Object.assign({},o,{text:ta}),op*(1-t));drawText(Object.assign({},o,{text:tb}),op*t)}else drawText(o,op)});return}
 withXf(o,()=>{const p=shapePath(o);const t=ctx.getTransform();ctx.__m=t;
  const c=t.transformPoint(new DOMPoint(o.x+o.w/2,o.y+o.h/2));o.__cx=c.x;o.__cy=c.y;
  if(o.shadow)drawShadow(p,o,o.shadow,op);
  if(o.fill){if(o.fill.type==='bg'){ctx.save();ctx.clip(p);if(!o.__spin)ctx.setTransform(1,0,0,1,0,0);drawBg(op);ctx.restore()}
   else{ctx.globalAlpha=op;ctx.fillStyle=fillStyle(o.fill,o);ctx.fill(p);ctx.globalAlpha=1}}
  if(o.kind==='flowChartSummingJunction'&&o.line){ctx.save();ctx.clip(p);const cx=o.x+o.w/2,cy=o.y+o.h/2,r=o.w/2*Math.SQRT1_2;ctx.beginPath();ctx.moveTo(cx-r,cy-r);ctx.lineTo(cx+r,cy+r);ctx.moveTo(cx+r,cy-r);ctx.lineTo(cx-r,cy+r);
   ctx.strokeStyle=rgba(o.line.color,(o.line.alpha??1)*op);ctx.lineWidth=o.line.w*PT;ctx.stroke();ctx.restore()}
  if(o.line){ctx.strokeStyle=rgba(o.line.color,(o.line.alpha??1)*op);ctx.lineWidth=o.line.w*PT;ctx.stroke(p)}
  if(o.text)drawText(o,op)})}
function drawText(o,op){const s=o.textstyle;const px=s.size*PT;const fam=`"${s.ea||s.font}","${s.font}",sans-serif`;
 ctx.font=`${s.bold?'700':'400'} ${px}px ${s.font==s.ea?`"${s.font}"`:`"${s.font}","${s.ea}"`},sans-serif`;
 ctx.letterSpacing=((s.spacing||0)*PT)+'px';ctx.textBaseline='alphabetic';
 const lines=o.text.split('\n');const lh=px*1.2*(s.line_spacing||1);const total=lh*lines.length;
 let y0=o.y;if((s.valign||'middle')==='middle')y0=o.y+(o.h-total)/2;else if(s.valign==='bottom')y0=o.y+o.h-total;
 const fs=s.fill?fillStyle(s.fill,o):rgba(s.color||'FFFFFF',s.alpha??1);
 ctx.globalAlpha=op;
 if(s.shadow){ctx.shadowColor=rgba(s.shadow.color||'000000',s.shadow.alpha);ctx.shadowBlur=s.shadow.blur*PT;ctx.shadowOffsetX=(s.shadow.dist||0)*PT*Math.cos((s.shadow.dir||0)*Math.PI/180);ctx.shadowOffsetY=(s.shadow.dist||0)*PT*Math.sin((s.shadow.dir||0)*Math.PI/180)}
 lines.forEach((ln,i)=>{let x=o.x;const ind=s.bullet?s.bullet_indent:0;const w=ctx.measureText(ln).width-(s.spacing||0)*PT;
  const al=s.align||'left';if(al==='center')x=o.x+(o.w-w)/2;else if(al==='right')x=o.x+o.w-w;else x=o.x+ind;
  const by=y0+i*lh+lh*0.5+px*0.36;ctx.fillStyle=fs;
  if(s.bullet){ctx.beginPath();ctx.arc(o.x+px*0.3,by-px*0.33,px*0.2,0,Math.PI*2);ctx.fill()}
  if(s.outline){ctx.save();ctx.shadowColor='transparent';ctx.strokeStyle=rgba(s.outline.color,1);ctx.lineWidth=s.outline.w*PT;ctx.lineJoin='round';ctx.strokeText(ln,x,by);ctx.restore()}ctx.fillText(ln,x,by)});
 ctx.shadowColor='transparent';ctx.globalAlpha=1;ctx.letterSpacing='0px'}
// ---- 平滑切换插值 ----
const NUM=['x','y','w','h','rot','opacity'];
function mixObj(a,b,t){const o=Object.assign({},b);for(const k of NUM){const va=a[k]??(k==='opacity'?1:0),vb=b[k]??(k==='opacity'?1:0);o[k]=lerp(va,vb,t)}
 if(a.kind==='group'&&b.kind==='group')o.children=mixList(a.children,b.children,t);
 if(a.textstyle&&b.textstyle){o.textstyle=Object.assign({},b.textstyle,{color:lerpHex(a.textstyle.color||'FFFFFF',b.textstyle.color||'FFFFFF',t),alpha:lerp(a.textstyle.alpha??1,b.textstyle.alpha??1,t)});
  if(a.text!==b.text){o.__xfade=[a.text,b.text,t]}}
 if(a.fill&&b.fill&&a.fill.type==='solid'&&b.fill.type==='solid')o.fill={type:'solid',color:lerpHex(a.fill.color,b.fill.color,t),alpha:lerp(a.fill.alpha??1,b.fill.alpha??1,t)};
 return o}
function lerpHex(a,b,t){const x=parseInt(a,16),y=parseInt(b,16);const c=[16,8,0].map(s=>Math.round(lerp(x>>s&255,y>>s&255,t)));return c.map(v=>v.toString(16).padStart(2,'0')).join('')}
function mixList(A,B,t){const out=[];const bn=new Set(B.map(o=>o.name)),an=new Map(A.map(o=>[o.name,o]));
 for(const o of A)if(!bn.has(o.name))out.push(Object.assign({},o,{opacity:(o.opacity??1)*(1-t)}));
 for(const o of B){if(an.has(o.name))out.push(mixObj(an.get(o.name),o,t));else out.push(Object.assign({},o,{opacity:(o.opacity??1)*t}))}return out}
function render(i,t){const S=SC.slides;ctx.clearRect(0,0,W,H);
 if(t===undefined||t>=1||i===0&&t===undefined){BG=[{img:S[i].bg,alpha:1}];drawBg();for(const o of S[i].objects)drawObj(o,1);return}
 const a=S[i-1],b=S[i];const e=ease(t);BG=a.bg===b.bg?[{img:b.bg,alpha:1}]:[{img:a.bg,alpha:1},{img:b.bg,alpha:e}];drawBg();
 for(const o of mixList(a.objects,b.objects,e))drawObj(o,1)}
let cur=0,raf=null,live=null;
function playVids(n){for(const k in VID){const v=VID[k];if(SC.slides[n].bg===k){if(v.paused){v.currentTime=0;v.play().catch(()=>{})}}else v.pause()}}
function go(n,instant){n=Math.max(0,Math.min(N-1,n));cancelAnimationFrame(raf);cancelAnimationFrame(live);const from=cur;cur=n;document.getElementById('pg').textContent=(n+1)+' / '+N;
 playVids(n);const s0=performance.now();CLK=0;
 const tick=()=>{if(window.__freeze)return;CLK=performance.now()-s0;render(n);live=requestAnimationFrame(tick)};
 if(instant||n!==from+1||!SC.slides[n].morph){tick();return}const D=SC.slides[n].morph;
 (function f(now){const t=Math.min(1,(now-s0)/D);CLK=now-s0;render(n,t);if(t<1)raf=requestAnimationFrame(f);else tick()})(s0)}
window.__setClock=(ms)=>{CLK=ms};
window.__seekVideo=async(k,sec)=>{const v=VID[k];v.pause();v.currentTime=sec%v.duration;await new Promise(r=>v.onseeked=r);return sec}
function ready(){const fams=new Set();const walk=o=>{if(o.textstyle){fams.add(o.textstyle.font);fams.add(o.textstyle.ea||o.textstyle.font)}(o.children||[]).forEach(walk)};SC.slides.forEach(s=>s.objects.forEach(walk));
 Promise.all([...fams].flatMap(f=>['400','700'].map(w=>document.fonts.load(`${w} 40px "${f}"`,'汉字AB01')))).then(()=>{go(0,true);window.__ready=true})}
document.getElementById('stage').onclick=()=>go(cur+1);
document.getElementById('next').onclick=()=>go(cur+1);
document.getElementById('prev').onclick=()=>go(cur-1,true);
document.getElementById('play').onclick=()=>{go(0,true);let k=0,tt=600;for(let i=1;i<N;i++){setTimeout(()=>go(i),tt);tt+=(SC.slides[i].morph||1500)+1400}};
addEventListener('keydown',e=>{if(['ArrowRight','ArrowDown',' ','PageDown','Enter'].includes(e.key)){e.preventDefault();go(cur+1)}if(['ArrowLeft','ArrowUp','PageUp'].includes(e.key))go(cur-1,true)});
window.__render=render;
</script></body></html>
'''


# ---------------- 页面内动画（PPT 动画窗格里的效果） ----------------
def shape_ids(slide):
    """名字 → 形状 id（cNvPr）。"""
    out = {}
    for el in slide._element.iter('{%s}cNvPr' % NS['p']):
        out[el.get('name')] = el.get('id')
    return out


def anim_xml(effects):
    """effects: [(kind, spid, opts)]，全部在进入页面时自动开始，同时播放。
    kind: spin（陀螺旋，opts: dur 毫秒, repeat）、move（直线路径，opts: dy0, dy1 占页面高度比例, dur, repeat）、
          fade（淡化进入，opts: dur, delay）。"""
    nid_ = [4]

    def n():
        nid_[0] += 1
        return nid_[0]
    pars = []
    first = True
    for kind, spid, o in effects:
        node = 'afterEffect' if first else 'withEffect'
        first = False
        rep = ' repeatCount="indefinite"' if o.get('repeat') else ''
        delay = int(o.get('delay', 0))
        if kind == 'spin':
            pars.append(f'<p:par><p:cTn id="{n()}" presetID="8" presetClass="emph" presetSubtype="0"{rep} fill="hold" nodeType="{node}">'
                        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
                        f'<p:animRot by="21600000"><p:cBhvr><p:cTn id="{n()}" dur="{int(o["dur"])}" fill="hold"/>'
                        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>'
                        f'</p:childTnLst></p:cTn></p:par>')
        elif kind == 'move':
            path = f'M 0 {o["dy0"]:.6f} L 0 {o["dy1"]:.6f} E'
            pars.append(f'<p:par><p:cTn id="{n()}" presetID="42" presetClass="path" presetSubtype="0" accel="0" decel="0"{rep} fill="hold" nodeType="{node}">'
                        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
                        f'<p:animMotion origin="layout" path="{path}" pathEditMode="relative" rAng="0" ptsTypes="AA">'
                        f'<p:cBhvr><p:cTn id="{n()}" dur="{int(o["dur"])}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
                        f'<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>'
                        f'<p:rCtr x="0" y="0"/></p:animMotion></p:childTnLst></p:cTn></p:par>')
        elif kind == 'fade':
            pars.append(f'<p:par><p:cTn id="{n()}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="{node}">'
                        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
                        f'<p:set><p:cBhvr><p:cTn id="{n()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                        f'<p:to><p:strVal val="visible"/></p:to></p:set>'
                        f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{n()}" dur="{int(o["dur"])}"/>'
                        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
    return (f'<p:timing xmlns:p="{NS["p"]}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            f'<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst>'
            f'<p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            + ''.join(pars) +
            f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
            f'</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
            f'</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')


def add_animations(pptx_path, per_slide):
    """per_slide: {页序号: [(kind, 形状名, opts)]}，写回 pptx。"""
    prs = Presentation(pptx_path)
    for i, effects in per_slide.items():
        s = prs.slides[i]
        ids = shape_ids(s)
        eff = [(k, ids[name], o) for k, name, o in effects if name in ids]
        if not eff:
            continue
        el = etree.fromstring(anim_xml(eff))
        root = s._element
        old = root.find('p:timing', NS)
        if old is not None:
            root.remove(old)
        # timing 要放在 transition 后面、extLst 前面
        ext = root.find('p:extLst', NS)
        if ext is not None:
            ext.addprevious(el)
        else:
            root.append(el)
    prs.save(pptx_path)


# ---------------- 唱片机结束页 ----------------
TITLE_FONT = 'Title'
BODY_FONT = 'Body'

RCX, RCY, RR = 489, 439, 347          # 唱片圆环：圆心、外半径
SCROLL_SPEED = 240                     # 字幕上滚速度（像素/秒，页面宽 1920）
TEXT_X = 1040                          # 字幕左边


def disc_path(size, inner_ratio=0.5):
    c = size / 2
    return pathops.op(circle_path(c, c, c), circle_path(c, c, c * inner_ratio), pathops.PathOp.DIFFERENCE)


def stroke_path(pts, width):
    """折线 → 有宽度的路径（圆角接头），用来画唱臂。"""
    p = pathops.Path()
    for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
        dx, dy = x1 - x0, y1 - y0
        L = math.hypot(dx, dy); nx, ny = -dy / L * width / 2, dx / L * width / 2
        q = pathops.Path(); pen = q.getPen()
        pen.moveTo((x0 + nx, y0 + ny)); pen.lineTo((x1 + nx, y1 + ny)); pen.lineTo((x1 - nx, y1 - ny)); pen.lineTo((x0 - nx, y0 - ny)); pen.closePath()
        p = pathops.op(p, q, pathops.PathOp.UNION)
    for x, y in pts:
        p = pathops.op(p, circle_path(x, y, width / 2), pathops.PathOp.UNION)
    return p


def path_obj(name, p, fill, **kw):
    x0, y0, x1, y1 = p.bounds
    o = {'name': name, 'kind': 'path', 'segs': path_segments_off(p, x0, y0, x1 - x0, y1 - y0),
         'x': x0, 'y': y0, 'w': x1 - x0, 'h': y1 - y0, 'fill': fill}
    o.update(kw)
    return o


def tonearm():
    silver = {'type': 'grad', 'angle': 0, 'stops': [(0, 'B9BDC2', 1), (0.5, 'F4F5F6', 1), (1, '8E9399', 1)]}
    white = {'type': 'grad', 'angle': 90, 'stops': [(0, 'FFFFFF', 1), (1, 'D9DCDF', 1)]}
    sh = {'blur': 6, 'dist': 3, 'dir': 45, 'alpha': 0.45}
    arm = [(910, 136), (853, 338), (847, 385), (803, 498)]
    kids = [path_obj('唱臂杆', stroke_path(arm, 7), silver, shadow=sh),
            {'name': '唱臂底座', 'kind': 'ellipse', 'x': 912 - 33, 'y': 118 - 33, 'w': 66, 'h': 66, 'fill': white, 'shadow': sh},
            {'name': '唱臂旋钮', 'kind': 'ellipse', 'x': 906 - 17, 'y': 112 - 17, 'w': 34, 'h': 34,
             'fill': {'type': 'grad', 'angle': 45, 'stops': [(0, 'FFFFFF', 1), (1, 'C9CDD2', 1)]}},
            {'name': '唱针', 'kind': 'rect', 'x': 770, 'y': 534, 'w': 6, 'h': 22, 'rot': -30, 'fill': silver},
            {'name': '唱头', 'kind': 'roundRect', 'x': 761 - 34, 'y': 524 - 18, 'w': 68, 'h': 36, 'rot': -30,
             'av': '<a:gd name="adj" fmla="val 18000"/>', 'fill': white, 'shadow': sh}]
    return {'name': '唱臂', 'kind': 'group', 'x': 727, 'y': 85, 'w': 218, 'h': 475, 'children': kids}


def text(name, s, x, y, w, h, **st):
    st.setdefault('font', BODY_FONT); st.setdefault('ea', BODY_FONT)
    return {'name': name, 'kind': 'text', 'text': s, 'x': x, 'y': y, 'w': w, 'h': h, 'textstyle': st}


def credits(a, top):
    """标题 + 几段小字，作为一个整体从下往上滚。返回 (组, 高度)。"""
    px = a.body_size * 2; lh = px * 1.2 * 1.5
    th = a.title_size * 2 * 1.25
    kids = [text('结尾标题', a.title, TEXT_X, top, 860, th, font=TITLE_FONT, ea=TITLE_FONT, size=a.title_size, align='left', spacing=4)]
    y = top + th + 50
    paras = [p for p in a.paras if p.strip()]
    for i, para in enumerate(paras):
        lines = para.split('\n')
        h = lh * len(lines)
        kids.append(text(f'结尾正文{i + 1}', para, TEXT_X + 6, y, 760, h, size=a.body_size, align='left', valign='top',
                         line_spacing=1.5, alpha=0.95))
        y += h + lh * 0.8
    y -= lh * 0.8
    return {'name': '字幕', 'kind': 'group', 'x': TEXT_X, 'y': top, 'w': 860, 'h': y - top, 'children': kids}, y - top


def build(a):
    ring = path_obj('唱片', moved_path(disc_path(2 * RR), RCX - RR, RCY - RR), {'type': 'bg'},
                    shadow={'blur': 20, 'dist': 0, 'alpha': 1.0, 'size': 1.04, 'color': '000000'})
    ring['anim'] = {'spin': a.spin * 1000}
    top = 267
    grp, gh = credits(a, top)
    # 起点：整块在页面下边外；终点：整块移出页面上边。循环滚动。
    d0, d1 = H - top, -(top + gh)
    dur = (d0 - d1) / SCROLL_SPEED * 1000 / a.scroll_speed
    grp['anim'] = {'move': [d0, d1], 'dur': dur, 'loop': True}
    o = [{'name': '压暗', 'kind': 'rect', 'x': 0, 'y': 0, 'w': W, 'h': H, 'fill': {'type': 'solid', 'color': '000000', 'alpha': a.dim}},
         ring, tonearm(), grp]
    scene = {'images': {'bg': a.image}, 'slides': [{'bg': 'bg', 'objects': o}], 'vector_text': True}
    anims = {0: [('spin', '唱片', {'dur': a.spin * 1000, 'repeat': True}),
                 ('move', '字幕', {'dy0': d0 / H, 'dy1': d1 / H, 'dur': dur, 'repeat': True})]}
    return scene, anims


def moved_path(p, dx, dy):
    from fontTools.pens.transformPen import TransformPen
    q = pathops.Path(); p.draw(TransformPen(q.getPen(), (1, 0, 0, 1, dx, dy)))
    return q


DEFAULT_PARAS = [
    '谢谢你认真听完\n也谢谢每一次点头和提问',
    '今天的结束\n是下一次出发前的停顿',
    '愿我们各自向前\n下次再见',
]


def main():
    ap = argparse.ArgumentParser(description='唱片机结束页：一张底图 + 透出原图的唱片圆环（一直转）+ 唱臂 + 从下往上滚动的结尾字幕')
    ap.add_argument('--title', default='谢谢聆听', help='结尾标题（毛笔字），4～6 个字最好看')
    ap.add_argument('--para', action='append', metavar='段落', help='标题下面的一段小字，可传多次，每段用 \\n 换行；不传用默认的三段')
    ap.add_argument('--image', default=str(HERE / 'example.jpg'), help='底图，16:9 横图最佳，建议 2K 以上；左边的内容会透在唱片里')
    ap.add_argument('--dim', type=float, default=0.3, help='底图压暗程度，0～1，默认 0.3（30% 黑）')
    ap.add_argument('--spin', type=float, default=7, help='唱片转一圈的秒数')
    ap.add_argument('--scroll-speed', type=float, default=1.0, help='字幕滚动速度倍数，1 是默认速度，2 是两倍快')
    ap.add_argument('--title-size', type=float, default=72, help='标题字号（磅）')
    ap.add_argument('--body-size', type=float, default=24, help='正文字号（磅）')
    ap.add_argument('--title-font', default=str(FONTS / 'MaShanZheng-Regular.ttf'), help='标题字体文件')
    ap.add_argument('--body-font', default=str(FONTS / 'NotoSerifSC-Regular.ttf'), help='正文字体文件')
    ap.add_argument('--out', default='唱片机结束页.pptx', help='输出的 PPT 文件名')
    ap.add_argument('--preview', default=None, help='同时输出网页预览，例如 --preview 预览.html')
    a = ap.parse_args()
    a.title = a.title.replace('\\n', '\n')
    a.paras = [p.replace('\\n', '\n') for p in (a.para or DEFAULT_PARAS)]

    FONT_FILES.update({(TITLE_FONT, False): a.title_font, (BODY_FONT, False): a.body_font})
    scene, anims = build(a)
    build_pptx(scene, a.out)
    add_animations(a.out, anims)
    print('已生成 PPT：', a.out)
    if a.preview:
        build_html(scene, [(TITLE_FONT, a.title_font, 400), (BODY_FONT, a.body_font, 400)], a.preview, '唱片机结束页')
        print('已生成网页预览：', a.preview)


if __name__ == '__main__':
    main()
