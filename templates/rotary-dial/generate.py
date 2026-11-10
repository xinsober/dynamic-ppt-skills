"""旋转转盘 · Rotary Dial

生成一份 4 页的 PPT，页与页之间用「平滑」切换：
  左边是半个双层圆盘（外圆带十字分隔线和 4 个标签），每翻一页逆时针转 90°，
  同时换一张底图，右边文字上下滚动切换。每页文字排法不同：
  第 1 页封面大标题，第 2 页章节标题，第 3 页要点列表，第 4 页展望收尾。

圆盘用的是「幻灯片背景填充」，所以盘面里透出的是当页的底图。所有文字都会转成形状，
换一台没装字体的电脑打开，效果也不会变。PowerPoint 2019 及以上 / Microsoft 365 / WPS 都能播放平滑切换。

用法：
  python generate.py --title "Annual\\nWork\\nReview\\n2026" --images 1.jpg 2.jpg 3.jpg 4.jpg
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
const IM={};let pending=0;
for(const k in SC.images){pending++;const im=new Image();im.onload=()=>{if(--pending===0)ready()};im.src=SC.images[k];IM[k]=im}
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
function drawBg(alphaMul){for(const b of BG){ctx.globalAlpha=b.alpha*(alphaMul??1);ctx.drawImage(IM[b.img],0,0,W,H)}ctx.globalAlpha=1}
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
function drawObj(o,op){op*=(o.opacity??1);if(op<=0.002)return;
 if(o.kind==='group'){withXf(o,()=>{for(const k of o.children)drawObj(k,op)});return}
 if(o.kind==='picture'){withXf(o,()=>{ctx.globalAlpha=op*(o.alpha??1);ctx.drawImage(IM[o.image],o.x,o.y,o.w,o.h);ctx.globalAlpha=1});return}
 if(o.kind==='text'){withXf(o,()=>{if(o.__xfade){const [ta,tb,t]=o.__xfade;drawText(Object.assign({},o,{text:ta}),op*(1-t));drawText(Object.assign({},o,{text:tb}),op*t)}else drawText(o,op)});return}
 withXf(o,()=>{const p=shapePath(o);const t=ctx.getTransform();ctx.__m=t;
  const c=t.transformPoint(new DOMPoint(o.x+o.w/2,o.y+o.h/2));o.__cx=c.x;o.__cy=c.y;
  if(o.shadow)drawShadow(p,o,o.shadow,op);
  if(o.fill){if(o.fill.type==='bg'){ctx.save();ctx.clip(p);ctx.setTransform(1,0,0,1,0,0);drawBg(op);ctx.restore()}
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
let cur=0,raf=null;
function go(n,instant){n=Math.max(0,Math.min(N-1,n));cancelAnimationFrame(raf);const from=cur;cur=n;document.getElementById('pg').textContent=(n+1)+' / '+N;
 if(instant||n!==from+1){render(n);return}const D=SC.slides[n].morph||1500;const t0=performance.now();
 (function f(now){const t=Math.min(1,(now-t0)/D);render(n,t);if(t<1)raf=requestAnimationFrame(f)})(t0)}
function ready(){const fams=new Set();const walk=o=>{if(o.textstyle){fams.add(o.textstyle.font);fams.add(o.textstyle.ea||o.textstyle.font)}(o.children||[]).forEach(walk)};SC.slides.forEach(s=>s.objects.forEach(walk));
 Promise.all([...fams].flatMap(f=>['400','700'].map(w=>document.fonts.load(`${w} 40px "${f}"`,'汉字AB01')))).then(()=>{render(0);document.getElementById('pg').textContent='1 / '+N;window.__ready=true})}
document.getElementById('stage').onclick=()=>go(cur+1);
document.getElementById('next').onclick=()=>go(cur+1);
document.getElementById('prev').onclick=()=>go(cur-1,true);
document.getElementById('play').onclick=()=>{go(0,true);let k=0,tt=600;for(let i=1;i<N;i++){setTimeout(()=>go(i),tt);tt+=(SC.slides[i].morph||1500)+1400}};
addEventListener('keydown',e=>{if(['ArrowRight','ArrowDown',' ','PageDown','Enter'].includes(e.key)){e.preventDefault();go(cur+1)}if(['ArrowLeft','ArrowUp','PageUp'].includes(e.key))go(cur-1,true)});
window.__render=render;
</script></body></html>
'''


# ---------------- 旋转转盘 ----------------
CN = 'Noto Sans SC'
EN = 'Poppins'

CX, CY = 0.0992 * W, 0.5 * H          # 转盘圆心（大半在画面左边外面）
RO, RI = 0.663 * H, 0.293 * H         # 外圆、内圆半径
SH = {'blur': 20, 'dist': 10, 'dir': 0, 'alpha': 0.4}
# 每个标签离圆心的距离（占页面高度的比例）和在转盘上的角度
LABELS = {1: (0.445, 180), 2: (0.455, 270), 3: (0.369, 0), 4: (0.397, 90)}


def text(name, s, x, y, w, h, **st):
    st.setdefault('font', CN); st.setdefault('ea', CN)
    return {'name': name, 'kind': 'text', 'text': s, 'x': x, 'y': y, 'w': w, 'h': h, 'textstyle': st}


def dial(rot, prefix):
    kids = [{'name': '外圆', 'kind': 'flowChartSummingJunction', 'x': CX - RO, 'y': CY - RO, 'w': 2 * RO, 'h': 2 * RO,
             'fill': {'type': 'bg'}, 'line': {'w': 1.5, 'color': 'FFFFFF'}, 'shadow': SH}]
    lw, lh = 300, 70
    for n, (r, ang) in LABELS.items():
        a = math.radians(ang); px = CX + r * H * math.sin(a); py = CY - r * H * math.cos(a)
        label = f'{prefix} 0{n}'.strip()
        kids.append(text(f'标签{n}', label, px - lw / 2, py - lh / 2, lw, lh, font=EN, size=29, bold=True, align='center'))
        kids[-1]['rot'] = ang
    return {'name': '!!转盘', 'kind': 'group', 'x': CX - RO, 'y': CY - RO, 'w': 2 * RO, 'h': 2 * RO, 'rot': rot, 'children': kids}


OVER = {'name': '!!压暗', 'kind': 'rect', 'x': 0, 'y': 0, 'w': W, 'h': H,
        'fill': {'type': 'grad', 'angle': 90, 'stops': [(0, '000000', 0.45), (1, '000000', 0.62)]}}


def page_texts(k, a):
    """第 k 页右边的文字，四页各有一种排法。"""
    T = dict(font=EN, ea=CN, bold=True)
    B = dict(font=EN, ea=CN)
    out = []
    if k == 0:
        if a.title:
            out.append(text('t0_title', a.title, 1080, 225, 690, 430, size=48, align='right', line_spacing=0.95, **T))
            out.append({'name': 't0_bar', 'kind': 'rect', 'x': 1830, 'y': 372, 'w': 4, 'h': 160, 'fill': {'type': 'solid', 'color': 'FFFFFF'}})
        body = '\n'.join(x for x in (a.subtitle, a.reporter) if x)
        if body:
            out.append(text('t0_body', body, 900, 720, 914, 120, size=18, align='right', valign='top', line_spacing=1.1, **B))
        return out
    title, body = a.pages[k - 1]
    if k == 1:
        if title:
            out.append(text('t1_title', title, 1000, 730, 835, 150, size=40, align='right', valign='bottom', line_spacing=0.95, **T))
            out.append({'name': 't1_bar', 'kind': 'rect', 'x': 1660, 'y': 890, 'w': 175, 'h': 4, 'fill': {'type': 'solid', 'color': 'FFFFFF'}})
        if body:
            out.append(text('t1_body', body, 1000, 910, 835, 60, size=15, align='right', valign='top', **B))
    elif k == 2:
        if title:
            out.append(text('t2_title', title, 1080, 600, 780, 120, size=48, align='left', **T))
        if body:
            out.append(text('t2_body', body, 1085, 735, 820, 200, size=18, align='left', valign='top', bullet=True,
                            bullet_indent=40, line_spacing=1.1, **B))
    else:
        if title:
            out.append(text('t3_title', title, 1000, 810, 840, 100, size=40, align='right', **T))
        if body:
            out.append(text('t3_body', body, 900, 915, 940, 60, size=15, align='right', valign='top', **B))
    return out


def build(a):
    slides = []
    for k in range(4):
        o = [OVER, dial(270 - 90 * k, a.label),
             {'name': '!!内圆', 'kind': 'flowChartSummingJunction', 'x': CX - RI, 'y': CY - RI, 'w': 2 * RI, 'h': 2 * RI,
              'fill': {'type': 'bg'}, 'shadow': SH}]
        # 四页的文字都放在每一页里：当前页在画面中，前面的页在上方画外，后面的页在下方画外
        for j in range(4):
            dy = 0 if j == k else (-H if j < k else H)
            for t in page_texts(j, a):
                t = dict(t); t['name'] = '!!' + t['name']; t['y'] = t['y'] + dy
                o.append(t)
        slides.append({'bg': f'bg{k}', 'objects': o, 'morph': a.duration if k else None})
    return {'images': {f'bg{k}': a.images[k] for k in range(4)}, 'slides': slides, 'vector_text': True}


def main():
    ap = argparse.ArgumentParser(description='旋转转盘：生成带平滑切换的 4 页 PPT，左边半个转盘每页转 90°，同时换底图和文字')
    ap.add_argument('--images', nargs='+', metavar='图片',
                    default=[str(HERE / n) for n in ('example.jpg', 'example-2.jpg', 'example-3.jpg', 'example-4.jpg')],
                    help='4 张底图，按页的顺序（每页一张）；少于 4 张会循环使用')
    ap.add_argument('--title', default='Annual\\nWork\\nReview\\n2026', help='第 1 页大标题，用 \\n 换行')
    ap.add_argument('--subtitle', default='Look back. Learn. Move forward.', help='第 1 页标题下面的一句话，留空则不显示')
    ap.add_argument('--reporter', default='汇报人：XXX', help='第 1 页署名，留空则不显示')
    ap.add_argument('--page', action='append', metavar='"标题|正文"',
                    help='第 2、3、4 页的文字，按顺序传 3 次。用 | 分开标题和正文，用 \\n 换行；第 3 页正文每行是一条要点')
    ap.add_argument('--label', default='PART', help='转盘上标签的前缀，显示成「PART 01」…；传空字符串只显示数字')
    ap.add_argument('--en-font', default=str(FONTS / 'Poppins-Regular.ttf'), help='英文正文字体文件')
    ap.add_argument('--en-bold-font', default=str(FONTS / 'Poppins-Bold.ttf'), help='英文标题、标签字体文件')
    ap.add_argument('--cn-font', default=str(FONTS / 'NotoSansSC-Medium.ttf'), help='中文正文字体文件')
    ap.add_argument('--cn-bold-font', default=str(FONTS / 'NotoSansSC-Black.ttf'), help='中文标题字体文件')
    ap.add_argument('--duration', type=int, default=1400, help='平滑切换时长（毫秒）')
    ap.add_argument('--out', default='旋转转盘.pptx', help='输出的 PPT 文件名')
    ap.add_argument('--preview', default=None, help='同时输出网页预览，例如 --preview 预览.html')
    a = ap.parse_args()

    nl = lambda s: s.replace('\\n', '\n').strip()
    a.title, a.subtitle, a.reporter = nl(a.title), nl(a.subtitle), nl(a.reporter)
    raw = a.page or ['Chapter One\\nThe Year in Review|12 projects. 100% on time.',
                     'Key Results|Satisfaction up to 95%\\n3 new business lines\\n6 new team leads',
                     'Looking Ahead|Go deeper. Grow stronger.']
    if len(raw) != 3:
        raise SystemExit('--page 要正好传 3 次（第 2、3、4 页各一次）')
    a.pages = [tuple(nl(p) for p in (s.split('|') + [''])[:2]) for s in raw]
    a.images = [a.images[i % len(a.images)] for i in range(4)]

    FONT_FILES.update({(EN, False): a.en_font, (EN, True): a.en_bold_font, (CN, False): a.cn_font, (CN, True): a.cn_bold_font})
    scene = build(a)
    build_pptx(scene, a.out)
    print('已生成 PPT：', a.out)
    if a.preview:
        build_html(scene, [(EN, a.en_font, 400), (EN, a.en_bold_font, 700), (CN, a.cn_font, 400), (CN, a.cn_bold_font, 700)],
                   a.preview, '旋转转盘')
        print('已生成网页预览：', a.preview)


if __name__ == '__main__':
    main()
