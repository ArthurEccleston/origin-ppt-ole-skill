"""PowerPoint / WPS OLE helpers for Origin graphs.

Key facts learned the hard way:
1. An Origin OLE object stores an *embedded Origin document* in its Contents
   stream. A standalone OPJU saved by originpro is NOT the same format.
   Replacing Contents with a standalone OPJU makes the slide look right
   (fallback image) but double-click opens nothing.
2. The reliable way to create a valid embedded Origin object is:
       Origin: gp.copy_page('OLE')  -> WPS: Shapes.PasteSpecial(10)
   This module's make_factory_ole() does exactly that.
3. WPS uses the VML <mc:Choice> branch (often a WMF preview); modern Office
   uses the <mc:Fallback> EMF picture. Removing the VML choice forces WPS to
   use the EMF fallback and avoids broken WMF rendering.
4. OLE object numbering is not sequential. Always parse slide rels.
5. Validate with WPS, not PowerPoint (the user's machine has no Office license).

This module is Windows-only.
"""

import argparse
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'

# WPS is driven through COM: w.Dispatch('KWPP.Application').
# If you need the executable path, discover it yourself or set WPP_EXE.


# ---------------------------------------------------------------------------
# OLE inventory
# ---------------------------------------------------------------------------

def list_ole_objects(pptx):
    """Return a list of dicts: slide, rid, embedding, fallback_media, has_choice."""
    out = []
    with zipfile.ZipFile(pptx) as z:
        slides = sorted(
            [n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)],
            key=lambda s: int(re.search(r'(\d+)\.xml', s).group(1))
        )
        for sn_xml in slides:
            sn = int(re.search(r'(\d+)\.xml', sn_xml).group(1))
            reln = sn_xml.replace('slides/', 'slides/_rels/') + '.rels'
            rels = {rel.get('Id'): rel.get('Target')
                    for rel in ET.fromstring(z.read(reln))}
            root = ET.fromstring(z.read(sn_xml))
            for frame in root.iter(P + 'graphicFrame'):
                oles = list(frame.iter(P + 'oleObj'))
                if not oles:
                    continue
                rid = oles[0].get(R + 'id')
                target = rels.get(rid, '')
                m = re.search(r'oleObject(\d+)\.bin$', target)
                if not m:
                    continue
                fallback = []
                for pic in frame.iter(P + 'pic'):
                    for blip in pic.iter(A + 'blip'):
                        er = blip.get(R + 'embed')
                        if er in rels:
                            fallback.append(rels[er])
                has_choice = frame.find('.//' + MC + 'Choice') is not None
                out.append({
                    'slide': sn, 'rid': rid,
                    'embedding': f'ppt/embeddings/oleObject{m.group(1)}.bin',
                    'fallback_media': fallback,
                    'has_choice': has_choice,
                })
    return out


def extract_contents(pptx, embedding, out_opju):
    """Extract the embedded Origin document (Contents stream) to an OPJU file."""
    import olefile
    with zipfile.ZipFile(pptx) as z:
        data = z.read(embedding)
    tmp = str(out_opju) + '.bin'
    with open(tmp, 'wb') as f:
        f.write(data)
    ole = olefile.OleFileIO(tmp)
    with open(out_opju, 'wb') as f:
        f.write(ole.openstream('Contents').read())
    try:
        os.remove(tmp)
    except OSError:
        pass
    return out_opju


# ---------------------------------------------------------------------------
# factory OLE: Origin copy -> WPS PasteSpecial
# ---------------------------------------------------------------------------

def make_factory_ole(opju, out_pptx, copy_dpi=300, ratio=100):
    """Create a minimal PPTX containing one genuine embedded Origin graph.

    Returns (embedding_entry, media_entry) names inside out_pptx.
    """
    import win32com.client as w
    import originpro as op

    op.set_show(False)
    op.open(str(opju))
    gps = op.graph_list()
    if not gps:
        raise RuntimeError(f'no graph in {opju}')
    gp = gps[0]
    gp.activate()
    gp.copy_page('OLE', copy_dpi, ratio)

    app = w.Dispatch('KWPP.Application')
    try:
        try:
            app.Visible = 1
        except Exception:
            pass
        pres = app.Presentations.Add()
        sl = pres.Slides.Add(1, 12)  # ppLayoutBlank
        r = sl.Shapes.PasteSpecial(10)  # ppPasteOLEObject
        sh = r.Item(1)
        if sh.Type != 7:
            raise RuntimeError(f'pasted shape is not OLE: Type={sh.Type}')
        pres.SaveAs(str(out_pptx))
        pres.Close()
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        try:
            op.exit()
        except Exception:
            pass

    with zipfile.ZipFile(out_pptx) as z:
        embs = [n for n in z.namelist() if n.startswith('ppt/embeddings/oleObject')]
        medias = [n for n in z.namelist() if n.startswith('ppt/media/image')]
    if not embs:
        raise RuntimeError('factory PPTX has no OLE embedding')
    return embs[0], (medias[0] if medias else None)


# ---------------------------------------------------------------------------
# replace an OLE embedding + fallback EMF (+ remove VML choice)
# ---------------------------------------------------------------------------

def _remove_choice_around(xml, rid):
    idx = xml.find(f'r:id="{rid}"')
    if idx == -1:
        return xml
    start = xml.rfind('<mc:Choice', 0, idx)
    end = xml.find('</mc:Choice>', idx)
    if start == -1 or end == -1:
        return xml
    return xml[:start] + xml[end + len('</mc:Choice>'):]


def replace_ole(pptx, target_embedding, factory_pptx, fallback_media,
                out_pptx, remove_vml=True):
    """Replace a target OLE bin and its fallback EMF with factory versions.

    target_embedding: 'ppt/embeddings/oleObject7.bin'
    fallback_media:   'ppt/media/image15.emf' (from list_ole_objects)
    """
    with zipfile.ZipFile(factory_pptx) as zf:
        fac_embs = [n for n in zf.namelist() if n.startswith('ppt/embeddings/oleObject')]
        fac_media = [n for n in zf.namelist() if n.startswith('ppt/media/image')]
        if not fac_embs:
            raise RuntimeError('factory has no embedding')
        new_bin = zf.read(fac_embs[0])
        new_media = zf.read(fac_media[0]) if fac_media else None

    # map embedding -> rid on each slide (for VML removal)
    ole_map = list_ole_objects(pptx)
    targets = [o for o in ole_map if o['embedding'] == target_embedding]
    slide_rids = {}
    for o in targets:
        slide_rids.setdefault(o['slide'], set()).add(o['rid'])

    with zipfile.ZipFile(pptx) as zin, zipfile.ZipFile(out_pptx, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == target_embedding:
                data = new_bin
            elif fallback_media and item.filename == fallback_media and new_media is not None:
                data = new_media
            elif remove_vml and re.match(r'ppt/slides/slide\d+\.xml$', item.filename):
                sn = int(re.search(r'(\d+)\.xml', item.filename).group(1))
                if sn in slide_rids:
                    xml = data.decode('utf-8')
                    for rid in slide_rids[sn]:
                        xml = _remove_choice_around(xml, rid)
                    data = xml.encode('utf-8')
            zout.writestr(item, data)
    return out_pptx


# ---------------------------------------------------------------------------
# WPS helpers: delete shapes / verify OLE / render slide
# ---------------------------------------------------------------------------

def _wps():
    import win32com.client as w
    return w.Dispatch('KWPP.Application')


def delete_shapes(pptx, slide, names, out_pptx):
    """Delete named shapes on one slide, recursively through groups."""
    app = _wps()
    try:
        try:
            app.Visible = 1
        except Exception:
            pass
        pres = app.Presentations.Open(str(pptx), False, False, True)
        sl = pres.Slides(slide)
        sl.Select()
        names = set(names)

        def walk(shapes):
            for i in range(shapes.Count, 0, -1):
                sh = shapes(i)
                if sh.Name in names:
                    sh.Delete()
                elif sh.Type == 6:
                    walk(sh.GroupItems)

        walk(sl.Shapes)
        pres.SaveAs(str(out_pptx))
        pres.Close()
    finally:
        try:
            app.Quit()
        except Exception:
            pass
    return out_pptx


def render_slide(pptx, slide, png):
    app = _wps()
    try:
        try:
            app.Visible = 1
        except Exception:
            pass
        pres = app.Presentations.Open(str(pptx), True, False, True)
        pres.Slides(slide).Export(str(png), 'PNG', 1920, 1080)
        pres.Close()
    finally:
        try:
            app.Quit()
        except Exception:
            pass
    return png


def verify_ole(pptx, slide):
    """Open in WPS, activate the first OLE on a slide, and inspect Origin.

    Returns dict(ok, origin_window, graphs, plots, x_title).
    """
    import time
    import subprocess
    import win32gui
    import win32process
    import win32com.client as w

    def kill_origin():
        subprocess.run(
            ['powershell', '-NoProfile', '-Command',
             'Get-Process Origin64 -ErrorAction SilentlyContinue | Stop-Process -Force'],
            capture_output=True)
        time.sleep(2)

    def origin_windows():
        found = []

        def cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            t = win32gui.GetWindowText(hwnd)
            if not t or 'Origin' not in t:
                return
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            kids = []

            def cb2(h, _):
                kt = win32gui.GetWindowText(h)
                if kt:
                    kids.append(kt)

            win32gui.EnumChildWindows(hwnd, cb2, None)
            found.append((pid, t, [k for k in kids if 'Graph' in k][:3]))

        win32gui.EnumWindows(cb, None)
        return found

    kill_origin()
    app = _wps()
    result = {'ok': False, 'origin_window': None, 'graphs': [], 'plots': None, 'x_title': None}
    try:
        try:
            app.Visible = 1
        except Exception:
            pass
        pres = app.Presentations.Open(str(pptx), False, False, True)
        sl = pres.Slides(slide)
        sl.Select()
        target = None

        def walk(shapes):
            nonlocal target
            for i in range(1, shapes.Count + 1):
                sh = shapes(i)
                try:
                    if sh.Type == 7:
                        target = sh
                        return
                except Exception:
                    pass
                if sh.Type == 6:
                    walk(sh.GroupItems)

        walk(sl.Shapes)
        if target is None:
            return result
        target.OLEFormat.DoVerb(1)
        time.sleep(9)
        result['origin_window'] = origin_windows() or None
        try:
            import originpro as op
            op.attach()
            gps = op.graph_list()
            result['graphs'] = [gp.name for gp in gps]
            if gps:
                gl = gps[0][0]
                result['plots'] = len(gl.plot_list())
                result['x_title'] = gl.axis('x').title
            result['ok'] = bool(gps)
        except Exception as e:
            result['error'] = f'{type(e).__name__}: {e}'
        pres.Close()
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        kill_origin()
    return result


if __name__ == '__main__':
    ap = argparse.ArgumentParser(description='PowerPoint/Origin OLE helpers')
    ap.add_argument('--list', help='list OLE objects in a PPTX')
    ap.add_argument('--extract', nargs=2, metavar=('PPTX', 'EMBEDDING'))
    ap.add_argument('--out', default='')
    a = ap.parse_args()
    if a.list:
        for o in list_ole_objects(a.list):
            print(o)
    elif a.extract:
        print(extract_contents(a.extract[0], a.extract[1], a.out))
