"""Inspect Origin OLE objects inside PPTX files without opening PowerPoint.

Why this module exists
----------------------
Replacing OLE objects is a visual job, but rendering a 100 MB deck through WPS
for every iteration is slow and burns a lot of context. Most defects can be
diagnosed from the package itself:

* Does the factory OLE's native size (``p:oleObj imgW/imgH``) match the slide
  frame's aspect ratio?
* Does the factory's EMF preview actually contain the expected content (or did
  an off-page object creep in)?
* Is the slide showing an extra preview picture on top of the OLE?

Usage
-----
    python ole_inspect.py frames  deck.pptx
    python ole_inspect.py factory factory.pptx
    python ole_inspect.py preview factory.pptx preview.png
    python ole_inspect.py strip   deck.pptx  deck_no_preview.pptx

Requires: Pillow (only for the ``preview`` command).
"""

from __future__ import annotations

import argparse
import io
import os
import re
import struct
import zipfile
from typing import Optional

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

SLIDE_RE = re.compile(r"ppt/slides/slide(\d+)\.xml$")


def _rels(z: zipfile.ZipFile, slide: int) -> dict[str, str]:
    name = f"ppt/slides/_rels/slide{slide}.xml.rels"
    if name not in z.namelist():
        return {}
    import xml.etree.ElementTree as ET

    return {rel.get("Id"): rel.get("Target") for rel in ET.fromstring(z.read(name))}


def list_frames(pptx: str) -> list[dict]:
    """Return one row per OLE graphicFrame with frame/native size info.

    The values are in EMU. ``pct`` is the relative aspect-ratio difference
    between the preview's native size and the slide frame; more than ~3%
    usually means the picture will be letter-boxed or overflow.
    """
    import xml.etree.ElementTree as ET

    rows = []
    with zipfile.ZipFile(pptx) as z:
        slides = sorted(
            (int(SLIDE_RE.match(n).group(1)), n)
            for n in z.namelist()
            if SLIDE_RE.match(n)
        )
        for sn, name in slides:
            rels = _rels(z, sn)
            root = ET.fromstring(z.read(name))
            for gf in root.iter(P + "graphicFrame"):
                ole = gf.find(".//" + P + "oleObj")
                if ole is None:
                    continue
                # graphicFrame uses <p:xfrm>; some producers use <a:xfrm>.
                ext = gf.find(".//" + P + "xfrm/" + A + "ext")
                if ext is None:
                    ext = gf.find(".//" + A + "xfrm/" + A + "ext")
                rid = ole.get(R + "id")
                iw, ih = ole.get("imgW"), ole.get("imgH")
                cx, cy = (ext.get("cx"), ext.get("cy")) if ext is not None else (None, None)
                row = {
                    "slide": sn,
                    "rid": rid,
                    "embedding": (rels.get(rid) or "").replace("../", "ppt/"),
                    "imgW": int(iw) if iw else None,
                    "imgH": int(ih) if ih else None,
                    "frameW": int(cx) if cx else None,
                    "frameH": int(cy) if cy else None,
                }
                if row["imgW"] and row["frameW"]:
                    ia = row["imgW"] / row["imgH"]
                    fa = row["frameW"] / row["frameH"]
                    row["img_aspect"] = round(ia, 3)
                    row["frame_aspect"] = round(fa, 3)
                    row["pct"] = round(100 * (ia - fa) / fa, 1)
                rows.append(row)
    return rows


def emf_frame_size(data: bytes) -> Optional[tuple[int, int]]:
    """Return the EMF's physical size in 1/100 mm (``rclFrame``).

    EMF header layout: iType(4), nSize(4), rclBounds(16), rclFrame(16).
    ``rclBounds`` is in device units; ``rclFrame`` is the physical size.
    """
    if len(data) < 40 or data[:4] != b"\x01\x00\x00\x00":
        return None
    left, top, right, bottom = struct.unpack_from("<4i", data, 24)
    w, h = right - left, bottom - top
    if w <= 0 or h <= 0:
        return None
    return w, h


def factory_info(factory_pptx: str) -> dict:
    """Describe a factory PPTX: native size, preview media, EMF aspect."""
    with zipfile.ZipFile(factory_pptx) as z:
        xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
        m = re.search(r'imgW="(\d+)" imgH="(\d+)"', xml)
        medias = [
            n
            for n in z.namelist()
            if n.startswith("ppt/media/") and len(z.read(n)) > 40
        ]
        info = {
            "imgW": int(m.group(1)) if m else None,
            "imgH": int(m.group(2)) if m else None,
            "media": medias[0] if medias else None,
        }
        if medias:
            size = emf_frame_size(z.read(medias[0]))
            if size:
                info["emf_frame_0_01mm"] = size
                info["emf_aspect"] = round(size[0] / size[1], 3)
        if info["imgW"]:
            info["native_aspect"] = round(info["imgW"] / info["imgH"], 3)
    return info


def render_preview(factory_pptx: str, out_png: str) -> str:
    """Render the factory's EMF preview to PNG with Pillow (Windows only)."""
    from PIL import Image

    with zipfile.ZipFile(factory_pptx) as z:
        medias = [
            n
            for n in z.namelist()
            if n.startswith("ppt/media/") and len(z.read(n)) > 40
        ]
        if not medias:
            raise RuntimeError("no preview media in factory")
        data = z.read(medias[0])
    Image.open(io.BytesIO(data)).save(out_png)
    return out_png


def strip_ole_previews(src: str, out: str) -> int:
    """Remove the ``<p:pic>`` fallback preview inside every ``<p:oleObj>``.

    WPS can display the OLE preview *and* the fallback picture at once, which
    looks like a ghost image layered over the graph. The OLE itself still
    renders from its embedded presentation after the picture is removed.

    Returns the number of removed pictures.
    """
    ole_re = re.compile(r"(<p:oleObj\b[^>]*>)(.*?)(</p:oleObj>)", re.S)
    pic_re = re.compile(r"<p:pic>.*?</p:pic>", re.S)
    removed = 0

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(
        out, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if SLIDE_RE.match(item.filename):
                xml = data.decode("utf-8")

                def repl(m: re.Match) -> str:
                    nonlocal removed
                    inner = m.group(2)
                    removed += inner.count("<p:pic>")
                    return m.group(1) + pic_re.sub("", inner) + m.group(3)

                xml = ole_re.sub(repl, xml)
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    return removed


def _print_frames(rows: list[dict]) -> None:
    print(f"{'slide':>5} {'rid':<6} {'imgW x imgH':<24} {'frame':<24} "
          f"{'imgAsp':>7} {'frmAsp':>7} {'delta%':>7}  embedding")
    for r in rows:
        print(
            f"{r['slide']:>5} {str(r['rid']):<6} "
            f"{str(r['imgW']) + ' x ' + str(r['imgH']):<24} "
            f"{str(r['frameW']) + ' x ' + str(r['frameH']):<24} "
            f"{str(r.get('img_aspect', '')):>7} {str(r.get('frame_aspect', '')):>7} "
            f"{str(r.get('pct', '')):>7}  {r['embedding']}"
        )


def main() -> None:
    ap = argparse.ArgumentParser(description="Origin OLE inspector for PPTX")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("frames", help="list OLE frames and aspect ratios")
    p.add_argument("pptx")

    p = sub.add_parser("factory", help="describe a factory PPTX")
    p.add_argument("pptx")

    p = sub.add_parser("preview", help="render a factory's EMF preview to PNG")
    p.add_argument("pptx")
    p.add_argument("out_png")

    p = sub.add_parser("strip", help="remove preview pictures inside OLE objects")
    p.add_argument("src")
    p.add_argument("out")

    a = ap.parse_args()
    if a.cmd == "frames":
        _print_frames(list_frames(a.pptx))
    elif a.cmd == "factory":
        info = factory_info(a.pptx)
        for k, v in info.items():
            print(f"{k:20} {v}")
    elif a.cmd == "preview":
        print(render_preview(a.pptx, a.out_png))
    elif a.cmd == "strip":
        n = strip_ole_previews(a.src, a.out)
        print(f"removed {n} preview pictures -> {a.out}")


if __name__ == "__main__":
    main()
