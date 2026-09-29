"""시트 안의 링크와 이미지를 꺼내는 도구.

- 링크: xlsx 로 내보내면 한 셀에 링크가 여러 개일 때 하나만 남아서, 시트의 HTML 보기(htmlview)에서
  셀별 (링크 문구, 주소) 목록을 읽는다. 셀 글자 안에 ⟦문구|주소⟧ 표시로 끼워 넣으면 앱이 링크로 그린다.
- 이미지: xlsx 안의 그림(셀에 걸쳐 둔 이미지)을 꺼내 모바일용으로 줄여 media/ 에 저장하고,
  시트별 {(행, 열): [파일 경로]} 로 돌려준다.
"""
import hashlib
import html
import io
import posixpath
import re
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from html.parser import HTMLParser

MAX_SIDE = 1080  # 이미지 긴 변 최대 픽셀
NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
LINK_OPEN, LINK_SEP, LINK_CLOSE = "⟦", "|", "⟧"


# ---------------------------------------------------------------- 링크
def _get(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read().decode("utf-8", "replace")


def _unwrap(url):
    url = html.unescape(url)
    if url.startswith("https://www.google.com/url?"):
        q = urllib.parse.parse_qs(urllib.parse.urlparse(url).query).get("q")
        if q:
            return q[0]
    return url


class _SheetTable(HTMLParser):
    """htmlview 표 → {(행, 열): [(문구, 주소)]} (행·열은 1부터, 병합 칸 고려)"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.links, self.row, self.col = {}, None, 0
        self.busy = {}  # 위에서 rowspan 으로 내려온 칸: row → set(col)
        self.in_th = self.in_td = False
        self.cell = None
        self.anchor = None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "th" and (a.get("id") or "").rsplit("R", 1)[-1].isdigit() and "R" in (a.get("id") or ""):
            self.row = int(a["id"].rsplit("R", 1)[-1]) + 1
            self.col = 0
        elif tag == "td" and self.row is not None:
            if "freezebar" in (a.get("class") or ""):
                return
            self.col += 1
            while self.col in self.busy.get(self.row, ()):
                self.col += 1
            span_c, span_r = int(a.get("colspan", 1)), int(a.get("rowspan", 1))
            for dr in range(span_r):
                for dc in range(span_c):
                    if dr or dc:
                        self.busy.setdefault(self.row + dr, set()).add(self.col + dc)
            self.cell = (self.row, self.col)
            self.col += span_c - 1
            self.in_td = True
        elif tag == "a" and self.in_td and a.get("href"):
            self.anchor = [_unwrap(a["href"]), ""]
        elif tag == "br" and self.anchor is not None:
            self.anchor[1] += "\n"

    def handle_endtag(self, tag):
        if tag == "a" and self.anchor is not None:
            url, label = self.anchor
            self.links.setdefault(self.cell, []).append((label.strip(), url))
            self.anchor = None
        elif tag == "td":
            self.in_td = False
        elif tag == "tr":
            self.row = None

    def handle_data(self, data):
        if self.anchor is not None:
            self.anchor[1] += data


def fetch_links(sheet_id):
    """{시트 이름: {(행, 열): [(문구, 주소)]}} — 실패하면 빈 dict"""
    base = f"https://docs.google.com/spreadsheets/d/{sheet_id}"
    page = _get(f"{base}/htmlview")
    items = re.findall(r'items\.push\(\{name: "((?:[^"\\]|\\.)*)", pageUrl: "[^"]*", gid: "(\d+)"', page)
    out = {}
    for raw_name, gid in items:
        name = raw_name.encode().decode("unicode_escape").encode("latin-1").decode("utf-8")
        name = name.strip().strip("[]").strip()
        t = _SheetTable()
        t.feed(_get(f"{base}/htmlview/sheet?headers=true&gid={gid}"))
        out[name] = t.links
    return out


def mark_links(text, links):
    """셀 글자에서 링크 문구를 찾아 ⟦문구|주소⟧ 로 감싼다. 문구를 못 찾으면 끝에 붙인다."""
    if not links:
        return text
    pos, tail = 0, []
    for label, url in links:
        key = label.split("\n")[0].strip()
        i = text.find(key, pos) if key else -1
        if i < 0:
            tail.append(f"{LINK_OPEN}{key or '링크'}{LINK_SEP}{url}{LINK_CLOSE}")
            continue
        piece = f"{LINK_OPEN}{key}{LINK_SEP}{url}{LINK_CLOSE}"
        text = text[:i] + piece + text[i + len(key):]
        pos = i + len(piece)
    return (text + "\n" + " ".join(tail)).strip() if tail else text


def apply_links(wb, links_by_sheet):
    """워크북 셀 값에 링크 표시를 끼워 넣는다 (htmlview 가 없으면 xlsx 의 셀 링크로 대체)"""
    for ws in wb.worksheets:
        links = links_by_sheet.get(ws.title.strip().strip("[]").strip())
        if links is None:
            links = {(c.row, c.column): [(str(c.value or "").strip(), c.hyperlink.target)]
                     for row in ws.iter_rows() for c in row
                     if c.hyperlink is not None and c.hyperlink.target}
        for (r, c), ls in links.items():
            cell = ws.cell(r, c)
            if isinstance(cell.value, str) or cell.value is None:
                cell.value = mark_links(str(cell.value or ""), ls)


# ---------------------------------------------------------------- 이미지
def _rels(z, path):
    d, f = posixpath.split(path)
    rp = posixpath.join(d, "_rels", f + ".rels")
    if rp not in z.namelist():
        return {}
    out = {}
    for r in ET.fromstring(z.read(rp)).findall("rel:Relationship", NS):
        t = r.get("Target")
        out[r.get("Id")] = t if r.get("TargetMode") == "External" else posixpath.normpath(posixpath.join(d, t))
    return out


def _save_small(data, out_dir):
    from PIL import Image
    name = hashlib.sha1(data).hexdigest()[:16] + ".jpg"
    dest = out_dir / name
    if not dest.exists():
        im = Image.open(io.BytesIO(data))
        im.thumbnail((MAX_SIDE, MAX_SIDE))
        if im.mode in ("RGBA", "LA", "P"):
            im = im.convert("RGBA")
            bg = Image.new("RGB", im.size, "white")
            bg.paste(im, mask=im.split()[-1])
            im = bg
        im.convert("RGB").save(dest, "JPEG", quality=80, optimize=True, progressive=True)
    return name


def extract_images(xlsx_path, out_dir):
    """{시트 이름: {(행, 열): [파일 이름]}} — 이미지를 out_dir 에 jpg 로 저장"""
    out_dir.mkdir(exist_ok=True)
    z = zipfile.ZipFile(xlsx_path)
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    wb_rels = _rels(z, "xl/workbook.xml")
    saved, result = {}, {}
    for s in wb.find("m:sheets", NS):
        path = wb_rels[s.get(f"{{{NS['r']}}}id")]
        s_rels = _rels(z, path)
        cells = {}
        for dr in ET.fromstring(z.read(path)).findall("m:drawing", NS):
            d_path = s_rels[dr.get(f"{{{NS['r']}}}id")]
            d_rels = _rels(z, d_path)
            anchors = []
            for anc in ET.fromstring(z.read(d_path)):
                fr, blip = anc.find("xdr:from", NS), anc.find(".//a:blip", NS)
                if fr is None or blip is None:
                    continue
                media = d_rels.get(blip.get(f"{{{NS['r']}}}embed"))
                if not media or media not in z.namelist():
                    continue
                row = int(fr.find("xdr:row", NS).text) + 1
                col = int(fr.find("xdr:col", NS).text) + 1
                off = int(fr.find("xdr:rowOff", NS).text or 0)
                anchors.append((row, col, off, media))
            for row, col, off, media in sorted(anchors):  # 같은 칸 안에서는 위에서 아래 순서
                if media not in saved:
                    try:
                        saved[media] = _save_small(z.read(media), out_dir)
                    except Exception as e:  # 깨진 이미지는 건너뜀
                        print("skip image", media, e)
                        saved[media] = None
                if saved[media]:
                    cells.setdefault((row, col), []).append(saved[media])
        result[s.get("name").strip().strip("[]").strip()] = cells
    return result
