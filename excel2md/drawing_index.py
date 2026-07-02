# -*- coding: utf-8 -*-
"""Workbook-level DrawingML index.

画像抽出 (image_extraction) と shapes モードの Mermaid 図形抽出 (mermaid_generator)
は、いずれも「シート名 → sheet rId → sheet ファイル番号 → drawing ファイル」の解決の
ために xlsx を ZIP として開き、workbook.xml / workbook.xml.rels / sheet rels を
シートごと・用途ごとに解析し直していた (Issue #26)。

本モジュールはその解決をワークブック単位で1回に集約する。索引の生成時に ZIP を開いて
関係ファイルを解析し、シート名から drawing ファイルのパスと画像リレーション
(rId → メディアパス) を引けるようにする。ZIP ハンドルは索引が保持するため、
利用側は `read()` / `exists()` でアーカイブ内のファイルへアクセスし、
使い終わったら `close()` する (context manager としても使える)。
"""

import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, Optional

from .output import warn

_NS_MAIN = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
_NS_PKG_REL = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
_R_ID_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"


class WorkbookDrawingIndex:
    """xlsx ZIP を1回だけ解析し、シート名から DrawingML 情報を引く索引。

    開けない・解析できないファイルに対しては warn を出したうえで空の索引として
    振る舞う (drawing_path() は None、image_rels() は {} を返す)。
    """

    def __init__(self, xlsx_path: str):
        self.xlsx_path = xlsx_path
        self._zip: Optional[zipfile.ZipFile] = None
        self._names = set()
        # sheet_name -> {"drawing_path": Optional[str], "image_rels": Dict[rId, media path]}
        self._sheets: Dict[str, dict] = {}
        try:
            self._zip = zipfile.ZipFile(xlsx_path, "r")
        except Exception as e:
            warn(f"Failed to open xlsx file as ZIP: {e}")
            return
        try:
            self._names = set(self._zip.namelist())
            self._parse()
        except Exception as e:
            warn(f"Failed to parse workbook drawing information: {e}")
            self._sheets = {}

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def drawing_path(self, sheet_name: str) -> Optional[str]:
        """シートに対応する drawing ファイルの ZIP 内パスを返す (無ければ None)。"""
        return self._sheets.get(sheet_name, {}).get("drawing_path")

    def image_rels(self, sheet_name: str) -> Dict[str, str]:
        """シートの drawing が参照する画像リレーション (rId → ZIP 内パス) を返す。"""
        return self._sheets.get(sheet_name, {}).get("image_rels", {})

    def exists(self, member: str) -> bool:
        """ZIP 内にファイルが存在するかを返す。"""
        return member in self._names

    def read(self, member: str) -> bytes:
        """ZIP 内のファイルを読み出す。"""
        return self._zip.read(member)

    def close(self):
        if self._zip is not None:
            self._zip.close()
            self._zip = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False

    # ------------------------------------------------------------------
    # Parsing
    # ------------------------------------------------------------------

    def _parse(self):
        z = self._zip

        # 1. シート名 → rId (workbook.xml)
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        rid_by_sheet = {}
        for sheet in wb_root.findall(".//main:sheet", _NS_MAIN):
            name = sheet.get("name")
            r_id = sheet.get(_R_ID_ATTR)
            if name and r_id:
                rid_by_sheet[name] = r_id

        # 2. rId → シートファイル番号 (workbook.xml.rels; 例: worksheets/sheet3.xml → "3")
        rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        sheet_id_by_rid = {}
        for rel in rels_root.findall(".//r:Relationship", _NS_PKG_REL):
            match = re.search(r"sheet(\d+)\.xml", rel.get("Target") or "")
            if match:
                sheet_id_by_rid[rel.get("Id")] = match.group(1)

        # 3. シートごとに drawing パスと画像リレーションを解決
        for name, r_id in rid_by_sheet.items():
            info = {"drawing_path": None, "image_rels": {}}
            self._sheets[name] = info

            sheet_id = sheet_id_by_rid.get(r_id)
            if not sheet_id:
                continue
            sheet_rel_path = f"xl/worksheets/_rels/sheet{sheet_id}.xml.rels"
            if sheet_rel_path not in self._names:
                continue
            try:
                sheet_rels_root = ET.fromstring(z.read(sheet_rel_path))
            except Exception as e:
                warn(f"Failed to parse relationship file '{sheet_rel_path}': {e}")
                continue

            drawing_path = None
            for rel in sheet_rels_root.findall(".//r:Relationship", _NS_PKG_REL):
                if rel.get("Type") == _DRAWING_REL_TYPE:
                    target = rel.get("Target") or ""
                    if target.startswith("../"):
                        drawing_path = target.replace("../", "xl/")
                    else:
                        drawing_path = f"xl/worksheets/{target}"
                    break
            if not drawing_path or drawing_path not in self._names:
                continue
            info["drawing_path"] = drawing_path

            # drawing rels → 画像パス (rId → xl/media/...)
            drawing_filename = drawing_path.split("/")[-1]
            drawing_dir = "/".join(drawing_path.split("/")[:-1])
            drawing_rels_path = f"{drawing_dir}/_rels/{drawing_filename}.rels"
            if drawing_rels_path not in self._names:
                continue
            try:
                drawing_rels_root = ET.fromstring(z.read(drawing_rels_path))
            except Exception as e:
                warn(f"Failed to parse drawing rels '{drawing_rels_path}': {e}")
                continue
            for rel in drawing_rels_root.findall(".//r:Relationship", _NS_PKG_REL):
                if "image" in (rel.get("Type") or "").lower():
                    target = rel.get("Target") or ""
                    if target.startswith("../"):
                        image_path = "xl/" + target.replace("../", "")
                    else:
                        image_path = target
                    info["image_rels"][rel.get("Id")] = image_path
