# -*- coding: utf-8 -*-
"""Image extraction utilities.

仕様書参照: §8 画像抽出規約
"""

import re
import zipfile as _zipfile
import xml.etree.ElementTree as _ET
from pathlib import Path
from typing import Dict, Tuple

from .output import warn, info

# DrawingML名前空間
_DRAWINGML_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}

def extract_images_from_xlsx_drawing(xlsx_path: str, ws, output_dir: Path, sheet_name: str, md_basename: str, opts) -> Dict[Tuple[int, int], str]:
    """Extract images directly from xlsx ZIP archive using DrawingML.

    This function reads the xlsx file as a ZIP archive and extracts embedded images
    by parsing DrawingML files. This approach works even when openpyxl's ws._images
    is empty, which can happen with certain Excel file types or loading modes.

    Args:
        xlsx_path: Path to the Excel (.xlsx) file
        ws: openpyxl worksheet object (used to get sheet title)
        output_dir: Base directory where images will be saved
        sheet_name: Name of the current sheet
        md_basename: Base name of the markdown file (used as subdirectory name)
        opts: Options dictionary

    Returns:
        Dict mapping (row, col) tuples to relative image paths.
        Keys are 1-based Excel coordinates.

    Note:
        - Images are saved in a subdirectory named {md_basename}_images
        - Filename format: {sanitized_sheet_name}_img_{index}.{extension}
    """
    cell_to_image = {}

    try:
        z = _zipfile.ZipFile(xlsx_path, 'r')
    except Exception as e:
        warn(f"Failed to open xlsx file as ZIP: {e}")
        return cell_to_image

    try:
        # 1. Find sheet ID from workbook.xml
        sheet_id = None
        try:
            wb_xml = z.read('xl/workbook.xml').decode('utf-8')
            wb_root = _ET.fromstring(wb_xml)
            ns_wb = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                     'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
            sheets = wb_root.findall('.//main:sheet', ns_wb)
            for sheet in sheets:
                if sheet.get('name') == ws.title:
                    r_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if r_id:
                        wb_rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
                        wb_rels_root = _ET.fromstring(wb_rels_xml)
                        ns_rel = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                        for rel in wb_rels_root.findall('.//r:Relationship', ns_rel):
                            if rel.get('Id') == r_id:
                                target = rel.get('Target')
                                match = re.search(r'sheet(\d+)\.xml', target)
                                if match:
                                    sheet_id = match.group(1)
                                break
                    break
        except Exception as e:
            warn(f"Failed to find sheet ID for '{ws.title}': {e}")
            return cell_to_image

        if not sheet_id:
            return cell_to_image

        # 2. Find drawing file from sheet relationship file
        sheet_rel_path = f"xl/worksheets/_rels/sheet{sheet_id}.xml.rels"
        if sheet_rel_path not in z.namelist():
            return cell_to_image

        drawing_path = None
        drawing_rels_path = None
        try:
            rels_xml = z.read(sheet_rel_path)
            rels_root = _ET.fromstring(rels_xml)
            ns_rel = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            for rel in rels_root.findall('.//r:Relationship', ns_rel):
                if rel.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing':
                    target = rel.get('Target')
                    if target.startswith('../'):
                        drawing_path = target.replace('../', 'xl/')
                    else:
                        drawing_path = f"xl/worksheets/{target}"
                    # Calculate drawing rels path
                    drawing_filename = drawing_path.split('/')[-1]
                    drawing_dir = '/'.join(drawing_path.split('/')[:-1])
                    drawing_rels_path = f"{drawing_dir}/_rels/{drawing_filename}.rels"
                    break
        except Exception as e:
            warn(f"Failed to parse relationship file '{sheet_rel_path}': {e}")
            return cell_to_image

        if not drawing_path or drawing_path not in z.namelist():
            return cell_to_image

        # 3. Parse drawing relationship file to get image paths
        image_rels = {}  # rId -> image path in zip
        if drawing_rels_path and drawing_rels_path in z.namelist():
            try:
                drawing_rels_xml = z.read(drawing_rels_path)
                drawing_rels_root = _ET.fromstring(drawing_rels_xml)
                ns_rel = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                for rel in drawing_rels_root.findall('.//r:Relationship', ns_rel):
                    if 'image' in rel.get('Type', '').lower():
                        target = rel.get('Target')
                        r_id = rel.get('Id')
                        if target.startswith('../'):
                            image_path = 'xl/' + target.replace('../', '')
                        else:
                            image_path = target
                        image_rels[r_id] = image_path
            except Exception as e:
                warn(f"Failed to parse drawing rels '{drawing_rels_path}': {e}")

        if not image_rels:
            return cell_to_image

        # 4. Parse drawing XML to get image positions
        drawing_xml = z.read(drawing_path)
        drawing_root = _ET.fromstring(drawing_xml)

        ns = {
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

        # Create output directory
        images_dir = output_dir / f"{md_basename}_images"
        images_dir.mkdir(parents=True, exist_ok=True)
        sanitized_sheet = sanitize_sheet_name(sheet_name)

        img_idx = 0
        # Find all picture elements (xdr:pic)
        for anchor in drawing_root.findall('.//xdr:twoCellAnchor', ns) + drawing_root.findall('.//xdr:oneCellAnchor', ns):
            pic = anchor.find('.//xdr:pic', ns)
            if pic is None:
                continue

            # Get image reference (blip)
            blip = pic.find('.//a:blip', ns)
            if blip is None:
                continue

            # Get rId from r:embed attribute
            embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if not embed_id or embed_id not in image_rels:
                continue

            # Get cell position from anchor
            from_elem = anchor.find('xdr:from', ns)
            if from_elem is not None:
                col_elem = from_elem.find('xdr:col', ns)
                row_elem = from_elem.find('xdr:row', ns)
                if col_elem is not None and row_elem is not None:
                    try:
                        cell_col = int(col_elem.text) + 1  # Convert 0-based to 1-based
                        cell_row = int(row_elem.text) + 1
                    except ValueError:
                        continue
                else:
                    continue
            else:
                continue

            # Extract and save image
            img_idx += 1
            image_zip_path = image_rels[embed_id]
            if image_zip_path not in z.namelist():
                warn(f"Image file not found in xlsx: {image_zip_path}")
                continue

            try:
                img_data = z.read(image_zip_path)

                # Determine extension from path or magic bytes
                file_extension = image_zip_path.split('.')[-1].lower()
                if file_extension not in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'webp'):
                    if img_data.startswith(b'\x89PNG'):
                        file_extension = "png"
                    elif img_data.startswith(b'\xff\xd8\xff'):
                        file_extension = "jpg"
                    elif img_data.startswith(b'GIF'):
                        file_extension = "gif"
                    else:
                        file_extension = "png"  # default

                img_filename = f"{sanitized_sheet}_img_{img_idx}.{file_extension}"
                img_path = images_dir / img_filename

                with open(img_path, 'wb') as f:
                    f.write(img_data)

                relative_path = f"{md_basename}_images/{img_filename}"
                cell_to_image[(cell_row, cell_col)] = relative_path

                info(f"Extracted image to {relative_path} at cell ({cell_row}, {cell_col})")

            except Exception as e:
                warn(f"Failed to extract image from {image_zip_path}: {e}")
                continue

    finally:
        z.close()

    return cell_to_image


def extract_images_from_sheet(ws, output_dir: Path, sheet_name: str, md_basename: str, opts, xlsx_path: str = None) -> Dict[Tuple[int, int], str]:
    """Extract images from worksheet and save to external files.

    This function first tries to use the new DrawingML-based extraction (which works
    with all Excel files), then falls back to openpyxl's ws._images if needed.

    This function processes all images in the worksheet, saves them to a subdirectory,
    and returns a mapping of cell coordinates to image paths for Markdown link generation.

    Args:
        ws: openpyxl worksheet object containing images
        output_dir: Base directory where images will be saved
        sheet_name: Name of the current sheet (used in filename generation)
        md_basename: Base name of the markdown file (used as subdirectory name)
        opts: Options dictionary (for future extensibility)
        xlsx_path: Path to the Excel file (required for DrawingML extraction)

    Returns:
        Dict mapping (row, col) tuples to relative image paths.
        Keys are 1-based Excel coordinates, values are paths relative to output_dir.

    Note:
        - Images are saved in a subdirectory named {md_basename}_images
        - Filename format: {sanitized_sheet_name}_img_{index}.{extension}
        - Supports PNG, JPG, GIF, and other common image formats
    """
    cell_to_image = {}

    # Try DrawingML-based extraction first (works with all Excel files)
    if xlsx_path:
        cell_to_image = extract_images_from_xlsx_drawing(
            xlsx_path, ws, output_dir, sheet_name, md_basename, opts
        )
        if cell_to_image:
            return cell_to_image

    # Fallback to openpyxl's ws._images (may be empty depending on load mode)
    if not hasattr(ws, '_images') or not ws._images:
        return cell_to_image

    # Create subdirectory for images (named after markdown file with _images suffix)
    images_dir = output_dir / f"{md_basename}_images"
    images_dir.mkdir(parents=True, exist_ok=True)

    sanitized_sheet = sanitize_sheet_name(sheet_name)

    for img_idx, img in enumerate(ws._images, start=1):
        try:
            # Extract raw image data from openpyxl image object
            img_data = img._data()

            # Determine file extension from image format
            # Priority: 1) format attribute, 2) magic bytes detection
            file_extension = "png"  # default fallback
            if hasattr(img, 'format'):
                file_extension = img.format.lower()
            elif hasattr(img, '_data'):
                # Detect format from magic bytes (file signatures)
                # PNG: 89 50 4E 47 (\x89PNG)
                # JPEG: FF D8 FF
                # GIF: 47 49 46 (GIF)
                if img_data.startswith(b'\x89PNG'):
                    file_extension = "png"
                elif img_data.startswith(b'\xff\xd8\xff'):
                    file_extension = "jpg"
                elif img_data.startswith(b'GIF'):
                    file_extension = "gif"

            # Generate unique filename for this image
            img_filename = f"{sanitized_sheet}_img_{img_idx}.{file_extension}"
            img_path = images_dir / img_filename

            # Write image data to file
            with open(img_path, 'wb') as f:
                f.write(img_data)

            # Extract cell position from image anchor
            # Different anchor types store position information differently
            if hasattr(img, 'anchor'):
                anchor = img.anchor
                cell_row = None
                cell_col = None

                # TwoCellAnchor: has _from attribute with row/col
                if hasattr(anchor, '_from'):
                    cell_ref = anchor._from
                    # Convert from 0-based (openpyxl) to 1-based (Excel)
                    cell_row = cell_ref.row + 1
                    cell_col = cell_ref.col + 1
                # OneCellAnchor or other types: may have direct row/col
                elif hasattr(anchor, 'row') and hasattr(anchor, 'col'):
                    # Convert from 0-based to 1-based
                    cell_row = anchor.row + 1
                    cell_col = anchor.col + 1
                else:
                    # Unknown anchor type - skip this image
                    warn(f"Unknown anchor type for image {img_idx}, skipping position detection")
                    continue

                # Store mapping with relative path (for portability)
                relative_path = f"{md_basename}_images/{img_filename}"
                cell_to_image[(cell_row, cell_col)] = relative_path

                info(f"Extracted image to {relative_path} at cell ({cell_row}, {cell_col})")

        except Exception as e:
            warn(f"Failed to extract image {img_idx}: {e}")
            continue

    return cell_to_image
def sanitize_sheet_name(sheet_name: str) -> str:
    """Sanitize sheet name for use in filenames."""
    # Replace filesystem-unsafe characters with underscore
    unsafe_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    sanitized = sheet_name
    for ch in unsafe_chars:
        sanitized = sanitized.replace(ch, '_')
    return sanitized
