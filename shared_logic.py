import os
import zipfile
import re
from copy import copy
from openpyxl.styles import Font, Alignment, Border, Side, Color, PatternFill

# --- structural constants and styles ---
WON_FMT  = r'_-"₩"* #,##0_-;\\-"₩"* #,##0_-;_-"₩"* "-"_-;_-@_-'
USD_FMT  = r'\$#,##0.00_);[Red]\(\$#,##0.00\)'

def copy_style(src_c, tgt_c):
    if src_c.has_style:
        tgt_c.font = copy(src_c.font)
        tgt_c.border = copy(src_c.border)
        tgt_c.fill = copy(src_c.fill)
        tgt_c.number_format = copy(src_c.number_format)
        tgt_c.alignment = copy(src_c.alignment)

def sc(ws, coord, val=None, fnt=None, al=None, border=None, fill=None, nf=None):
    c = ws[coord]
    if val is not None: c.value = val
    if fnt: c.font = fnt
    if al: c.alignment = al
    if border: c.border = border
    if fill: c.fill = fill
    if nf: c.number_format = nf
    return c

def reconcile_branding_and_xml(orig, out):
    # 1. Full Theme Replication is CRITICAL for correct colors (theme=7, tint=x)
    with zipfile.ZipFile(orig, 'r') as zin, zipfile.ZipFile(out, 'r') as zgen:
        orig_files = {n: zin.read(n) for n in zin.namelist()}
        gen_files = {n: zgen.read(n) for n in zgen.namelist()}
    
    # Inject corporate Theme XML
    if 'xl/theme/theme1.xml' in orig_files:
        gen_files['xl/theme/theme1.xml'] = orig_files['xl/theme/theme1.xml']
    
    # Inject Drawings, Media, VML
    for f in orig_files:
        if f.startswith('xl/media/') or f.startswith('xl/drawings/') or f.startswith('xl/vmlDrawing'):
            gen_files[f] = orig_files[f]
            
    # Inject relationships for sheet 1 (Logo link)
    if 'xl/worksheets/_rels/sheet1.xml.rels' in orig_files:
        gen_files['xl/worksheets/_rels/sheet1.xml.rels'] = orig_files['xl/worksheets/_rels/sheet1.xml.rels']
        
    # Patch [Content_Types].xml to ensure all re-injected types are registered
    ct_xml = gen_files['[Content_Types].xml'].decode('utf-8')
    def ensure_ct(xml, ext, ct):
        if f'Extension="{ext}"' not in xml:
            xml = xml.replace('</Types>', f'<Default Extension="{ext}" ContentType="{ct}"/></Types>')
        return xml
    ct_xml = ensure_ct(ct_xml, 'vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing')
    ct_xml = ensure_ct(ct_xml, 'jpeg', 'image/jpeg')
    ct_xml = ensure_ct(ct_xml, 'png', 'image/png')
    if 'drawing1.xml' not in ct_xml:
        ct_xml = ct_xml.replace('</Types>', '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    gen_files['[Content_Types].xml'] = ct_xml.encode('utf-8')

    # Patch sheet1.xml to link drawing (Logo/Stamp)
    s1_xml = gen_files['xl/worksheets/sheet1.xml'].decode('utf-8')
    if 'xmlns:r' not in s1_xml:
        s1_xml = s1_xml.replace('<worksheet', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    if '<drawing' not in s1_xml:
        s1_xml = s1_xml.replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>')
    gen_files['xl/worksheets/sheet1.xml'] = s1_xml.encode('utf-8')

    # Re-write the ZIP
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, data in gen_files.items(): zout.writestr(n, data)
