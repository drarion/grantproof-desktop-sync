from __future__ import annotations

import math
import re
import tempfile
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

import shapefile  # pyshp
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFilter, ImageOps

IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
VIDEO_EXTENSIONS = {'.mp4', '.mov', '.m4v', '.avi', '.mkv', '.webm'}

BLUE = '0A57B5'
BLUE_DARK = '064890'
BLUE_MID = '1665C1'
BLUE_SOFT = 'EEF5FF'
BLUE_PALE = 'F6FAFF'
ORANGE = 'F56D28'
DARK = '1D2D44'
MUTED = '5D6B7C'
BORDER = 'D9E4F2'
BORDER_LIGHT = 'E8EEF7'
PANEL = 'FFFFFF'
PANEL_SOFT = 'F7FAFF'

COUNTRY_ALIASES = {
    'niger': 'Niger',
    'france': 'France',
    'république démocratique du congo': 'Dem. Rep. Congo',
    'republique democratique du congo': 'Dem. Rep. Congo',
    'rd congo': 'Dem. Rep. Congo',
    'rdc': 'Dem. Rep. Congo',
    'democratic republic of congo': 'Dem. Rep. Congo',
    'drc': 'Dem. Rep. Congo',
    'burkina faso': 'Burkina Faso',
    'mali': 'Mali',
    'tchad': 'Chad',
    'chad': 'Chad',
    'nigeria': 'Nigeria',
    'sénégal': 'Senegal',
    'senegal': 'Senegal',
    'cameroun': 'Cameroon',
    'cameroon': 'Cameroon',
    'côte d’ivoire': "Côte d'Ivoire",
    'côte d\'ivoire': "Côte d'Ivoire",
    'cote d ivoire': "Côte d'Ivoire",
    'south sudan': 'S. Sudan',
    'soudan du sud': 'S. Sudan',
    'sudan': 'Sudan',
}

CAPITAL_COORDS = {
    'Niger': (13.5116, 2.1254, 'Niamey'),
    'France': (48.8566, 2.3522, 'Paris'),
    'Dem. Rep. Congo': (-4.4419, 15.2663, 'Kinshasa'),
    'Burkina Faso': (12.3714, -1.5197, 'Ouagadougou'),
    'Mali': (12.6392, -8.0029, 'Bamako'),
    'Chad': (12.1348, 15.0557, "N'Djamena"),
    'Nigeria': (9.0765, 7.3986, 'Abuja'),
    'Senegal': (14.7167, -17.4677, 'Dakar'),
    'Cameroon': (3.8480, 11.5021, 'Yaoundé'),
    "Côte d'Ivoire": (5.3600, -4.0083, 'Abidjan'),
    'S. Sudan': (4.8594, 31.5713, 'Juba'),
    'Sudan': (15.5007, 32.5599, 'Khartoum'),
}

COUNTRY_DISPLAY_FR = {
    'Niger': 'Niger',
    'France': 'France',
    'Dem. Rep. Congo': 'RD Congo',
    'Burkina Faso': 'Burkina Faso',
    'Mali': 'Mali',
    'Chad': 'Tchad',
    'Nigeria': 'Nigeria',
    'Senegal': 'Sénégal',
    'Cameroon': 'Cameroun',
    "Côte d'Ivoire": "Côte d’Ivoire",
    'S. Sudan': 'Soudan du Sud',
    'Sudan': 'Soudan',
}

SECTOR_RULES = [
    {
        'key': 'shelter',
        'fr': 'Abris',
        'en': 'Shelter',
        'icon': 'shelter',
        'keywords': ['abri', 'abris', 'shelter', 'logement', 'construction', 'réhabilitation', 'rehabilitation', 'maison', 'habitat', 'ame', 'nfi'],
    },
    {
        'key': 'food_security',
        'fr': 'Sécurité alimentaire',
        'en': 'Food security',
        'icon': 'food',
        'keywords': ['sécurité alimentaire', 'food security', 'vivres', 'food', 'ration', 'nutrition sensitive', 'maraîcher', 'maraicher', 'semence', 'intrants', 'agricole', 'agriculture', 'cash for food'],
    },
    {
        'key': 'agriculture',
        'fr': 'Agriculture',
        'en': 'Agriculture',
        'icon': 'agriculture',
        'keywords': ['agriculture', 'agricole', 'maraîcher', 'maraicher', 'semence', 'irrigation', 'élevage', 'livelihood', 'moyens d’existence', 'moyens existence'],
    },
    {
        'key': 'wash',
        'fr': 'Eau, hygiène et assainissement',
        'en': 'WASH',
        'icon': 'wash',
        'keywords': ['wash', 'eha', 'eau', 'hygiène', 'hygiene', 'assainissement', 'latrine', 'chloration', 'forage', 'water', 'sanitation'],
    },
    {
        'key': 'health',
        'fr': 'Santé',
        'en': 'Health',
        'icon': 'health',
        'keywords': ['santé', 'health', 'clinique', 'centre de santé', 'médicament', 'vaccination', 'choléra', 'cholera', 'soins'],
    },
    {
        'key': 'nutrition',
        'fr': 'Nutrition',
        'en': 'Nutrition',
        'icon': 'nutrition',
        'keywords': ['nutrition', 'malnutrition', 'anjé', 'anje', 'mas', 'mam', 'dépistage', 'dépister'],
    },
    {
        'key': 'education',
        'fr': 'Éducation',
        'en': 'Education',
        'icon': 'education',
        'keywords': ['éducation', 'education', 'école', 'ecole', 'enseignant', 'classe', 'élève', 'scolaire', 'school'],
    },
    {
        'key': 'protection',
        'fr': 'Protection',
        'en': 'Protection',
        'icon': 'protection',
        'keywords': ['protection', 'vbg', 'gbv', 'violence', 'enfant', 'psychosocial', 'psea', 'réunification', 'protection de l’enfance'],
    },
    {
        'key': 'cash',
        'fr': 'Transferts monétaires',
        'en': 'Cash assistance',
        'icon': 'cash',
        'keywords': ['cash', 'transfert monétaire', 'transferts monétaires', 'voucher', 'coupon', 'assistance monétaire'],
    },
    {
        'key': 'coordination',
        'fr': 'Coordination',
        'en': 'Coordination',
        'icon': 'coordination',
        'keywords': ['coordination', 'réunion', 'meeting', 'cluster', 'atelier', 'planification'],
    },
]

QUANTITY_PATTERNS = [
    ('shelter', r'(?:construction|construit|construits|réhabilitation|rehabilitation)?\s*(?:de\s*)?(\d[\d\s.,]*)\s*(abris?|shelters?)', 'Abris', 'Shelters'),
    ('latrine', r'(\d[\d\s.,]*)\s*(latrines?)', 'Latrines', 'Latrines'),
    ('school', r'(\d[\d\s.,]*)\s*(écoles?|ecoles?|schools?)', 'Écoles', 'Schools'),
    ('kit', r'(\d[\d\s.,]*)\s*(kits?)', 'Kits', 'Kits'),
    ('household', r'(\d[\d\s.,]*)\s*(ménages?|menages?|households?)', 'Ménages', 'Households'),
    ('beneficiary', r'(\d[\d\s.,]*)\s*(bénéficiaires?|beneficiaires?|beneficiaries|participants?|personnes?)', 'Bénéficiaires', 'Beneficiaries'),
]


def _safe_text(value: Any, fallback: str = '') -> str:
    if value is None:
        return fallback
    text = str(value).strip()
    return text if text else fallback


def _clean_spaces(value: Any) -> str:
    return ' '.join(_safe_text(value).replace('\n', ' ').split())


def _sentence(text: str) -> str:
    text = _clean_spaces(text)
    if text and text[-1] not in '.!?':
        text += '.'
    return text


def _hex_color(hex_value: str) -> RGBColor:
    return RGBColor.from_string(hex_value)


def _parse_float(value: Any) -> float | None:
    try:
        if value is None:
            return None
        if isinstance(value, str):
            value = value.replace(',', '.').strip()
        return float(value)
    except Exception:
        return None


def _format_date(value: Any, lang: str) -> str:
    text = _safe_text(value)
    if not text:
        return datetime.now().strftime('%d/%m/%Y' if lang == 'fr' else '%Y-%m-%d')
    try:
        dt = datetime.fromisoformat(text.replace('Z', '+00:00'))
        return dt.strftime('%d/%m/%Y' if lang == 'fr' else '%Y-%m-%d')
    except Exception:
        return text[:10] if len(text) >= 10 else text


def _set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tc_pr.append(shd)
    shd.set(qn('w:fill'), fill)


def _set_cell_border(cell, **kwargs) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_borders = tc_pr.first_child_found_in('w:tcBorders')
    if tc_borders is None:
        tc_borders = OxmlElement('w:tcBorders')
        tc_pr.append(tc_borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        if edge not in kwargs:
            continue
        tag = f'w:{edge}'
        element = tc_borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            tc_borders.append(element)
        for key, value in kwargs[edge].items():
            element.set(qn(f'w:{key}'), str(value))


def _clear_cell_border(cell) -> None:
    _set_cell_border(
        cell,
        top={'val': 'nil'}, bottom={'val': 'nil'}, left={'val': 'nil'}, right={'val': 'nil'},
        insideH={'val': 'nil'}, insideV={'val': 'nil'},
    )


def _clear_table_borders(table) -> None:
    for row in table.rows:
        for cell in row.cells:
            _clear_cell_border(cell)


def _set_cell_margins(cell, top: int = 80, start: int = 80, bottom: int = 80, end: int = 80) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in('w:tcMar')
    if tc_mar is None:
        tc_mar = OxmlElement('w:tcMar')
        tc_pr.append(tc_mar)
    for margin_name, value in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        node = tc_mar.find(qn(f'w:{margin_name}'))
        if node is None:
            node = OxmlElement(f'w:{margin_name}')
            tc_mar.append(node)
        node.set(qn('w:w'), str(value))
        node.set(qn('w:type'), 'dxa')


def _set_cell_width(cell, width_twips: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn('w:tcW'))
    if tc_w is None:
        tc_w = OxmlElement('w:tcW')
        tc_pr.append(tc_w)
    tc_w.set(qn('w:w'), str(width_twips))
    tc_w.set(qn('w:type'), 'dxa')


def _add_run(paragraph, text: str, *, bold: bool = False, italic: bool = False, size: float | None = None, color: str | None = None, all_caps: bool = False):
    run = paragraph.add_run(text.upper() if all_caps else text)
    run.bold = bold
    run.italic = italic
    run.font.name = 'Aptos'
    if size is not None:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = _hex_color(color)
    try:
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Aptos')
    except Exception:
        pass
    return run


def _configure_paragraph(paragraph, *, before: float = 0, after: float = 0, line: float = 1.0, align=None) -> None:
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line
    if align is not None:
        paragraph.alignment = align



def _clear_cell_content(cell) -> None:
    for paragraph in list(cell.paragraphs):
        parent = paragraph._element.getparent()
        if parent is not None:
            parent.remove(paragraph._element)
    for table in list(cell.tables):
        parent = table._element.getparent()
        if parent is not None:
            parent.remove(table._element)

def _first_paragraph(cell):
    if cell.paragraphs:
        return cell.paragraphs[0]
    return cell.add_paragraph()


class PremiumActivityReportBuilder:
    """Generates an NGO-owned, editable DOCX activity report from synced evidence.

    The layout is intentionally compact: page 1 is a premium dashboard-style report,
    and page 2 is created only when there are additional images to annex.
    """

    def __init__(self, base_folder: Path) -> None:
        self.base_folder = Path(base_folder)
        self.assets_dir = Path(__file__).resolve().parent / 'assets'
        self.maps_dir = self.assets_dir / 'maps' / 'naturalearth_lowres'

    def build(self, output: Path, project_code: str, project_name: str, records: list[Any], lang: str) -> None:
        lang = 'en' if str(lang).lower().startswith('en') else 'fr'
        if not records:
            return
        data = self._prepare_data(project_code, project_name, records, lang)
        with tempfile.TemporaryDirectory(prefix='grantproof_report_assets_') as temp_root:
            temp_dir = Path(temp_root)
            visual_assets = self._create_visual_assets(data, temp_dir, lang)
            doc = Document()
            self._configure_document(doc)
            self._build_first_page(doc, data, visual_assets, lang)
            if data['annex_images']:
                self._build_visual_annex(doc, data, visual_assets, lang)
            doc.save(output)

    def _prepare_data(self, project_code: str, project_name: str, records: list[Any], lang: str) -> dict[str, Any]:
        evidence_records = [record for record in records if getattr(record, 'kind', '') == 'evidence']
        story_records = [record for record in records if getattr(record, 'kind', '') == 'story']
        primary_record = self._select_primary_record(evidence_records or records)
        project = self._project_payload(records)
        org_name = self._extract_org_name(project, records)
        country_raw = _safe_text(project.get('country') or self._best_raw_value(records, ['country', 'projectCountry']))
        country = self._canonical_country(country_raw)
        country_display = self._display_country(country, country_raw, lang)
        location = self._extract_location(primary_record, records, country_display)
        report_date = _format_date(getattr(primary_record, 'raw', {}).get('createdAt') or getattr(primary_record, 'created_at', ''), lang)
        activity = self._extract_activity(primary_record, records, lang)
        title = self._report_title(activity, location, country_display, lang)
        sectors = self._detect_sectors(records, activity, project_name, lang)
        primary_quantity = self._extract_primary_quantity(records, sectors, lang)
        media_files = self._all_media(records)
        image_files = [path for path in media_files if path.suffix.lower() in IMAGE_EXTENSIONS]
        hero_image = self._select_hero_image(image_files, primary_record)
        annex_images = [path for path in image_files if hero_image is None or path != hero_image]
        gps = self._extract_gps(primary_record, records)
        map_point = gps or self._fallback_point(country, location)
        description = _safe_text(getattr(primary_record, 'description', '') or getattr(primary_record, 'raw', {}).get('description'))
        summary = self._narrative(activity, project_name, description, primary_quantity, location, lang)
        highlights = self._highlights(activity, project_code, description, primary_quantity, len(media_files), lang)

        return {
            'project_code': project_code,
            'project_name': project_name,
            'org_name': org_name,
            'country': country,
            'country_display': country_display,
            'location': location,
            'report_date': report_date,
            'activity': activity,
            'title': title,
            'sectors': sectors,
            'quantity': primary_quantity,
            'media_count': len(media_files),
            'evidence_count': len(evidence_records),
            'story_count': len(story_records),
            'hero_image': hero_image,
            'annex_images': annex_images,
            'map_point': map_point,
            'gps_available': gps is not None,
            'map_location_label': map_point[2] if map_point else location,
            'description': description,
            'summary': summary,
            'highlights': highlights,
            'primary_record': primary_record,
            'records': records,
        }

    def _configure_document(self, doc: Document) -> None:
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11.69)
        section.page_height = Inches(8.27)
        section.top_margin = Inches(0.12)
        section.bottom_margin = Inches(0.12)
        section.left_margin = Inches(0.16)
        section.right_margin = Inches(0.16)
        section.header_distance = Inches(0.0)
        section.footer_distance = Inches(0.0)
        styles = doc.styles
        styles['Normal'].font.name = 'Aptos'
        styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Aptos')
        styles['Normal'].font.size = Pt(8.6)
        styles['Normal'].font.color.rgb = _hex_color(DARK)

    def _build_first_page(self, doc: Document, data: dict[str, Any], assets: dict[str, Path], lang: str) -> None:
        self._add_header(doc, data, assets, lang)
        body = doc.add_table(rows=1, cols=3)
        body.alignment = WD_TABLE_ALIGNMENT.CENTER
        body.autofit = False
        _clear_table_borders(body)
        widths = [3650, 4450, 4050]
        for idx, cell in enumerate(body.rows[0].cells):
            _set_cell_width(cell, widths[idx])
            _set_cell_margins(cell, top=45, bottom=35, start=50, end=50)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        self._overview_card(body.cell(0, 0), data, assets, lang)
        self._kpi_section(body.cell(0, 1), data, assets, lang)
        self._media_sector_column(body.cell(0, 2), data, assets, lang)

        spacer = doc.add_table(rows=1, cols=1)
        _clear_table_borders(spacer)
        cell = spacer.cell(0, 0)
        _set_cell_margins(cell, top=1, bottom=1, start=0, end=0)
        _set_cell_shading(cell, BORDER_LIGHT)

        bottom = doc.add_table(rows=1, cols=2)
        bottom.alignment = WD_TABLE_ALIGNMENT.CENTER
        bottom.autofit = False
        _clear_table_borders(bottom)
        for idx, cell in enumerate(bottom.rows[0].cells):
            _set_cell_width(cell, [5650, 6550][idx])
            _set_cell_margins(cell, top=30, bottom=20, start=50, end=50)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        self._highlights_card(bottom.cell(0, 0), data, assets, lang)
        self._map_card(bottom.cell(0, 1), data, assets, lang)

    def _add_header(self, doc: Document, data: dict[str, Any], assets: dict[str, Path], lang: str) -> None:
        table = doc.add_table(rows=1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        _clear_table_borders(table)
        widths = [10100, 750, 1650]
        for idx, cell in enumerate(table.rows[0].cells):
            _set_cell_width(cell, widths[idx])
            _set_cell_shading(cell, BLUE_DARK)
            _set_cell_margins(cell, top=58, bottom=48, start=115, end=115)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        left, mid, right = table.rows[0].cells
        for _cell in (left, mid, right):
            _clear_cell_content(_cell)
        p = _first_paragraph(left)
        _configure_paragraph(p, after=3, line=1.0)
        title_run = _add_run(p, data['title'].upper(), bold=True, size=13.2, color='FFFFFF')
        title_run.font.name = 'Arial Narrow'
        try:
            title_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial Narrow')
        except Exception:
            pass
        p = left.add_paragraph()
        _configure_paragraph(p, after=0, line=1.0)
        subtitle_label = 'Projet : ' if lang == 'fr' else 'Project: '
        _add_run(p, subtitle_label, bold=True, size=8.5, color='FFFFFF')
        _add_run(p, f"{data['project_name']} ({data['project_code']})", size=8.5, color='FFFFFF')

        if assets.get('country_header'):
            p = _first_paragraph(mid)
            _configure_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER)
            try:
                p.add_run().add_picture(str(assets['country_header']), width=Inches(0.72))
            except Exception:
                _add_run(p, data['country_display'], bold=True, color='FFFFFF', size=9)

        _set_cell_border(right, left={'val': 'single', 'sz': '8', 'color': 'D6E6FF'})
        first_line = True
        for icon, text in [('user', data['org_name']), ('pin', data['location']), ('calendar', data['report_date'])]:
            line = _first_paragraph(right) if first_line else right.add_paragraph()
            first_line = False
            _configure_paragraph(line, after=1, line=1.0)
            _add_run(line, self._small_symbol(icon), size=8.0, color='FFFFFF')
            _add_run(line, f'  {text}', size=8.1, color='FFFFFF')

    def _overview_card(self, cell, data: dict[str, Any], assets: dict[str, Path], lang: str) -> None:
        _clear_cell_content(cell)
        _set_cell_shading(cell, PANEL)
        _set_cell_border(cell, top={'val': 'single', 'sz': '6', 'color': BORDER}, bottom={'val': 'single', 'sz': '6', 'color': BORDER}, left={'val': 'single', 'sz': '6', 'color': BORDER}, right={'val': 'single', 'sz': '6', 'color': BORDER})
        _set_cell_margins(cell, top=70, bottom=45, start=90, end=90)
        self._section_title(cell, 'APERÇU DU PROJET' if lang == 'fr' else 'PROJECT OVERVIEW', assets['icon_document'])
        p = cell.add_paragraph()
        _configure_paragraph(p, after=3, line=1.02)
        _add_run(p, data['summary'], size=7.45, color=DARK)
        rows = [
            ('pin', 'Lieu' if lang == 'fr' else 'Location', data['location']),
            ('people', 'Public cible' if lang == 'fr' else 'Target group', self._target_group_label(data, lang)),
            ('camera', 'Type de preuve' if lang == 'fr' else 'Evidence type', 'Preuve photo' if lang == 'fr' else 'Photo evidence'),
        ]
        for icon_key, label, value in rows:
            self._info_row(cell, assets[f'icon_{icon_key}'], label, value)

    def _kpi_section(self, cell, data: dict[str, Any], assets: dict[str, Path], lang: str) -> None:
        _clear_cell_content(cell)
        self._section_title(cell, 'CHIFFRES CLÉS' if lang == 'fr' else 'KEY FIGURES', assets['icon_barchart'], line_width=3.75)
        p = cell.add_paragraph()
        _configure_paragraph(p, after=1)
        cards = cell.add_table(rows=1, cols=3)
        cards.autofit = False
        _clear_table_borders(cards)
        for c in cards.rows[0].cells:
            _set_cell_width(c, 1400)
            _set_cell_margins(c, top=70, bottom=70, start=45, end=45)
            _set_cell_shading(c, PANEL)
            _set_cell_border(c, top={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, bottom={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, left={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, right={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT})
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        quantity = data['quantity']
        kpis = [
            (assets['icon_people_round'], quantity['value'], quantity['label']),
            (assets['icon_media_round'], str(data['media_count']), 'Médias liés' if lang == 'fr' else 'Linked media'),
            (assets['icon_check_round'], str(data['evidence_count']), 'Preuve consolidée' if lang == 'fr' else 'Consolidated evidence'),
        ]
        for c, (icon_path, value, label) in zip(cards.rows[0].cells, kpis):
            p = _first_paragraph(c)
            _configure_paragraph(p, after=4, align=WD_ALIGN_PARAGRAPH.CENTER)
            try:
                p.add_run().add_picture(str(icon_path), width=Inches(0.39))
            except Exception:
                pass
            p = c.add_paragraph()
            _configure_paragraph(p, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
            _add_run(p, value, bold=True, size=20.5, color=BLUE)
            p = c.add_paragraph()
            _configure_paragraph(p, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
            _add_run(p, '—', bold=True, size=8, color=ORANGE)
            p = c.add_paragraph()
            _configure_paragraph(p, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
            _add_run(p, label, size=7.8, color=DARK)

    def _media_sector_column(self, cell, data: dict[str, Any], assets: dict[str, Path], lang: str) -> None:
        _clear_cell_content(cell)
        if assets.get('hero'):
            p = _first_paragraph(cell)
            _configure_paragraph(p, after=6, align=WD_ALIGN_PARAGRAPH.CENTER)
            p.add_run().add_picture(str(assets['hero']), width=Inches(2.70))
        self._section_title(cell, 'ALIGNEMENT SECTORIEL' if lang == 'fr' else 'SECTOR ALIGNMENT', assets['icon_sector'], line_width=2.65)
        sector_table = cell.add_table(rows=1, cols=max(1, len(data['sectors'][:2])))
        sector_table.autofit = False
        _clear_table_borders(sector_table)
        for idx, sector in enumerate(data['sectors'][:2]):
            c = sector_table.rows[0].cells[idx]
            _set_cell_width(c, 1850)
            _set_cell_margins(c, top=25, bottom=20, start=40, end=40)
            if idx > 0:
                _set_cell_border(c, left={'val': 'single', 'sz': '6', 'color': BORDER})
            p = _first_paragraph(c)
            _configure_paragraph(p, after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
            icon_path = assets.get(f"sector_{sector['key']}") or assets['icon_sector']
            try:
                p.add_run().add_picture(str(icon_path), width=Inches(0.36))
            except Exception:
                pass
            p = c.add_paragraph()
            _configure_paragraph(p, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
            _add_run(p, sector['label'], size=7.4, color=DARK)

    def _highlights_card(self, cell, data: dict[str, Any], assets: dict[str, Path], lang: str) -> None:
        _clear_cell_content(cell)
        _set_cell_shading(cell, PANEL_SOFT)
        _set_cell_border(cell, top={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, bottom={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, left={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, right={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT})
        _set_cell_margins(cell, top=65, bottom=50, start=100, end=90)
        self._section_title(cell, 'FAITS SAILLANTS' if lang == 'fr' else 'HIGHLIGHTS', assets['icon_star'], line_width=3.4)
        for item in data['highlights'][:3]:
            p = cell.add_paragraph()
            _configure_paragraph(p, after=2, line=1.03)
            p.paragraph_format.left_indent = Inches(0.16)
            p.paragraph_format.first_line_indent = Inches(-0.16)
            _add_run(p, '• ', bold=True, size=8.8, color=BLUE)
            _add_run(p, item, size=7.8, color=DARK)

    def _map_card(self, cell, data: dict[str, Any], assets: dict[str, Path], lang: str) -> None:
        _clear_cell_content(cell)
        _set_cell_shading(cell, PANEL_SOFT)
        _set_cell_border(cell, top={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, bottom={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, left={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, right={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT})
        _set_cell_margins(cell, top=55, bottom=45, start=90, end=90)
        inner = cell.add_table(rows=1, cols=2)
        inner.autofit = False
        _clear_table_borders(inner)
        text_cell, map_cell = inner.rows[0].cells
        _set_cell_width(text_cell, 2500)
        _set_cell_width(map_cell, 3550)
        _set_cell_margins(text_cell, top=0, bottom=0, start=0, end=100)
        _set_cell_margins(map_cell, top=0, bottom=0, start=30, end=0)
        text_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        map_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _clear_cell_content(text_cell)
        p = _first_paragraph(text_cell)
        _configure_paragraph(p, after=2, line=1.0)
        try:
            p.add_run().add_picture(str(assets['icon_pin_round']), width=Inches(0.20))
        except Exception:
            pass
        _add_run(p, '  ' + ('LOCALISATION' if lang == 'fr' else 'LOCATION'), bold=True, color=BLUE, size=8.0)
        p = text_cell.add_paragraph()
        _configure_paragraph(p, after=0, line=1.0)
        _add_run(p, '━━━━━━', color=BLUE, size=4)
        p = text_cell.add_paragraph()
        _configure_paragraph(p, before=20, after=0, line=1.0)
        _add_run(p, 'Localisation de l’activité' if lang == 'fr' else 'Activity location', bold=True, color=BLUE, size=7.3)
        p = text_cell.add_paragraph()
        _configure_paragraph(p, after=0)
        _add_run(p, '—', bold=True, color=ORANGE, size=11)
        p = map_cell.paragraphs[0]
        _configure_paragraph(p, after=1, align=WD_ALIGN_PARAGRAPH.CENTER)
        p.add_run().add_picture(str(assets['map']), width=Inches(2.75))
        p = map_cell.add_paragraph()
        _configure_paragraph(p, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        _add_run(p, data['map_location_label'] or data['location'], bold=True, color=BLUE, size=7.5)

    def _build_visual_annex(self, doc: Document, data: dict[str, Any], assets: dict[str, Path], lang: str) -> None:
        doc.add_page_break()
        p = doc.add_paragraph()
        _configure_paragraph(p, after=4, line=1.0)
        _add_run(p, 'ANNEXE VISUELLE' if lang == 'fr' else 'VISUAL ANNEX', bold=True, color=BLUE, size=18)
        p = doc.add_paragraph()
        _configure_paragraph(p, after=8, line=1.0)
        _add_run(p, 'Images complémentaires liées à l’activité, conservées pour appui documentaire et relecture interne.' if lang == 'fr' else 'Additional images linked to the activity, kept for documentation support and internal review.', size=10, color=MUTED)
        annex_assets = self._prepare_annex_images(data['annex_images'], Path(tempfile.mkdtemp(prefix='grantproof_annex_images_')))
        for index in range(0, len(annex_assets), 4):
            if index > 0:
                doc.add_page_break()
            chunk = annex_assets[index:index + 4]
            table = doc.add_table(rows=math.ceil(len(chunk) / 2), cols=2)
            table.autofit = False
            _clear_table_borders(table)
            flat_cells = [c for row in table.rows for c in row.cells]
            for c in flat_cells:
                _set_cell_width(c, 5650)
                _set_cell_margins(c, top=90, bottom=100, start=100, end=100)
                _set_cell_shading(c, PANEL_SOFT)
                _set_cell_border(c, top={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, bottom={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, left={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT}, right={'val': 'single', 'sz': '6', 'color': BORDER_LIGHT})
            for c, image_path in zip(flat_cells, chunk):
                p = c.paragraphs[0]
                _configure_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
                try:
                    p.add_run().add_picture(str(image_path), width=Inches(2.7))
                except Exception:
                    _add_run(p, image_path.name, size=9, color=MUTED)
                p = c.add_paragraph()
                _configure_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
                _add_run(p, image_path.name, italic=True, size=8.2, color=MUTED)
            for c in flat_cells[len(chunk):]:
                c.text = ''

    def _section_title(self, cell, title: str, icon_path: Path | None = None, line_width: float = 2.55) -> None:
        title_table = cell.add_table(rows=1, cols=2)
        title_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        title_table.autofit = False
        _clear_table_borders(title_table)
        icon_cell, text_cell = title_table.rows[0].cells
        _set_cell_width(icon_cell, 440)
        _set_cell_width(text_cell, 2400)
        _set_cell_margins(icon_cell, top=0, bottom=0, start=0, end=60)
        _set_cell_margins(text_cell, top=0, bottom=0, start=0, end=0)
        if icon_path:
            p = _first_paragraph(icon_cell)
            _configure_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER)
            try:
                p.add_run().add_picture(str(icon_path), width=Inches(0.21))
            except Exception:
                pass
        p = _first_paragraph(text_cell)
        _configure_paragraph(p, after=0, line=1.0)
        _add_run(p, title, bold=True, color=BLUE, size=8.0)
        p = text_cell.add_paragraph()
        _configure_paragraph(p, after=2, line=1.0)
        _add_run(p, '━' * max(2, int(line_width * 4)), color=BLUE, size=4)

    def _info_row(self, cell, icon_path: Path, label: str, value: str) -> None:
        table = cell.add_table(rows=1, cols=2)
        table.autofit = False
        _clear_table_borders(table)
        icon_cell, text_cell = table.rows[0].cells
        _set_cell_width(icon_cell, 500)
        _set_cell_width(text_cell, 2800)
        for c in (icon_cell, text_cell):
            _set_cell_margins(c, top=12, bottom=12, start=0, end=12)
            _set_cell_border(c, top={'val': 'single', 'sz': '4', 'color': BORDER_LIGHT})
        p = _first_paragraph(icon_cell)
        _configure_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER)
        try:
            p.add_run().add_picture(str(icon_path), width=Inches(0.17))
        except Exception:
            pass
        p = _first_paragraph(text_cell)
        _configure_paragraph(p, after=0, line=1.0)
        _add_run(p, label, bold=True, color=BLUE, size=7.1)
        p = text_cell.add_paragraph()
        _configure_paragraph(p, after=0, line=1.0)
        _add_run(p, value, color=DARK, size=6.9)

    def _create_visual_assets(self, data: dict[str, Any], temp_dir: Path, lang: str) -> dict[str, Path]:
        assets: dict[str, Path] = {}
        for key in ['document', 'barchart', 'pin_round', 'pin', 'people', 'camera', 'media_round', 'people_round', 'check_round', 'sector', 'star']:
            path = temp_dir / f'{key}.png'
            self._draw_icon(key, path)
            assets[f'icon_{key}'] = path
        for sector in data['sectors']:
            path = temp_dir / f"sector_{sector['key']}.png"
            self._draw_sector_icon(sector['icon'], path)
            assets[f"sector_{sector['key']}"] = path
        assets['country_header'] = self._render_country_header(data, temp_dir / 'country_header.png')
        assets['map'] = self._render_map(data, temp_dir / 'map.png', lang)
        if data['hero_image']:
            assets['hero'] = self._prepare_hero_image(data['hero_image'], temp_dir / 'hero.png')
        return assets

    def _draw_icon(self, key: str, path: Path) -> None:
        im = Image.new('RGBA', (160, 160), (255, 255, 255, 0))
        d = ImageDraw.Draw(im)
        blue = (10, 87, 181, 255)
        pale = (239, 246, 255, 255)
        orange = (245, 109, 40, 255)
        if key in {'document', 'pin_round', 'people_round', 'media_round', 'check_round', 'sector', 'star'}:
            d.ellipse((10, 10, 150, 150), fill=pale if key not in {'pin_round', 'sector', 'star'} else blue)
        if key in {'document'}:
            d.rectangle((58, 42, 108, 118), outline=blue, width=7)
            d.line((68, 62, 98, 62), fill=blue, width=5)
            d.line((68, 80, 98, 80), fill=blue, width=5)
            d.line((68, 98, 90, 98), fill=blue, width=5)
        elif key == 'barchart':
            d.rectangle((35, 92, 52, 128), fill=blue)
            d.rectangle((66, 70, 83, 128), fill=blue)
            d.rectangle((97, 45, 114, 128), fill=blue)
        elif key in {'pin', 'pin_round'}:
            color = blue if key == 'pin' else (255, 255, 255, 255)
            d.ellipse((55, 35, 105, 85), outline=color, width=8)
            d.polygon([(80, 125), (55, 75), (105, 75)], outline=color, fill=None)
            d.ellipse((71, 54, 89, 72), fill=color)
        elif key in {'people', 'people_round'}:
            color = blue
            d.ellipse((61, 38, 99, 76), outline=color, width=6)
            d.arc((44, 72, 116, 140), 190, 350, fill=color, width=8)
            d.ellipse((34, 56, 64, 86), outline=color, width=5)
            d.ellipse((96, 56, 126, 86), outline=color, width=5)
        elif key in {'camera'}:
            d.rounded_rectangle((38, 58, 122, 112), radius=10, outline=blue, width=7)
            d.rectangle((56, 46, 84, 62), fill=blue)
            d.ellipse((66, 70, 94, 98), outline=blue, width=6)
        elif key == 'media_round':
            d.rounded_rectangle((45, 55, 105, 105), radius=7, outline=blue, width=6)
            d.rounded_rectangle((58, 42, 118, 92), radius=7, outline=blue, width=5)
            d.polygon([(77, 59), (77, 82), (98, 70)], fill=blue)
        elif key == 'check_round':
            d.rounded_rectangle((48, 46, 112, 118), radius=7, outline=blue, width=6)
            d.line((62, 82, 76, 96, 101, 65), fill=blue, width=7)
            d.rectangle((66, 36, 94, 52), outline=blue, width=5)
        elif key == 'sector':
            d.ellipse((40, 40, 120, 120), outline=(255, 255, 255, 255), width=7)
            d.line((80, 26, 80, 134), fill=(255, 255, 255, 255), width=5)
            d.line((26, 80, 134, 80), fill=(255, 255, 255, 255), width=5)
            d.ellipse((72, 72, 88, 88), fill=orange)
        elif key == 'star':
            pts = []
            for i in range(10):
                angle = -math.pi / 2 + i * math.pi / 5
                r = 48 if i % 2 == 0 else 22
                pts.append((80 + r * math.cos(angle), 80 + r * math.sin(angle)))
            d.line(pts + [pts[0]], fill=(255, 255, 255, 255), width=7, joint='curve')
        im.save(path)

    def _draw_sector_icon(self, key: str, path: Path) -> None:
        im = Image.new('RGBA', (180, 180), (255, 255, 255, 0))
        d = ImageDraw.Draw(im)
        blue = (10, 87, 181, 255)
        d.ellipse((12, 12, 168, 168), fill=(246, 250, 255, 255), outline=(217, 228, 242, 255), width=2)
        if key == 'shelter':
            d.polygon([(45, 92), (90, 50), (135, 92)], outline=blue, fill=None)
            d.rectangle((58, 92, 122, 132), outline=blue, width=7)
            d.rectangle((82, 105, 98, 132), outline=blue, width=5)
        elif key in {'food', 'nutrition'}:
            d.arc((46, 76, 134, 140), 0, 180, fill=blue, width=7)
            d.line((52, 108, 128, 108), fill=blue, width=7)
            for x in (68, 88, 108):
                d.ellipse((x, 56, x + 16, 72), outline=blue, width=4)
                d.line((x + 8, 72, x + 8, 90), fill=blue, width=4)
        elif key == 'agriculture':
            d.line((90, 130, 90, 58), fill=blue, width=7)
            d.ellipse((58, 78, 90, 104), outline=blue, width=6)
            d.ellipse((90, 63, 126, 92), outline=blue, width=6)
            d.line((55, 132, 125, 132), fill=blue, width=6)
        elif key == 'wash':
            d.polygon([(90, 40), (122, 92), (90, 130), (58, 92)], outline=blue, width=7)
            d.line((70, 95, 110, 95), fill=blue, width=5)
        elif key == 'health':
            d.rectangle((78, 42, 102, 136), fill=blue)
            d.rectangle((42, 78, 138, 102), fill=blue)
        elif key == 'education':
            d.polygon([(42, 74), (90, 48), (138, 74), (90, 100)], outline=blue, width=6)
            d.line((58, 88, 58, 116), fill=blue, width=5)
            d.arc((64, 90, 116, 138), 0, 180, fill=blue, width=5)
        elif key == 'protection':
            d.polygon([(90, 42), (132, 60), (124, 112), (90, 138), (56, 112), (48, 60)], outline=blue, width=6)
            d.line((72, 88, 86, 104, 112, 70), fill=blue, width=7)
        elif key == 'cash':
            d.rounded_rectangle((45, 62, 135, 116), radius=8, outline=blue, width=7)
            d.ellipse((76, 70, 104, 108), outline=blue, width=5)
            d.line((90, 45, 90, 135), fill=blue, width=5)
        else:
            d.ellipse((50, 50, 130, 130), outline=blue, width=7)
            d.line((90, 50, 90, 130), fill=blue, width=5)
            d.line((50, 90, 130, 90), fill=blue, width=5)
        im.save(path)

    def _prepare_hero_image(self, source: Path, output: Path) -> Path:
        try:
            img = Image.open(source).convert('RGB')
        except Exception:
            img = Image.new('RGB', (900, 520), '#F0F4FA')
        target_w, target_h = 900, 520
        img_ratio = img.width / max(img.height, 1)
        target_ratio = target_w / target_h
        if img_ratio > target_ratio:
            new_h = img.height
            new_w = int(new_h * target_ratio)
            x = (img.width - new_w) // 2
            img = img.crop((x, 0, x + new_w, img.height))
        else:
            new_w = img.width
            new_h = int(new_w / target_ratio)
            y = max(0, (img.height - new_h) // 2)
            img = img.crop((0, y, img.width, y + new_h))
        img = img.resize((target_w, target_h), Image.LANCZOS)
        mask = Image.new('L', (target_w, target_h), 0)
        d = ImageDraw.Draw(mask)
        d.rounded_rectangle((0, 0, target_w, target_h), radius=28, fill=255)
        rgba = img.convert('RGBA')
        rgba.putalpha(mask)
        canvas = Image.new('RGBA', (target_w + 10, target_h + 10), (255, 255, 255, 0))
        canvas.alpha_composite(rgba, (5, 5))
        canvas.save(output)
        return output

    def _prepare_annex_images(self, images: list[Path], temp_dir: Path) -> list[Path]:
        temp_dir.mkdir(parents=True, exist_ok=True)
        out: list[Path] = []
        for idx, source in enumerate(images, start=1):
            try:
                img = Image.open(source).convert('RGB')
                img.thumbnail((900, 620), Image.LANCZOS)
                canvas = Image.new('RGB', (900, 620), 'white')
                canvas.paste(img, ((900 - img.width) // 2, (620 - img.height) // 2))
                path = temp_dir / f'image_{idx:02d}_{source.name}.png'
                canvas.save(path)
                out.append(path)
            except Exception:
                continue
        return out

    def _render_country_header(self, data: dict[str, Any], output: Path) -> Path | None:
        shape = self._country_shape(data['country'])
        if not shape:
            return None
        img = Image.new('RGBA', (360, 240), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        polygons, bbox = self._shape_polygons(shape)
        if not polygons:
            return None
        mapper = self._map_transform(bbox, 320, 200, 20)
        for poly in polygons:
            pts = [mapper(x, y) for x, y in poly]
            if len(pts) >= 3:
                draw.polygon(pts, fill=(255, 255, 255, 240), outline=(255, 255, 255, 255))
        point = data.get('map_point')
        if point:
            lat, lon, _ = point
            px, py = mapper(lon, lat)
            draw.ellipse((px - 7, py - 7, px + 7, py + 7), fill=(245, 109, 40, 255))
        img.save(output)
        return output

    def _render_map(self, data: dict[str, Any], output: Path, lang: str) -> Path:
        width, height = 780, 420
        img = Image.new('RGB', (width, height), '#F8FAFD')
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle((4, 4, width - 5, height - 5), radius=22, fill='#FFFFFF', outline='#D8DEE8', width=3)
        target_shape = self._country_shape(data['country'])
        if not target_shape:
            draw.text((width // 2 - 70, height // 2 - 8), data['country_display'], fill='#7B8492')
            img.save(output)
            return output
        polygons, bbox = self._shape_polygons(target_shape)
        minx, miny, maxx, maxy = bbox
        dx = maxx - minx
        dy = maxy - miny
        bbox_expanded = (minx - dx * 0.45, miny - dy * 0.35, maxx + dx * 0.45, maxy + dy * 0.35)
        mapper = self._map_transform(bbox_expanded, width - 70, height - 60, 35)
        # draw neighbors first
        for sr in self._iter_shapes():
            name = self._record_name(sr.record)
            shape = sr.shape
            sx1, sy1, sx2, sy2 = shape.bbox
            if sx2 < bbox_expanded[0] or sx1 > bbox_expanded[2] or sy2 < bbox_expanded[1] or sy1 > bbox_expanded[3]:
                continue
            shape_polys, _ = self._shape_polygons(shape)
            is_target = name == self._shape_country_name(data['country'])
            fill = '#FFFFFF' if is_target else '#F1F3F6'
            outline = '#D3D9E1' if is_target else '#DEE3EA'
            for poly in shape_polys:
                pts = [mapper(x, y) for x, y in poly]
                if len(pts) >= 3:
                    draw.polygon(pts, fill=fill, outline=outline)
        # labels
        target_label = data['country_display'].upper()
        cx, cy = mapper((minx + maxx) / 2, (miny + maxy) / 2)
        draw.text((cx - 38, cy - 10), target_label, fill='#6B7280')
        labels_added = 0
        for sr in self._iter_shapes():
            name = self._record_name(sr.record)
            if name == self._shape_country_name(data['country']):
                continue
            shape = sr.shape
            sx1, sy1, sx2, sy2 = shape.bbox
            if sx2 < bbox_expanded[0] or sx1 > bbox_expanded[2] or sy2 < bbox_expanded[1] or sy1 > bbox_expanded[3]:
                continue
            if labels_added >= 5:
                break
            lx, ly = mapper((sx1 + sx2) / 2, (sy1 + sy2) / 2)
            draw.text((lx - 22, ly - 6), self._display_country(name, name, lang).upper(), fill='#9AA2AE')
            labels_added += 1
        point = data.get('map_point')
        if point:
            lat, lon, label = point
            px, py = mapper(lon, lat)
            self._draw_pin(draw, px, py, 18)
            draw.text((px + 14, py - 5), label or data['location'], fill='#28384F')
        img.save(output)
        return output

    def _draw_pin(self, draw: ImageDraw.ImageDraw, x: float, y: float, size: int) -> None:
        orange = '#F56D28'
        draw.ellipse((x - size * 0.55, y - size * 1.05, x + size * 0.55, y + size * 0.05), fill=orange)
        draw.polygon([(x, y + size * 0.85), (x - size * 0.42, y - size * 0.05), (x + size * 0.42, y - size * 0.05)], fill=orange)
        draw.ellipse((x - size * 0.2, y - size * 0.7, x + size * 0.2, y - size * 0.3), fill='white')

    def _shape_file(self) -> Path:
        return self.maps_dir / 'naturalearth_lowres.shp'

    def _iter_shapes(self):
        if not self._shape_file().exists():
            return []
        reader = shapefile.Reader(str(self._shape_file()), encoding='latin1')
        return reader.iterShapeRecords()

    def _country_shape(self, country: str):
        target = self._shape_country_name(country)
        for sr in self._iter_shapes():
            if self._record_name(sr.record) == target:
                return sr.shape
        return None

    def _record_name(self, record) -> str:
        try:
            return _safe_text(record['name'])
        except Exception:
            try:
                return _safe_text(record[2])
            except Exception:
                return ''

    def _shape_country_name(self, country: str) -> str:
        return COUNTRY_ALIASES.get(_safe_text(country).lower(), _safe_text(country))

    def _shape_polygons(self, shape) -> tuple[list[list[tuple[float, float]]], tuple[float, float, float, float]]:
        points = shape.points
        parts = list(shape.parts) + [len(points)]
        polygons: list[list[tuple[float, float]]] = []
        for index in range(len(parts) - 1):
            pts = points[parts[index]:parts[index + 1]]
            if len(pts) >= 3:
                polygons.append([(float(x), float(y)) for x, y in pts])
        minx, miny, maxx, maxy = [float(v) for v in shape.bbox]
        return polygons, (minx, miny, maxx, maxy)

    def _map_transform(self, bbox: tuple[float, float, float, float], width: int, height: int, pad: int):
        minx, miny, maxx, maxy = bbox
        scale = min((width - pad * 2) / max(maxx - minx, 0.001), (height - pad * 2) / max(maxy - miny, 0.001))
        offset_x = (width - (maxx - minx) * scale) / 2
        offset_y = (height - (maxy - miny) * scale) / 2

        def mapper(lon: float, lat: float) -> tuple[float, float]:
            x = offset_x + (lon - minx) * scale + pad / 2
            y = offset_y + (maxy - lat) * scale + pad / 2
            return x, y
        return mapper

    def _project_payload(self, records: list[Any]) -> dict[str, Any]:
        for record in records:
            raw = getattr(record, 'raw', {}) or {}
            project = raw.get('project')
            if isinstance(project, dict) and project:
                return project
        return {}

    def _extract_org_name(self, project: dict[str, Any], records: list[Any]) -> str:
        keys = ['organizationName', 'organisationName', 'ngoName', 'ongName', 'clientName', 'implementingPartner', 'partnerName', 'organization', 'organisation', 'ngo']
        for key in keys:
            value = project.get(key)
            if isinstance(value, dict):
                value = value.get('name') or value.get('displayName')
            text = _safe_text(value)
            if text:
                return text
        for record in records:
            raw = getattr(record, 'raw', {}) or {}
            for key in keys:
                value = raw.get(key)
                if isinstance(value, dict):
                    value = value.get('name') or value.get('displayName')
                text = _safe_text(value)
                if text:
                    return text
        return 'Organisation'

    def _best_raw_value(self, records: list[Any], keys: list[str]) -> str:
        for record in records:
            raw = getattr(record, 'raw', {}) or {}
            for key in keys:
                text = _safe_text(raw.get(key))
                if text:
                    return text
        return ''

    def _canonical_country(self, country_raw: str) -> str:
        text = _safe_text(country_raw)
        return COUNTRY_ALIASES.get(text.lower(), text or 'Niger')

    def _display_country(self, canonical: str, raw: str, lang: str) -> str:
        if lang == 'fr':
            return COUNTRY_DISPLAY_FR.get(canonical, _safe_text(raw, canonical))
        return canonical or raw

    def _select_primary_record(self, records: list[Any]) -> Any:
        def score(record: Any) -> tuple[int, int, int, str]:
            raw = getattr(record, 'raw', {}) or {}
            media_count = len([p for p in getattr(record, 'media_files', []) if p.suffix.lower() in IMAGE_EXTENSIONS])
            gps = 1 if raw.get('latitude') is not None and raw.get('longitude') is not None else 0
            desc_len = len(_safe_text(getattr(record, 'description', '') or raw.get('description') or raw.get('summary')))
            date = _safe_text(raw.get('createdAt') or getattr(record, 'created_at', ''))
            return (media_count, gps, desc_len, date)
        return sorted(records, key=score, reverse=True)[0]

    def _extract_location(self, primary_record: Any, records: list[Any], country_display: str) -> str:
        for value in [getattr(primary_record, 'location', ''), getattr(primary_record, 'raw', {}).get('locationLabel')]:
            text = _safe_text(value)
            if text and text != '-':
                if country_display and country_display.lower() not in text.lower():
                    return f'{text}, {country_display}'
                return text
        for record in records:
            text = _safe_text(getattr(record, 'location', '') or getattr(record, 'raw', {}).get('locationLabel'))
            if text and text != '-':
                return f'{text}, {country_display}' if country_display and country_display.lower() not in text.lower() else text
        return country_display

    def _extract_activity(self, primary_record: Any, records: list[Any], lang: str) -> str:
        for value in [getattr(primary_record, 'raw', {}).get('activity'), getattr(primary_record, 'title', '')]:
            text = _clean_spaces(value)
            if text and text != '-':
                return text[:1].upper() + text[1:]
        counts = Counter(_safe_text(getattr(record, 'raw', {}).get('activity')) for record in records if _safe_text(getattr(record, 'raw', {}).get('activity')))
        if counts:
            return counts.most_common(1)[0][0]
        return 'Activité terrain' if lang == 'fr' else 'Field activity'

    def _report_title(self, activity: str, location: str, country: str, lang: str) -> str:
        short_location = location
        if ',' in short_location:
            short_location = ', '.join([part.strip() for part in short_location.split(',')[:2] if part.strip()])
        if lang == 'fr':
            return f"Rapport d’activité : {activity} à {short_location}"
        return f'Activity report: {activity} in {short_location}'

    def _detect_sectors(self, records: list[Any], activity: str, project_name: str, lang: str) -> list[dict[str, str]]:
        corpus = ' '.join([
            activity,
            *[_safe_text(getattr(record, 'title', '')) for record in records],
            *[_safe_text(getattr(record, 'description', '')) for record in records],
            *[_safe_text((getattr(record, 'raw', {}) or {}).get('activity')) for record in records],
            *[_safe_text((getattr(record, 'raw', {}) or {}).get('output')) for record in records],
        ]).lower()
        scored: list[tuple[int, dict[str, str]]] = []
        for rule in SECTOR_RULES:
            score = sum(1 for keyword in rule['keywords'] if keyword.lower() in corpus)
            if score:
                scored.append((score, {'key': rule['key'], 'label': rule[lang], 'icon': rule['icon']}))
        if not scored:
            return [{'key': 'coordination', 'label': 'Coordination' if lang == 'fr' else 'Coordination', 'icon': 'coordination'}]
        scored.sort(key=lambda item: (-item[0], item[1]['label']))
        result: list[dict[str, str]] = []
        seen = set()
        for _, sector in scored:
            if sector['key'] in seen:
                continue
            result.append(sector)
            seen.add(sector['key'])
            if len(result) == 2:
                break
        return result

    def _extract_primary_quantity(self, records: list[Any], sectors: list[dict[str, str]], lang: str) -> dict[str, str]:
        corpus = ' '.join([
            _safe_text(getattr(record, 'description', '')) + ' ' + _safe_text(getattr(record, 'title', '')) + ' ' + _safe_text((getattr(record, 'raw', {}) or {}).get('activity'))
            for record in records
        ])
        for key, pattern, fr_label, en_label in QUANTITY_PATTERNS:
            matches = list(re.finditer(pattern, corpus, flags=re.IGNORECASE))
            if matches:
                number = self._normalize_number(matches[0].group(1))
                label = fr_label if lang == 'fr' else en_label
                if key == 'beneficiary' and sectors and sectors[0]['key'] == 'education':
                    label = 'Participants' if lang == 'fr' else 'Participants'
                return {'value': number, 'label': label}
        return {'value': str(len([r for r in records if getattr(r, 'kind', '') == 'evidence']) or len(records)), 'label': 'Preuves' if lang == 'fr' else 'Evidence items'}

    def _normalize_number(self, value: str) -> str:
        cleaned = re.sub(r'[^0-9]', '', value)
        if not cleaned:
            return value.strip()
        try:
            return f'{int(cleaned):,}'.replace(',', ' ')
        except Exception:
            return cleaned

    def _all_media(self, records: list[Any]) -> list[Path]:
        media: list[Path] = []
        for record in records:
            media.extend([Path(p) for p in getattr(record, 'media_files', [])])
        return media

    def _select_hero_image(self, image_files: list[Path], primary_record: Any) -> Path | None:
        if not image_files:
            return None
        primary_images = [Path(p) for p in getattr(primary_record, 'media_files', []) if Path(p).suffix.lower() in IMAGE_EXTENSIONS]
        candidates = primary_images or image_files
        best = None
        best_score = -1
        for path in candidates:
            try:
                im = Image.open(path)
                w, h = im.size
                score = min(w * h, 4_000_000) + (300_000 if w >= h else 0)
                if score > best_score:
                    best = path
                    best_score = score
            except Exception:
                continue
        return best or candidates[0]

    def _extract_gps(self, primary_record: Any, records: list[Any]) -> tuple[float, float, str] | None:
        candidates = [primary_record, *records]
        for record in candidates:
            raw = getattr(record, 'raw', {}) or {}
            lat = _parse_float(raw.get('latitude') or raw.get('lat'))
            lon = _parse_float(raw.get('longitude') or raw.get('lng') or raw.get('lon'))
            if lat is not None and lon is not None:
                label = _safe_text(raw.get('locationLabel') or getattr(record, 'location', ''), 'Localisation')
                return (lat, lon, label)
        return None

    def _fallback_point(self, country: str, location: str) -> tuple[float, float, str] | None:
        if country in CAPITAL_COORDS:
            lat, lon, capital = CAPITAL_COORDS[country]
            location_lower = location.lower()
            if capital.lower() in location_lower:
                return (lat, lon, capital)
        return None

    def _target_group_label(self, data: dict[str, Any], lang: str) -> str:
        q = data['quantity']
        label = q['label'].lower()
        if 'abri' in label or 'shelter' in label:
            return 'Ménages / communautés bénéficiaires' if lang == 'fr' else 'Beneficiary households / communities'
        if 'école' in label or 'school' in label:
            return 'Élèves, enseignants et communautés scolaires' if lang == 'fr' else 'Students, teachers and school communities'
        if 'bénéficiaire' in label or 'participant' in label:
            return 'Bénéficiaires de la formation' if lang == 'fr' else 'Training beneficiaries'
        return 'Communautés ciblées' if lang == 'fr' else 'Targeted communities'

    def _narrative(self, activity: str, project_name: str, description: str, quantity: dict[str, str], location: str, lang: str) -> str:
        value = quantity['value']
        label = quantity['label'].lower()
        if lang == 'fr':
            if description:
                base = f"Cette activité de {activity.lower()} a été documentée à {location} dans le cadre du projet {project_name}. La description source mentionne {description.strip()}"
                base = _sentence(base)
            else:
                base = f"Cette activité de {activity.lower()} a été documentée à {location} dans le cadre du projet {project_name}."
            support = f"Les éléments collectés permettent d’illustrer la mise en œuvre de l’activité et de soutenir le reporting bailleur, sans aller au-delà des informations disponibles."
            if value and label and value not in base:
                return f"{base} Le dossier met en évidence {value} {label}. {support}"
            return f"{base} {support}"
        if description:
            base = f"This {activity.lower()} activity was documented in {location} under the project {project_name}. The source description states: {description.strip()}"
            base = _sentence(base)
        else:
            base = f"This {activity.lower()} activity was documented in {location} under the project {project_name}."
        support = 'The collected elements illustrate implementation and support donor reporting without going beyond the available information.'
        return f"{base} The file highlights {value} {label}. {support}"

    def _highlights(self, activity: str, project_code: str, description: str, quantity: dict[str, str], media_count: int, lang: str) -> list[str]:
        if lang == 'fr':
            highlights = [f"Une activité de {activity.lower()} a été documentée dans le cadre du projet {project_code}."]
            if quantity['value'] and quantity['label']:
                highlights.append(f"La description source permet de retenir {quantity['value']} {quantity['label'].lower()} comme chiffre clé de l’activité.")
            if media_count > 0:
                highlights.append(f"La preuve est appuyée par {media_count} média(s) lié(s), mobilisables comme support de redevabilité et d’illustration des résultats.")
            return highlights
        highlights = [f"A {activity.lower()} activity was documented under project {project_code}."]
        if quantity['value'] and quantity['label']:
            highlights.append(f"The source description supports {quantity['value']} {quantity['label'].lower()} as the key figure for this activity.")
        if media_count > 0:
            highlights.append(f"The evidence is supported by {media_count} linked media asset(s), available for accountability and results illustration.")
        return highlights

    def _small_symbol(self, kind: str) -> str:
        # Simple symbols render reliably in Word/LibreOffice with Aptos/Segoe UI Symbol fallbacks.
        return {'user': '○', 'pin': '⌖', 'calendar': '□'}.get(kind, '•')
