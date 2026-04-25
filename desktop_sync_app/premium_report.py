from __future__ import annotations

import math
import re
import tempfile
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

import shapefile  # pyshp
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont, ImageOps, ImageStat

IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
VIDEO_EXTENSIONS = {'.mp4', '.mov', '.m4v', '.avi', '.mkv', '.webm'}

BLUE = '064890'
BLUE_2 = '0A57B5'
BLUE_3 = '1665C1'
PALE_BLUE = 'EEF5FF'
SOFT_BLUE = 'F6FAFF'
ORANGE = 'F56D28'
DARK = '1D2D44'
MUTED = '5D6B7C'
BORDER = 'D9E4F2'
BORDER_LIGHT = 'E8EEF7'
WHITE = 'FFFFFF'
LIGHT_GREY = 'F4F6FA'

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
    "côte d'ivoire": "Côte d'Ivoire",
    'cote d ivoire': "Côte d'Ivoire",
    'south sudan': 'S. Sudan',
    'soudan du sud': 'S. Sudan',
    'sudan': 'Sudan',
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
    "Côte d'Ivoire": 'Côte d’Ivoire',
    'S. Sudan': 'Soudan du Sud',
    'Sudan': 'Soudan',
}

HUMANITARIAN_ICON_MAP = {
    'food': 'food_security.png',
    'agriculture': 'early_recovery.png',
    'shelter': 'shelter.png',
    'wash': 'wash.png',
    'health': 'health.png',
    'nutrition': 'nutrition.png',
    'education': 'education.png',
    'protection': 'protection.png',
    'cash': 'logistics.png',
    'coordination': 'cccm.png',
    'people': 'cccm.png',
    'beneficiary': 'cccm.png',
    'group': 'cccm.png',
    'kits': 'logistics.png',
    'box': 'logistics.png',
    'media': 'etc.png',
    'check': 'etc.png',
}

INDICATOR_PATTERNS = [
    ('beneficiary', r'(\d[\d\s.,]*)\s*(?:bénéficiaires?|beneficiaires?|beneficiaries|participants?|personnes?)', 'Bénéficiaires', 'Beneficiaries', 'people', 130),
    ('kit', r'(\d[\d\s.,]*)\s*(?:kits?)', 'Kits distribués', 'Distributed kits', 'kits', 125),
    ('group', r'(\d[\d\s.,]*)\s*(?:groupements?|groupes?|groups?)', 'Groupements', 'Groups', 'group', 120),
    ('shelter', r'(?:construction|construit|construits|réhabilitation|rehabilitation)?\s*(?:de\s*)?(\d[\d\s.,]*)\s*(?:abris?|shelters?)', 'Abris', 'Shelters', 'shelter', 125),
    ('farmer', r'(\d[\d\s.,]*)\s*(?:agriculteurs?|farmers?)', 'Agriculteurs', 'Farmers', 'agriculture', 116),
    ('household', r'(\d[\d\s.,]*)\s*(?:ménages?|menages?|households?)', 'Ménages', 'Households', 'people', 112),
    ('latrine', r'(\d[\d\s.,]*)\s*(?:latrines?)', 'Latrines', 'Latrines', 'wash', 108),
    ('school', r'(\d[\d\s.,]*)\s*(?:écoles?|ecoles?|schools?)', 'Écoles', 'Schools', 'education', 104),
]

SECTOR_RULES = [
    {
        'key': 'shelter', 'fr': 'Abris', 'en': 'Shelter', 'icon': 'shelter',
        'keywords': ['abri', 'abris', 'shelter', 'logement', 'construction', 'réhabilitation', 'rehabilitation', 'maison', 'habitat', 'ame', 'nfi'],
    },
    {
        'key': 'food_security', 'fr': 'Sécurité alimentaire', 'en': 'Food security', 'icon': 'food',
        'keywords': ['sécurité alimentaire', 'food security', 'vivres', 'food', 'ration', 'maraîcher', 'maraicher', 'semence', 'intrants', 'kit agricole', 'kits agricole', 'kits agricoles', 'agricole', 'cash for food'],
    },
    {
        'key': 'agriculture', 'fr': 'Agriculture', 'en': 'Agriculture', 'icon': 'agriculture',
        'keywords': ['agriculture', 'agricole', 'agricoles', 'maraîcher', 'maraicher', 'semence', 'irrigation', 'élevage', 'livelihood', 'moyens d’existence', 'moyens existence', 'agriculteur', 'agriculteurs'],
    },
    {
        'key': 'wash', 'fr': 'Eau, hygiène et assainissement', 'en': 'WASH', 'icon': 'wash',
        'keywords': ['wash', 'eha', 'eau', 'hygiène', 'hygiene', 'assainissement', 'latrine', 'chloration', 'forage', 'water', 'sanitation'],
    },
    {
        'key': 'health', 'fr': 'Santé', 'en': 'Health', 'icon': 'health',
        'keywords': ['santé', 'health', 'clinique', 'centre de santé', 'médicament', 'vaccination', 'choléra', 'cholera', 'soins'],
    },
    {
        'key': 'nutrition', 'fr': 'Nutrition', 'en': 'Nutrition', 'icon': 'nutrition',
        'keywords': ['nutrition', 'malnutrition', 'anjé', 'anje', 'mas', 'mam', 'dépistage', 'dépister'],
    },
    {
        'key': 'education', 'fr': 'Éducation', 'en': 'Education', 'icon': 'education',
        'keywords': ['éducation', 'education', 'école', 'ecole', 'enseignant', 'classe', 'élève', 'scolaire', 'school'],
    },
    {
        'key': 'protection', 'fr': 'Protection', 'en': 'Protection', 'icon': 'protection',
        'keywords': ['protection', 'vbg', 'gbv', 'violence', 'enfant', 'psychosocial', 'psea', 'réunification', 'protection de l’enfance'],
    },
    {
        'key': 'cash', 'fr': 'Transferts monétaires', 'en': 'Cash assistance', 'icon': 'cash',
        'keywords': ['cash', 'transfert monétaire', 'transferts monétaires', 'voucher', 'coupon', 'assistance monétaire'],
    },
    {
        'key': 'coordination', 'fr': 'Coordination', 'en': 'Coordination', 'icon': 'coordination',
        'keywords': ['coordination', 'réunion', 'meeting', 'cluster', 'atelier', 'planification'],
    },
]

QUANTITY_PATTERNS = [
    ('shelter', r'(?:construction|construit|construits|réhabilitation|rehabilitation)?\s*(?:de\s*)?(\d[\d\s.,]*)\s*(abris?|shelters?)', 'Abris', 'Shelters'),
    ('kit', r'(\d[\d\s.,]*)\s*(kits?)', 'Kits', 'Kits'),
    ('kit_context', r'(?:kit|kits)[^\d]{0,35}(\d[\d\s.,]*)', 'Kits', 'Kits'),
    ('farmer', r'(\d[\d\s.,]*)\s*(agriculteurs?|farmers?)', 'Agriculteurs', 'Farmers'),
    ('latrine', r'(\d[\d\s.,]*)\s*(latrines?)', 'Latrines', 'Latrines'),
    ('school', r'(\d[\d\s.,]*)\s*(écoles?|ecoles?|schools?)', 'Écoles', 'Schools'),
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


def _rgb(hex_color: str) -> RGBColor:
    return RGBColor.from_string(hex_color.strip().lstrip('#').upper())


def _pil_rgb(hex_color: str) -> tuple[int, int, int]:
    h = hex_color.strip().lstrip('#')
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _norm_lower(text: str) -> str:
    return _clean_spaces(text).lower()


def resource_path(relative: str) -> Path:
    candidates = [
        Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent.parent)) / relative,
        Path(__file__).resolve().parent / relative,
        Path(__file__).resolve().parent / relative.replace('desktop_sync_app/', ''),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


class PremiumActivityReportBuilder:
    """Editable premium DOCX activity report.

    Page 1 is built with real Word tables, paragraphs and images so the NGO can edit text.
    Maps, photos and pictograms remain images, but all title/narrative/KPI/caption text is editable.
    """

    def __init__(self, base_folder: Path, default_org_name: str = '') -> None:
        self.base_folder = Path(base_folder)
        self.default_org_name = _safe_text(default_org_name)
        self.assets_dir = resource_path('desktop_sync_app/assets')
        self.maps_dir = self.assets_dir / 'maps' / 'naturalearth_lowres'
        self._shape_cache: list[Any] | None = None
        self._icon_cache: dict[str, Path] = {}
        self._temp_dir: Path | None = None

    def build(self, output: Path, project_code: str, project_name: str, records: list[Any], lang: str) -> None:
        lang = 'en' if str(lang).lower().startswith('en') else 'fr'
        if not records:
            return
        data = self._prepare_data(project_code, project_name, records, lang)
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix='grantproof_report_assets_') as tmp:
            self._temp_dir = Path(tmp)
            self._icon_cache = {}
            data['hero_picture'] = self._prepare_photo(data.get('hero_image'), self._temp_dir / 'hero.png', (1040, 620))
            data['map_picture'] = self._render_map_image(data, self._temp_dir / 'map.png', (880, 500), lang)
            data['header_country_picture'] = self._render_country_silhouette(data, self._temp_dir / 'country.png', (260, 170), for_header=True)
            doc = Document()
            self._configure_docx(doc)
            self._build_page_one(doc, data, lang)
            self._build_annex_pages(doc, data, lang)
            doc.save(output)
            self._temp_dir = None
            self._icon_cache = {}

    # ------------------------------------------------------------------ data

    def _prepare_data(self, project_code: str, project_name: str, records: list[Any], lang: str) -> dict[str, Any]:
        evidence_records = [record for record in records if getattr(record, 'kind', '') == 'evidence']
        primary_record = self._select_primary_record(evidence_records or records)
        project = self._project_payload(records)
        org_name = self._extract_org_name(project, records)
        country_raw = _safe_text(project.get('country') or self._best_raw_value(records, ['country', 'projectCountry']))
        country = self._canonical_country(country_raw)
        country_display = self._display_country(country, country_raw, lang)
        location = self._extract_location(primary_record, records, country_display)
        report_date = _format_date(getattr(primary_record, 'raw', {}).get('createdAt') or getattr(primary_record, 'created_at', ''), lang)
        activity_raw = self._extract_activity(primary_record, records, lang)
        activity = self._polish_activity(activity_raw, lang)
        sectors = self._detect_sectors(records, activity, project_name, lang)
        media_files = self._all_media(records)
        image_files = [path for path in media_files if Path(path).suffix.lower() in IMAGE_EXTENSIONS]
        hero_image = self._select_hero_image(image_files, primary_record, records)
        annex_images = [Path(path) for path in image_files if hero_image is None or Path(path) != Path(hero_image)]
        gps = self._extract_gps(primary_record, records)
        description = self._polish_source_description(_safe_text(getattr(primary_record, 'description', '') or getattr(primary_record, 'raw', {}).get('description')))
        subtype = self._evidence_type_label(primary_record, lang)
        title = self._report_title(activity, location, country_display, lang)
        indicators = self._extract_indicators(records, lang)
        quantity = indicators[0] if indicators else {'value': str(len(evidence_records) or len(records)), 'label': 'Preuves' if lang == 'fr' else 'Evidence items', 'icon': 'check', 'kind': 'evidence'}
        summary = self._narrative(activity, project_name, description, indicators, location, sectors, lang)
        highlights = self._highlights(activity, project_code, description, indicators, len(media_files), location, sectors, lang)
        target_group = self._target_group_label(quantity, sectors, description, lang)
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
            'quantity': quantity,
            'kpis': indicators[:4],
            'media_count': len(media_files),
            'evidence_count': max(1, len(evidence_records) or len(records)),
            'hero_image': hero_image,
            'annex_images': annex_images,
            'gps_point': gps,
            'description': description,
            'evidence_type': subtype,
            'summary': summary,
            'highlights': highlights,
            'target_group': target_group,
            'map_location_label': gps[2] if gps else country_display,
        }

    def _configure_docx(self, doc: Document) -> None:
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11)
        section.page_height = Inches(8.25)
        section.top_margin = Inches(0.12)
        section.bottom_margin = Inches(0.10)
        section.left_margin = Inches(0.12)
        section.right_margin = Inches(0.12)
        section.header_distance = Inches(0)
        section.footer_distance = Inches(0)
        styles = doc.styles
        styles['Normal'].font.name = 'Arial'
        styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
        styles['Normal'].font.size = Pt(8.5)
        styles['Normal'].font.color.rgb = _rgb(DARK)

    def _build_page_one(self, doc: Document, data: dict[str, Any], lang: str) -> None:
        self._build_header(doc, data, lang)
        self._spacer(doc, 0.8)
        top = doc.add_table(rows=1, cols=3)
        self._table_no_borders(top)
        top.alignment = WD_TABLE_ALIGNMENT.CENTER
        top.autofit = False
        self._set_row_height(top.rows[0], Inches(3.55))
        widths = [3.20, 4.15, 3.42]
        for cell, width in zip(top.rows[0].cells, widths):
            self._set_cell_width(cell, Inches(width))
            self._cell_margins(cell, top=18, bottom=12, start=34, end=34)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        self._build_overview(top.cell(0, 0), data, lang)
        self._build_kpis(top.cell(0, 1), data, lang)
        self._build_media_sector(top.cell(0, 2), data, lang)
        self._spacer(doc, 1.0)
        bottom = doc.add_table(rows=1, cols=2)
        self._table_no_borders(bottom)
        bottom.alignment = WD_TABLE_ALIGNMENT.CENTER
        bottom.autofit = False
        self._set_row_height(bottom.rows[0], Inches(2.05))
        for cell, width in zip(bottom.rows[0].cells, [5.10, 5.67]):
            self._set_cell_width(cell, Inches(width))
            self._cell_margins(cell, top=18, bottom=14, start=36, end=36)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            self._shade_cell(cell, SOFT_BLUE)
            self._border_cell(cell, BORDER_LIGHT)
        self._build_highlights(bottom.cell(0, 0), data, lang)
        self._build_location(bottom.cell(0, 1), data, lang)

    def _build_header(self, doc: Document, data: dict[str, Any], lang: str) -> None:
        table = doc.add_table(rows=1, cols=3)
        self._table_no_borders(table)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        self._set_row_height(table.rows[0], Inches(0.78))
        for cell, width in zip(table.rows[0].cells, [8.25, 0.92, 1.60]):
            self._set_cell_width(cell, Inches(width))
            self._shade_cell(cell, BLUE)
            self._cell_margins(cell, top=20, bottom=16, start=46, end=36)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        title_cell, map_cell, meta_cell = table.rows[0].cells
        p = self._cell_p(title_cell)
        self._pconf(p, after=0, line=0.92)
        title_size = 11.3 if len(data['title']) > 56 else 12.0
        self._run(p, data['title'].upper(), bold=True, size=title_size, color=WHITE)
        p = title_cell.add_paragraph()
        self._pconf(p, after=0, line=0.92)
        prefix = 'Projet' if lang == 'fr' else 'Project'
        self._run(p, f"{prefix} : {data['project_name']} ({data['project_code']})", size=8.0, color=WHITE)
        p = self._cell_p(map_cell)
        self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
        if Path(data['header_country_picture']).exists():
            p.add_run().add_picture(str(data['header_country_picture']), width=Inches(0.76))
        meta = meta_cell.add_table(rows=3, cols=2)
        self._table_no_borders(meta)
        meta.autofit = False
        for row in meta.rows:
            self._set_row_height(row, Inches(0.18))
        values = [
            ('org', self._truncate_text(data['org_name'], 22)),
            ('pin', self._truncate_text(data['location'], 22)),
            ('calendar', data['report_date']),
        ]
        for i, (kind, value) in enumerate(values):
            left = meta.cell(i, 0)
            right = meta.cell(i, 1)
            self._set_cell_width(left, Inches(0.20))
            self._set_cell_width(right, Inches(1.26))
            self._cell_margins(left, top=0, bottom=0, start=0, end=10)
            self._cell_margins(right, top=0, bottom=0, start=0, end=0)
            p = self._cell_p(left)
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0, line=0.9)
            p.add_run().add_picture(str(self._icon(kind)), width=Inches(0.11))
            p = self._cell_p(right)
            self._pconf(p, after=0, line=0.9)
            self._run(p, value, size=7.2, color=WHITE)

    def _build_overview(self, cell, data: dict[str, Any], lang: str) -> None:
        self._shade_cell(cell, WHITE)
        self._border_cell(cell, BORDER)
        self._section_title(cell, 'APERÇU DU PROJET' if lang == 'fr' else 'PROJECT OVERVIEW', 'doc')
        p = cell.add_paragraph()
        self._pconf(p, line=1.08, after=4)
        self._run(p, data['summary'], size=6.7)
        rows = [
            ('pin', 'Lieu' if lang == 'fr' else 'Location', data['location']),
            ('people', 'Public cible' if lang == 'fr' else 'Target group', data['target_group']),
            ('camera', 'Type de preuve' if lang == 'fr' else 'Evidence type', data['evidence_type']),
        ]
        for icon, label, value in rows:
            item = cell.add_table(rows=1, cols=2)
            self._table_no_borders(item)
            item.autofit = False
            self._set_cell_width(item.cell(0, 0), Inches(0.34))
            self._set_cell_width(item.cell(0, 1), Inches(2.55))
            self._cell_margins(item.cell(0, 0), top=10, bottom=4, start=0, end=18)
            self._cell_margins(item.cell(0, 1), top=10, bottom=4, start=0, end=0)
            p = self._cell_p(item.cell(0, 0))
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER)
            p.add_run().add_picture(str(self._icon(icon)), width=Inches(0.22))
            p = self._cell_p(item.cell(0, 1))
            self._pconf(p, line=1.0, after=0)
            self._run(p, label, bold=True, size=6.8, color=BLUE_2)
            p = item.cell(0, 1).add_paragraph()
            self._pconf(p, line=1.0, after=0)
            self._run(p, self._truncate_text(value, 48), size=6.7)

    def _build_kpis(self, cell, data: dict[str, Any], lang: str) -> None:
        self._section_title(cell, 'CHIFFRES CLÉS' if lang == 'fr' else 'KEY FIGURES', 'bar')
        kpis = list(data.get('kpis') or [])
        normalized = []
        seen = set()
        preferred_order = ['beneficiaries', 'people', 'kits', 'households', 'sessions', 'groups', 'media', 'evidence']
        while len(kpis) < 3:
            kinds = {item.get('kind') for item in kpis}
            if 'media' not in kinds:
                kpis.append({'kind': 'media', 'value': str(data['media_count']), 'label': 'Médias liés' if lang == 'fr' else 'Linked media', 'icon': 'media'})
            elif 'evidence' not in kinds:
                kpis.append({'kind': 'evidence', 'value': str(data['evidence_count']), 'label': 'Preuve consolidée' if lang == 'fr' else 'Consolidated evidence', 'icon': 'check'})
            else:
                break
        def rank(item):
            kind = item.get('kind', '')
            try:
                priority = preferred_order.index(kind)
            except ValueError:
                priority = 99
            return priority, -len(str(item.get('value', '')))
        for item in sorted(kpis, key=rank):
            sig = (item.get('kind'), item.get('value'))
            if sig in seen:
                continue
            normalized.append(item)
            seen.add(sig)
            if len(normalized) == 3:
                break
        grid = cell.add_table(rows=1, cols=3)
        self._table_no_borders(grid)
        grid.autofit = False
        self._set_row_height(grid.rows[0], Inches(2.30))
        for kcell in grid.rows[0].cells:
            self._set_cell_width(kcell, Inches(1.28))
            self._cell_margins(kcell, top=14, bottom=10, start=10, end=10)
            self._shade_cell(kcell, WHITE)
            self._border_cell(kcell, BORDER_LIGHT)
            kcell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for idx, kpi in enumerate(normalized[:3]):
            kcell = grid.cell(0, idx)
            p = self._cell_p(kcell)
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
            p.add_run().add_picture(str(self._icon(kpi.get('icon', 'people'))), width=Inches(0.27))
            p = kcell.add_paragraph()
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0, line=0.9)
            self._run(p, self._truncate_text(str(kpi.get('value', '')), 14), bold=True, size=14.8, color=BLUE_2)
            p = kcell.add_paragraph()
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0, line=0.9)
            self._run(p, '—', bold=True, size=6.8, color=ORANGE)
            p = kcell.add_paragraph()
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0, line=0.88)
            self._run(p, self._truncate_text(str(kpi.get('label', '')), 22), size=6.4)

    def _build_media_sector(self, cell, data: dict[str, Any], lang: str) -> None:
        p = self._cell_p(cell)
        self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
        if Path(data['hero_picture']).exists():
            p.add_run().add_picture(str(data['hero_picture']), width=Inches(3.05), height=Inches(1.95))
        self._section_title(cell, 'ALIGNEMENT SECTORIEL' if lang == 'fr' else 'SECTOR ALIGNMENT', 'sector')
        sectors = data['sectors'][:2] or [{'key': 'coordination', 'label': 'Coordination', 'icon': 'coordination'}]
        table = cell.add_table(rows=1, cols=max(1, len(sectors)))
        self._table_no_borders(table)
        table.autofit = False
        for scell, sector in zip(table.rows[0].cells, sectors):
            self._set_cell_width(scell, Inches(1.50))
            self._cell_margins(scell, top=4, bottom=0, start=6, end=6)
            p = self._cell_p(scell)
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=1)
            p.add_run().add_picture(str(self._sector_icon(sector['icon'])), width=Inches(0.46))
            p = scell.add_paragraph()
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, line=0.95, after=0)
            self._run(p, sector['label'], size=6.8)

    def _build_highlights(self, cell, data: dict[str, Any], lang: str) -> None:
        self._section_title(cell, 'FAITS SAILLANTS' if lang == 'fr' else 'HIGHLIGHTS', 'star', line_chars='━━━━━━━━━━━━━━━━━━━━')
        for item in data['highlights'][:3]:
            p = cell.add_paragraph()
            self._pconf(p, line=1.05, after=2)
            p.paragraph_format.left_indent = Inches(0.16)
            p.paragraph_format.first_line_indent = Inches(-0.16)
            self._run(p, '• ', bold=True, color=BLUE_2, size=8.4)
            self._run(p, item, size=6.8)

    def _build_location(self, cell, data: dict[str, Any], lang: str) -> None:
        layout = cell.add_table(rows=1, cols=2)
        self._table_no_borders(layout)
        layout.autofit = False
        left, right = layout.rows[0].cells
        self._set_cell_width(left, Inches(1.95))
        self._set_cell_width(right, Inches(3.50))
        self._cell_margins(left, top=0, bottom=0, start=0, end=24)
        self._cell_margins(right, top=0, bottom=0, start=0, end=0)
        self._section_title(left, 'LOCALISATION' if lang == 'fr' else 'LOCATION', 'pin')
        p = left.add_paragraph()
        self._pconf(p, before=4, after=1)
        self._run(p, 'Localisation de l’activité' if lang == 'fr' else 'Activity location', bold=True, size=6.9, color=BLUE_2)
        p = left.add_paragraph()
        self._pconf(p, after=0, line=0.95)
        self._run(p, self._truncate_text(data['location'], 34), size=6.6, color=DARK)
        p = left.add_paragraph()
        self._pconf(p, after=0, line=0.95)
        if data.get('gps_point'):
            lat, lon, _ = data['gps_point']
            self._run(p, f'GPS : {lat:.4f}, {lon:.4f}', size=6.1, color=MUTED)
        else:
            self._run(p, 'Carte nationale affichée par défaut.', size=6.1, color=MUTED)
        p = self._cell_p(right)
        self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
        p.add_run().add_picture(str(data['map_picture']), width=Inches(3.35), height=Inches(1.70))
        p = right.add_paragraph()
        self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
        caption = data['map_location_label'] if data.get('gps_point') else data['country_display']
        self._run(p, caption, bold=True, size=7.0, color=BLUE_2)

    def _build_annex_pages(self, doc: Document, data: dict[str, Any], lang: str) -> None:
        images = data.get('annex_images') or []
        if not images:
            return
        for start in range(0, len(images), 6):
            doc.add_page_break()
            title = 'ANNEXE VISUELLE — Médias complémentaires' if lang == 'fr' else 'VISUAL ANNEX — Additional media'
            p = doc.add_paragraph()
            self._pconf(p, after=2)
            self._run(p, title, bold=True, size=16, color=BLUE)
            p = doc.add_paragraph()
            self._pconf(p, after=6)
            subtitle = 'Images complémentaires liées à l’activité et conservées dans le dossier technique de preuve.' if lang == 'fr' else 'Additional images linked to the activity and kept in the technical evidence folder.'
            self._run(p, subtitle, size=8.4, color=MUTED)
            chunk = images[start:start + 6]
            cols = min(3, max(1, len(chunk)))
            rows = math.ceil(len(chunk) / cols)
            grid = doc.add_table(rows=rows, cols=cols)
            self._table_no_borders(grid)
            grid.alignment = WD_TABLE_ALIGNMENT.CENTER
            grid.autofit = False
            card_width = 10.45 / cols
            card_height = 2.55 if rows > 1 else 4.25
            for row in grid.rows:
                self._set_row_height(row, Inches(card_height))
                for cell in row.cells:
                    self._set_cell_width(cell, Inches(card_width))
                    self._shade_cell(cell, SOFT_BLUE)
                    self._border_cell(cell, BORDER_LIGHT)
                    self._cell_margins(cell, top=55, bottom=35, start=55, end=55)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            for idx, path in enumerate(chunk):
                c = grid.cell(idx // cols, idx % cols)
                pic_px = (880, 560) if cols <= 2 else (720, 450)
                picture = self._prepare_photo(path, self._temp_dir / f'annex_{start + idx}.png', pic_px, fit='cover') if self._temp_dir else None
                p = self._cell_p(c)
                self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
                if picture and Path(picture).exists():
                    p.add_run().add_picture(str(picture), width=Inches(min(card_width - 0.45, 4.45)))
                p = c.add_paragraph()
                self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, line=1.0, after=0)
                label = f"Média complémentaire {start + idx + 1}" if lang == 'fr' else f'Additional media {start + idx + 1}'
                self._run(p, label, bold=True, size=7.8, color=BLUE_2)
                p = c.add_paragraph()
                self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, line=1.0, after=0)
                self._run(p, f"{data['activity']} | {data['location']} | {data['report_date']}", size=6.8, color=MUTED)

    # ------------------------------------------------------------------ images, icons, maps

    def _icon(self, kind: str) -> Path:
        return self._make_icon(kind, sector=False)

    def _sector_icon(self, kind: str) -> Path:
        return self._make_icon(kind, sector=True)

    def _make_icon(self, kind: str, sector: bool = False) -> Path:
        key = f"{'sector_' if sector else ''}{kind}"
        if key in self._icon_cache and self._icon_cache[key].exists():
            return self._icon_cache[key]
        if self._temp_dir is None:
            raise RuntimeError('temporary directory not initialized')
        path = self._temp_dir / f'{key}.png'
        size = 220
        im = Image.new('RGBA', (size, size), (255, 255, 255, 0))
        d = ImageDraw.Draw(im)
        if sector:
            self._draw_sector_icon_pil(d, kind, (size // 2, size // 2), (*_pil_rgb(BLUE_2), 255), scale=1.12)
        else:
            d.ellipse((20, 20, size - 20, size - 20), fill=(*_pil_rgb(PALE_BLUE), 255), outline=(*_pil_rgb(BORDER_LIGHT), 255), width=4)
            self._draw_line_icon_pil(d, kind, (size // 2, size // 2), (*_pil_rgb(BLUE_2), 255), scale=1.08)
        im.save(path)
        self._icon_cache[key] = path
        return path

    def _prepare_photo(self, source: Path | None, output: Path, size: tuple[int, int], fit: str = 'cover') -> Path | None:
        if source and Path(source).exists():
            try:
                im = Image.open(source).convert('RGB')
                im = self._contain_pad(im, size) if fit == 'contain' else self._cover_crop(im, size)
                im.save(output, quality=92)
                return output
            except Exception:
                pass
        im = Image.new('RGB', size, '#F1F5FA')
        d = ImageDraw.Draw(im)
        d.rounded_rectangle((0, 0, size[0] - 1, size[1] - 1), radius=18, fill='#F1F5FA', outline='#D9E4F2')
        font = self._font(22, bold=True)
        text = 'Photo principale indisponible'
        tw = self._text_w(d, text, font)
        d.text(((size[0] - tw) / 2, size[1] / 2 - 10), text, fill='#6B7280', font=font)
        im.save(output, quality=92)
        return output

    def _contain_pad(self, img: Image.Image, size: tuple[int, int]) -> Image.Image:
        contained = ImageOps.contain(img.convert('RGB'), size, Image.Resampling.LANCZOS)
        canvas = Image.new('RGB', size, '#FFFFFF')
        x = (size[0] - contained.width) // 2
        y = (size[1] - contained.height) // 2
        canvas.paste(contained, (x, y))
        d = ImageDraw.Draw(canvas)
        d.rounded_rectangle((0, 0, size[0] - 1, size[1] - 1), radius=18, outline='#D9E4F2', width=3)
        return canvas

    def _cover_crop(self, img: Image.Image, size: tuple[int, int]) -> Image.Image:
        tw, th = size
        ratio = tw / th
        img = img.convert('RGB')
        if img.width / max(img.height, 1) > ratio:
            new_w = int(img.height * ratio)
            x = (img.width - new_w) // 2
            img = img.crop((x, 0, x + new_w, img.height))
        else:
            new_h = int(img.width / ratio)
            y = max(0, int((img.height - new_h) * 0.35))
            img = img.crop((0, y, img.width, y + new_h))
        return img.resize((tw, th), Image.Resampling.LANCZOS)

    def _render_country_silhouette(self, data: dict[str, Any], output: Path, size: tuple[int, int], for_header: bool = False) -> Path:
        w, h = size
        im = Image.new('RGBA', size, (255, 255, 255, 0))
        d = ImageDraw.Draw(im)
        shape = self._country_shape(data['country'])
        if shape:
            polygons, bbox = self._shape_polygons(shape, data['country'])
            mapper = self._map_transform(bbox, w, h, 10)
            for poly in polygons:
                pts = [mapper(lon, lat) for lon, lat in poly]
                if len(pts) >= 3:
                    d.polygon(pts, fill=(255, 255, 255, 238), outline=(255, 255, 255, 255))
            point = data.get('gps_point')
            if point:
                lat, lon, _ = point
                if self._point_in_bbox(lon, lat, bbox):
                    px, py = mapper(lon, lat)
                    self._draw_pin_pil(d, px, py, 10)
        im.save(output)
        return output

    def _render_map_image(self, data: dict[str, Any], output: Path, size: tuple[int, int], lang: str) -> Path:
        width, height = size
        im = Image.new('RGB', size, '#FFFFFF')
        draw = ImageDraw.Draw(im)
        shape = self._country_shape(data['country'])
        if not shape:
            draw.rounded_rectangle((1, 1, width - 2, height - 2), radius=20, outline='#D4DAE3', fill='#FFFFFF', width=3)
            self._center_text(draw, data['country_display'], (0, 0, width, height), self._font(30, bold=True), '#6B7280')
            im.save(output, quality=95)
            return output
        polygons, bbox = self._shape_polygons(shape, data['country'])
        minx, miny, maxx, maxy = bbox
        dx = max(maxx - minx, 0.01)
        dy = max(maxy - miny, 0.01)
        bbox_expanded = (minx - dx * 0.26, miny - dy * 0.24, maxx + dx * 0.26, maxy + dy * 0.24)
        mapper = self._map_transform(bbox_expanded, width, height, 28)
        target_name = self._shape_country_name(data['country'])
        # background countries
        for sr in self._iter_shapes():
            shp = sr.shape
            sx1, sy1, sx2, sy2 = [float(v) for v in shp.bbox]
            if sx2 < bbox_expanded[0] or sx1 > bbox_expanded[2] or sy2 < bbox_expanded[1] or sy1 > bbox_expanded[3]:
                continue
            name = self._record_name(sr.record)
            is_target = name == target_name
            s_polys, _ = self._shape_polygons(shp, name)
            for poly in s_polys:
                pts = [mapper(lon, lat) for lon, lat in poly]
                if len(pts) >= 3:
                    draw.polygon(pts, fill='#EAF3FF' if is_target else '#F2F4F7', outline='#7EA6D9' if is_target else '#DCE3EC')
        country_font = self._font(34, bold=True)
        cx, cy = mapper((minx + maxx) / 2, (miny + maxy) / 2)
        label = data['country_display'].upper()
        draw.text((cx - self._text_w(draw, label, country_font) / 2, cy - 18), label, fill='#5F6B7A', font=country_font)
        label_font = self._font(16)
        labels = []
        for sr in self._iter_shapes():
            name = self._record_name(sr.record)
            if name == target_name:
                continue
            sx1, sy1, sx2, sy2 = [float(v) for v in sr.shape.bbox]
            if sx2 < bbox_expanded[0] or sx1 > bbox_expanded[2] or sy2 < bbox_expanded[1] or sy1 > bbox_expanded[3]:
                continue
            labels.append((name, (sx1 + sx2) / 2, (sy1 + sy2) / 2))
        for name, lon, lat in labels[:5]:
            lx, ly = mapper(lon, lat)
            txt = self._display_country(name, name, lang).upper()
            draw.text((lx - self._text_w(draw, txt, label_font) / 2, ly - 8), txt, fill='#A1AAB6', font=label_font)
        point = data.get('gps_point')
        if point:
            lat, lon, pin_label = point
            if self._point_in_bbox(lon, lat, bbox_expanded):
                px, py = mapper(lon, lat)
                self._draw_pin_pil(draw, px, py, 25)
                draw.text((px + 22, py - 16), self._truncate_text(pin_label or data['location'], 20), fill='#1D2D44', font=self._font(18, bold=True))
        draw.rounded_rectangle((1, 1, width - 2, height - 2), radius=20, outline='#D4DAE3', width=3)
        im.save(output, quality=95)
        return output

    # ------------------------------------------------------------------ extraction helpers
    # ------------------------------------------------------------------ extraction helpers

    def _project_payload(self, records: list[Any]) -> dict[str, Any]:
        for record in records:
            raw = getattr(record, 'raw', {}) or {}
            project = raw.get('project')
            if isinstance(project, dict) and project:
                return project
        return {}

    def _extract_org_name(self, project: dict[str, Any], records: list[Any]) -> str:
        keys = [
            'organizationName', 'organisationName', 'organization_name', 'organisation_name', 'orgName', 'org_name',
            'ngoName', 'ngo_name', 'ongName', 'ong_name', 'clientName', 'client_name', 'clientOrganizationName',
            'implementingPartner', 'implementing_partner', 'partnerName', 'partner_name', 'agencyName', 'agency_name',
            'workspaceName', 'workspace_name', 'organization', 'organisation', 'ngo', 'ong', 'accountName',
        ]
        for source in [project, *[(getattr(record, 'raw', {}) or {}) for record in records]]:
            for key in keys:
                value = source.get(key) if isinstance(source, dict) else None
                if isinstance(value, dict):
                    value = value.get('name') or value.get('displayName') or value.get('label')
                text = _safe_text(value)
                if text:
                    return text
        return self.default_org_name or 'Organisation'

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
            media_count = len([p for p in getattr(record, 'media_files', []) if Path(p).suffix.lower() in IMAGE_EXTENSIONS])
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

    def _polish_activity(self, activity: str, lang: str) -> str:
        text = _clean_spaces(activity)
        if lang != 'fr':
            return text[:1].upper() + text[1:] if text else 'Field activity'
        replacements = [
            (r'\bkits agricole\b', 'kits agricoles'),
            (r'\bkit agricole\b', 'kit agricole'),
            (r'\bgroupement agricole\b', 'groupement agricole'),
            (r'\bgroupements agricole\b', 'groupements agricoles'),
            (r'\bbénéficiaire\b', 'bénéficiaire'),
        ]
        for pattern, repl in replacements:
            text = re.sub(pattern, repl, text, flags=re.IGNORECASE)
        return text[:1].upper() + text[1:] if text else 'Activité terrain'

    def _polish_source_description(self, description: str) -> str:
        text = _clean_spaces(description)
        if not text:
            return ''
        text = re.sub(r'(\d)([A-Za-zÀ-ÿ])', r' ', text)
        text = re.sub(r'([A-Za-zÀ-ÿ])(\d)', r' ', text)
        text = re.sub(r'a\s+(\d)', r'à ', text, flags=re.IGNORECASE)
        fixes = [
            (r'kits agricole', 'kits agricoles'),
            (r'kit agricole', 'kit agricole'),
            (r'groupement et formations', 'groupements et formations'),
            (r'groupement agricole', 'groupement agricole'),
            (r'groupements agricole', 'groupements agricoles'),
            (r'agriculteur au sein', 'agriculteurs au sein'),
            (r'bénéficiaire issus', 'bénéficiaires issus'),
            (r'beneficiaire', 'bénéficiaire'),
            (r'beneficiaires', 'bénéficiaires'),
        ]
        for pattern, repl in fixes:
            text = re.sub(pattern, repl, text, flags=re.IGNORECASE)
        return text[:1].upper() + text[1:]

    def _report_title(self, activity: str, location: str, country: str, lang: str) -> str:
        short_location = location
        if ',' in short_location:
            parts = [part.strip() for part in short_location.split(',') if part.strip()]
            short_location = ', '.join(parts[:2])
        if lang == 'fr':
            return f"Rapport d’activité : {activity} à {short_location}"
        return f'Activity report: {activity} in {short_location}'

    def _detect_sectors(self, records: list[Any], activity: str, project_name: str, lang: str) -> list[dict[str, str]]:
        corpus = ' '.join([
            activity,
            project_name,
            *[_safe_text(getattr(record, 'title', '')) for record in records],
            *[_safe_text(getattr(record, 'description', '')) for record in records],
            *[_safe_text((getattr(record, 'raw', {}) or {}).get('activity')) for record in records],
            *[_safe_text((getattr(record, 'raw', {}) or {}).get('output')) for record in records],
        ]).lower()
        scored: list[tuple[int, int, dict[str, str]]] = []
        for order, rule in enumerate(SECTOR_RULES):
            score = sum(1 for keyword in rule['keywords'] if keyword.lower() in corpus)
            if score:
                scored.append((score, -order, {'key': rule['key'], 'label': rule[lang], 'icon': rule['icon']}))
        if not scored:
            return [{'key': 'coordination', 'label': 'Coordination', 'icon': 'coordination'}]
        scored.sort(key=lambda item: (-item[0], -item[1]))
        result: list[dict[str, str]] = []
        seen: set[str] = set()
        for _, __, sector in scored:
            if sector['key'] in seen:
                continue
            result.append(sector)
            seen.add(sector['key'])
            if len(result) == 2:
                break
        return result

    def _extract_indicators(self, records: list[Any], lang: str) -> list[dict[str, str]]:
        corpus = ' '.join([
            _safe_text(getattr(record, 'description', '')) + ' ' +
            _safe_text(getattr(record, 'title', '')) + ' ' +
            _safe_text((getattr(record, 'raw', {}) or {}).get('activity'))
            for record in records
        ])
        indicators: list[dict[str, str]] = []
        seen: set[tuple[str, str]] = set()
        for key, pattern, fr_label, en_label, icon, priority in INDICATOR_PATTERNS:
            for match in re.finditer(pattern, corpus, flags=re.IGNORECASE):
                number = self._normalize_number(match.group(1))
                if not number:
                    continue
                sig = (key, number)
                if sig in seen:
                    continue
                indicators.append({
                    'kind': key,
                    'value': number,
                    'label': fr_label if lang == 'fr' else en_label,
                    'icon': icon,
                    'priority': str(priority),
                })
                seen.add(sig)
        indicators.sort(key=lambda item: (-int(item.get('priority', '0')), item['label']))
        return indicators

    def _extract_primary_quantity(self, records: list[Any], sectors: list[dict[str, str]], lang: str) -> dict[str, str]:
        indicators = self._extract_indicators(records, lang)
        return indicators[0] if indicators else {'value': str(len([r for r in records if getattr(r, 'kind', '') == 'evidence']) or len(records)), 'label': 'Preuves' if lang == 'fr' else 'Evidence items', 'icon': 'check', 'kind': 'evidence'}

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

    def _select_hero_image(self, image_files: list[Path], primary_record: Any, records: list[Any]) -> Path | None:
        if not image_files:
            return None
        primary_images = [Path(p) for p in getattr(primary_record, 'media_files', []) if Path(p).suffix.lower() in IMAGE_EXTENSIONS]
        candidates = primary_images or image_files
        best = None
        best_score = -1.0
        for path in candidates:
            try:
                im = Image.open(path).convert('RGB')
                w, h = im.size
                small = im.resize((96, 96), Image.Resampling.LANCZOS).convert('L')
                stat = ImageStat.Stat(small)
                contrast = stat.stddev[0]
                mean = stat.mean[0]
                white_penalty = 200 if mean > 235 and contrast < 35 else 0
                aspect_bonus = 70 if w >= h else 0
                score = min(w * h / 50000, 120) + contrast + aspect_bonus - white_penalty
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

    def _target_group_label(self, quantity: dict[str, str], sectors: list[dict[str, str]], description: str, lang: str) -> str:
        label = quantity['label'].lower()
        corpus = f"{label} {description}".lower()
        if lang == 'fr':
            if 'agriculteur' in corpus or any(s['key'] == 'agriculture' for s in sectors):
                return 'Agriculteurs et groupements ciblés'
            if 'abri' in corpus or any(s['key'] == 'shelter' for s in sectors):
                return 'Ménages et communautés ciblées'
            if 'école' in corpus or any(s['key'] == 'education' for s in sectors):
                return 'Élèves, enseignants et communautés scolaires'
            if 'bénéficiaire' in corpus or 'participant' in corpus:
                return 'Bénéficiaires de l’activité'
            return 'Communautés ciblées'
        if 'farmer' in corpus or any(s['key'] == 'agriculture' for s in sectors):
            return 'Farmers and targeted groups'
        if 'shelter' in corpus or any(s['key'] == 'shelter' for s in sectors):
            return 'Beneficiary households and communities'
        return 'Targeted communities'

    def _evidence_type_label(self, primary_record: Any, lang: str) -> str:
        subtype = _safe_text(getattr(primary_record, 'subtype', '') or getattr(primary_record, 'raw', {}).get('type')).lower()
        if 'photo' in subtype or 'image' in subtype:
            return 'Preuve photo' if lang == 'fr' else 'Photo evidence'
        if 'video' in subtype or 'vidéo' in subtype:
            return 'Preuve vidéo' if lang == 'fr' else 'Video evidence'
        if 'document' in subtype or 'doc' in subtype:
            return 'Document source' if lang == 'fr' else 'Source document'
        return 'Preuve terrain' if lang == 'fr' else 'Field evidence'

    def _indicator_value(self, indicators: list[dict[str, str]], kind: str) -> str:
        for item in indicators:
            if item.get('kind') == kind:
                return item.get('value', '')
        return ''

    def _indicator_phrase(self, indicators: list[dict[str, str]], lang: str) -> str:
        focus = [item for item in indicators if item.get('kind') not in {'media', 'evidence'}][:3]
        if not focus:
            return ''
        return ', '.join([f"{item['value']} {item['label'].lower()}" for item in focus[:-1]]) + ((' et ' if lang == 'fr' else ' and ') + f"{focus[-1]['value']} {focus[-1]['label'].lower()}" if len(focus) > 1 else f"{focus[0]['value']} {focus[0]['label'].lower()}")

    def _narrative(self, activity: str, project_name: str, description: str, indicators: list[dict[str, str]], location: str, sectors: list[dict[str, str]], lang: str) -> str:
        act = activity[:1].lower() + activity[1:] if activity else ('activité terrain' if lang == 'fr' else 'field activity')
        sector_keys = {s.get('key') for s in sectors}
        beneficiaries = self._indicator_value(indicators, 'beneficiary')
        kits = self._indicator_value(indicators, 'kit')
        groups = self._indicator_value(indicators, 'group')
        shelters = self._indicator_value(indicators, 'shelter')
        phrase = self._indicator_phrase(indicators, lang)
        if lang == 'fr':
            if {'food_security', 'agriculture'} & sector_keys and (kits or beneficiaries):
                parts = [f"Cette activité de {act} documentée à {location} s’inscrit dans un appui direct aux moyens d’existence et à la sécurité alimentaire des ménages ciblés."]
                if phrase:
                    parts.append(f"Les informations disponibles font ressortir {phrase}, ce qui donne une lecture opérationnelle claire de l’effort réalisé.")
                if groups:
                    parts.append(f"Le travail avec {groups} groupements constitue un levier important : il favorise l’entraide entre membres, la diffusion des pratiques et la pérennité de l’investissement au-delà de la distribution initiale.")
                parts.append("La prochaine étape utile consistera à suivre l’utilisation effective des kits et l’évolution de la production, afin d’éclairer les décisions de réinvestissement et les possibilités de mise à l’échelle.")
                return ' '.join(_sentence(p) for p in parts)
            if 'shelter' in sector_keys or shelters:
                parts = [f"Cette activité de {act} documentée à {location} contribue à améliorer les conditions de vie et de protection des ménages ciblés."]
                if phrase:
                    parts.append(f"Les données disponibles mettent en avant {phrase}, donnant une indication concrète du volume d’appui fourni.")
                parts.append("Un suivi de l’occupation, de la qualité des abris et de la satisfaction des ménages permettra de confirmer la contribution de l’intervention à la sécurité et à la dignité des bénéficiaires.")
                return ' '.join(_sentence(p) for p in parts)
            parts = [f"Cette activité de {act} a été documentée à {location} dans le cadre du projet {project_name}."]
            if phrase:
                parts.append(f"Les informations disponibles mettent en avant {phrase} comme résultats immédiatement visibles de l’activité.")
            if description:
                parts.append(f"La description source précise : {description}.")
            parts.append("Le suivi des effets observés permettra de relier plus finement les réalisations documentées aux changements attendus pour les communautés ciblées.")
            return ' '.join(_sentence(p) for p in parts)
        # English fallback
        parts = [f"This {act} activity was documented in {location} under the project {project_name}."]
        if phrase:
            parts.append(f"Available information highlights {phrase} as the most visible immediate results of the activity.")
        parts.append('Follow-up on use and observed effects will help connect documented outputs with expected changes for targeted communities.')
        return ' '.join(_sentence(p) for p in parts)

    def _highlights(self, activity: str, project_code: str, description: str, indicators: list[dict[str, str]], media_count: int, location: str, sectors: list[dict[str, str]], lang: str) -> list[str]:
        act = activity[:1].lower() + activity[1:] if activity else ('activité terrain' if lang == 'fr' else 'field activity')
        phrase = self._indicator_phrase(indicators, lang)
        sector_text = ', '.join([s['label'] for s in sectors[:2]])
        if lang == 'fr':
            highlights = [f"Une activité de {act} a été documentée à {location} dans le cadre du projet {project_code}."]
            if phrase:
                highlights.append(f"Les données source font ressortir {phrase} comme principaux chiffres opérationnels.")
            if media_count > 0:
                highlights.append(f"{media_count} média(s) lié(s) apportent un appui visuel à la documentation de l’activité.")
            if sector_text:
                highlights.append(f"L’activité est alignée principalement avec les secteurs {sector_text.lower()}.")
            return [_sentence(item) for item in highlights]
        highlights = [f"A {act} activity was documented in {location} under project {project_code}."]
        if phrase:
            highlights.append(f"Source data highlights {phrase} as the main operational figures.")
        if media_count > 0:
            highlights.append(f"{media_count} linked media asset(s) provide visual support to document the activity.")
        if sector_text:
            highlights.append(f"The activity is primarily aligned with the {sector_text} sectors.")
        return [_sentence(item) for item in highlights]

    # ------------------------------------------------------------------ map helpers
    # ------------------------------------------------------------------ map helpers

    def _shape_file(self) -> Path:
        return resource_path('desktop_sync_app/assets/maps/naturalearth_lowres/naturalearth_lowres.shp')

    def _iter_shapes(self):
        if self._shape_cache is not None:
            return self._shape_cache
        if not self._shape_file().exists():
            self._shape_cache = []
            return self._shape_cache
        reader = shapefile.Reader(str(self._shape_file()), encoding='latin1')
        self._shape_cache = list(reader.iterShapeRecords())
        return self._shape_cache

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

    def _shape_polygons(self, shape, country: str) -> tuple[list[list[tuple[float, float]]], tuple[float, float, float, float]]:
        points = shape.points
        parts = list(shape.parts) + [len(points)]
        polygons: list[list[tuple[float, float]]] = []
        for index in range(len(parts) - 1):
            pts = points[parts[index]:parts[index + 1]]
            if len(pts) >= 3:
                poly = [(float(x), float(y)) for x, y in pts]
                if self._shape_country_name(country) == 'France':
                    xs = [p[0] for p in poly]
                    ys = [p[1] for p in poly]
                    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
                    if not (-6.5 <= cx <= 10.5 and 41.0 <= cy <= 52.5):
                        continue
                polygons.append(poly)
        if not polygons:
            return [], tuple(float(v) for v in shape.bbox)  # type: ignore[return-value]
        xs = [x for poly in polygons for x, _ in poly]
        ys = [y for poly in polygons for _, y in poly]
        return polygons, (min(xs), min(ys), max(xs), max(ys))

    def _map_transform(self, bbox: tuple[float, float, float, float], width: int, height: int, pad: int):
        minx, miny, maxx, maxy = bbox
        scale = min((width - pad * 2) / max(maxx - minx, 0.001), (height - pad * 2) / max(maxy - miny, 0.001))
        offset_x = (width - (maxx - minx) * scale) / 2
        offset_y = (height - (maxy - miny) * scale) / 2

        def mapper(lon: float, lat: float) -> tuple[float, float]:
            x = offset_x + (lon - minx) * scale
            y = offset_y + (maxy - lat) * scale
            return x, y
        return mapper

    def _point_in_bbox(self, lon: float, lat: float, bbox: tuple[float, float, float, float]) -> bool:
        return bbox[0] <= lon <= bbox[2] and bbox[1] <= lat <= bbox[3]

    # ------------------------------------------------------------------ docx formatting helpers

    def _spacer(self, doc: Document, points: float) -> None:
        p = doc.add_paragraph()
        self._pconf(p, before=0, after=0, line=0.1)
        p.add_run().font.size = Pt(points)

    def _section_title(self, cell, title: str, icon: str | None = None, line_chars: str = '──────────') -> None:
        table = cell.add_table(rows=1, cols=2)
        self._table_no_borders(table)
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.autofit = False
        self._set_cell_width(table.cell(0, 0), Inches(0.30))
        self._set_cell_width(table.cell(0, 1), Inches(2.72))
        self._cell_margins(table.cell(0, 0), top=0, bottom=0, start=0, end=10)
        self._cell_margins(table.cell(0, 1), top=0, bottom=0, start=0, end=0)
        if icon:
            p = self._cell_p(table.cell(0, 0))
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
            p.add_run().add_picture(str(self._icon(icon)), width=Inches(0.21))
        p = self._cell_p(table.cell(0, 1))
        self._pconf(p, after=0, line=0.95)
        self._run(p, title, bold=True, size=8.6, color=BLUE_2)
        p = table.cell(0, 1).add_paragraph()
        self._pconf(p, after=1, line=0.45)
        self._run(p, line_chars, size=2.2, color='#9CB7DE')

    def _cell_p(self, cell):
        return cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()

    def _run(self, paragraph, text: str, bold: bool = False, size: float = 8.5, color: str = DARK, italic: bool = False):
        run = paragraph.add_run(str(text))
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size)
        run.font.color.rgb = _rgb(color)
        run.font.name = 'Arial'
        try:
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
        except Exception:
            pass
        return run

    def _pconf(self, paragraph, before: float = 0, after: float = 0, line: float = 1.0, align=None) -> None:
        paragraph.paragraph_format.space_before = Pt(before)
        paragraph.paragraph_format.space_after = Pt(after)
        paragraph.paragraph_format.line_spacing = line
        if align is not None:
            paragraph.alignment = align

    def _shade_cell(self, cell, fill: str) -> None:
        tcPr = cell._tc.get_or_add_tcPr()
        shd = tcPr.find(qn('w:shd'))
        if shd is None:
            shd = OxmlElement('w:shd')
            tcPr.append(shd)
        shd.set(qn('w:fill'), fill)

    def _border_cell(self, cell, color: str = BORDER, size: int = 4) -> None:
        tcPr = cell._tc.get_or_add_tcPr()
        borders = tcPr.first_child_found_in('w:tcBorders')
        if borders is None:
            borders = OxmlElement('w:tcBorders')
            tcPr.append(borders)
        for edge in ('top', 'left', 'bottom', 'right'):
            el = borders.find(qn(f'w:{edge}'))
            if el is None:
                el = OxmlElement(f'w:{edge}')
                borders.append(el)
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), str(size))
            el.set(qn('w:color'), color)

    def _table_no_borders(self, table) -> None:
        for row in table.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                borders = tcPr.first_child_found_in('w:tcBorders')
                if borders is None:
                    borders = OxmlElement('w:tcBorders')
                    tcPr.append(borders)
                for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                    el = borders.find(qn(f'w:{edge}'))
                    if el is None:
                        el = OxmlElement(f'w:{edge}')
                        borders.append(el)
                    el.set(qn('w:val'), 'nil')

    def _set_cell_width(self, cell, width) -> None:
        twips = int(width.inches * 1440) if hasattr(width, 'inches') else int(width)
        tcPr = cell._tc.get_or_add_tcPr()
        tcW = tcPr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW')
            tcPr.append(tcW)
        tcW.set(qn('w:w'), str(twips))
        tcW.set(qn('w:type'), 'dxa')

    def _set_row_height(self, row, height) -> None:
        row.height = height
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    def _cell_margins(self, cell, top=70, bottom=70, start=70, end=70) -> None:
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = tcPr.first_child_found_in('w:tcMar')
        if tcMar is None:
            tcMar = OxmlElement('w:tcMar')
            tcPr.append(tcMar)
        for name, val in [('top', top), ('bottom', bottom), ('start', start), ('end', end)]:
            node = tcMar.find(qn(f'w:{name}'))
            if node is None:
                node = OxmlElement(f'w:{name}')
                tcMar.append(node)
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')

    # ------------------------------------------------------------------ PIL drawing helpers

    def _draw_line_icon_pil(self, draw: ImageDraw.ImageDraw, icon: str, center: tuple[int, int], color: tuple[int, int, int, int], scale: float = 1.0) -> None:
        x, y = center
        s = 48 * scale
        w = max(4, int(5 * scale))
        c = color
        if icon in {'doc', 'document'}:
            draw.rectangle((x - s * .28, y - s * .45, x + s * .28, y + s * .45), outline=c, width=w)
            for yy in [-.20, 0, .20]:
                draw.line((x - s * .15, y + s * yy, x + s * .15, y + s * yy), fill=c, width=w)
        elif icon in {'bar', 'barchart'}:
            draw.rectangle((x - s*.48, y, x - s*.27, y + s*.46), fill=c)
            draw.rectangle((x - s*.10, y - s*.25, x + s*.11, y + s*.46), fill=c)
            draw.rectangle((x + s*.28, y - s*.55, x + s*.49, y + s*.46), fill=c)
        elif icon in {'people', 'person', 'org'}:
            draw.ellipse((x - s*.16, y - s*.50, x + s*.16, y - s*.18), outline=c, width=w)
            draw.arc((x - s*.52, y - s*.15, x + s*.52, y + s*.58), 190, 350, fill=c, width=w)
            draw.ellipse((x - s*.68, y - s*.32, x - s*.42, y - s*.06), outline=c, width=max(2, w-1))
            draw.ellipse((x + s*.42, y - s*.32, x + s*.68, y - s*.06), outline=c, width=max(2, w-1))
        elif icon in {'media', 'camera'}:
            draw.rounded_rectangle((x - s*.46, y - s*.32, x + s*.46, y + s*.32), radius=8, outline=c, width=w)
            if icon == 'camera':
                draw.ellipse((x - s*.18, y - s*.17, x + s*.18, y + s*.19), outline=c, width=w)
            else:
                draw.polygon([(x - s*.08, y - s*.22), (x - s*.08, y + s*.22), (x + s*.28, y)], fill=c)
        elif icon in {'calendar'}:
            draw.rounded_rectangle((x - s*.42, y - s*.48, x + s*.42, y + s*.42), radius=8, outline=c, width=w)
            draw.rectangle((x - s*.42, y - s*.48, x + s*.42, y - s*.20), fill=c)
            draw.line((x - s*.20, y - s*.62, x - s*.20, y - s*.34), fill=c, width=w)
            draw.line((x + s*.20, y - s*.62, x + s*.20, y - s*.34), fill=c, width=w)
        elif icon in {'check'}:
            draw.rounded_rectangle((x - s*.38, y - s*.52, x + s*.38, y + s*.50), radius=8, outline=c, width=w)
            draw.line((x - s*.19, y, x - s*.02, y + s*.18, x + s*.28, y - s*.22), fill=c, width=w)
            draw.rectangle((x - s*.18, y - s*.62, x + s*.18, y - s*.42), outline=c, width=max(2, w - 1))
        elif icon in {'pin', 'pin_round'}:
            draw.ellipse((x - s*.28, y - s*.50, x + s*.28, y + s*.06), outline=c, width=w)
            draw.polygon([(x, y + s*.63), (x - s*.22, y), (x + s*.22, y)], outline=c)
            draw.ellipse((x - s*.09, y - s*.32, x + s*.09, y - s*.14), fill=c)
        elif icon == 'star':
            pts = []
            for i in range(10):
                a = -math.pi / 2 + i * math.pi / 5
                rr = s * .50 if i % 2 == 0 else s * .22
                pts.append((x + rr * math.cos(a), y + rr * math.sin(a)))
            draw.line(pts + [pts[0]], fill=c, width=w)
        elif icon == 'sector':
            draw.ellipse((x - s*.45, y - s*.45, x + s*.45, y + s*.45), outline=c, width=w)
            draw.line((x, y - s*.58, x, y + s*.58), fill=c, width=w)
            draw.line((x - s*.58, y, x + s*.58, y), fill=c, width=w)
        else:
            draw.ellipse((x - 8, y - 8, x + 8, y + 8), fill=c)

    def _draw_sector_icon_pil(self, draw: ImageDraw.ImageDraw, icon: str, center: tuple[int, int], color: tuple[int, int, int, int], scale: float = 1.0) -> None:
        x, y = center
        s = 54 * scale
        w = max(4, int(5 * scale))
        c = color
        if icon == 'shelter':
            draw.line((x - s*.70, y, x, y - s*.58, x + s*.70, y), fill=c, width=w)
            draw.rectangle((x - s*.48, y, x + s*.48, y + s*.52), outline=c, width=w)
            draw.rectangle((x - s*.14, y + s*.22, x + s*.14, y + s*.52), outline=c, width=max(2, w-1))
        elif icon in {'food', 'nutrition'}:
            draw.arc((x - s*.72, y - s*.12, x + s*.72, y + s*.72), 0, 180, fill=c, width=w)
            draw.line((x - s*.67, y + s*.32, x + s*.67, y + s*.32), fill=c, width=w)
            draw.line((x - s*.30, y - s*.55, x - s*.30, y + s*.05), fill=c, width=max(2, w-1))
            draw.ellipse((x - s*.42, y - s*.75, x - s*.18, y - s*.52), outline=c, width=max(2, w-1))
        elif icon == 'agriculture':
            draw.line((x, y + s*.70, x, y - s*.45), fill=c, width=w)
            draw.arc((x - s*.58, y - s*.55, x + s*.02, y + s*.05), 210, 50, fill=c, width=w)
            draw.arc((x - s*.02, y - s*.75, x + s*.62, y - s*.05), 130, 320, fill=c, width=w)
            draw.arc((x - s*.64, y - s*.05, x + s*.00, y + s*.62), 210, 50, fill=c, width=w)
            draw.arc((x, y - s*.05, x + s*.70, y + s*.62), 130, 320, fill=c, width=w)
        elif icon == 'wash':
            draw.line((x, y - s*.70, x - s*.42, y + s*.05, x, y + s*.65, x + s*.42, y + s*.05, x, y - s*.70), fill=c, width=w)
        elif icon == 'health':
            draw.rectangle((x - s*.16, y - s*.62, x + s*.16, y + s*.62), fill=c)
            draw.rectangle((x - s*.62, y - s*.16, x + s*.62, y + s*.16), fill=c)
        elif icon == 'education':
            draw.line((x - s*.70, y - s*.10, x, y - s*.50, x + s*.70, y - s*.10, x, y + s*.28, x - s*.70, y - s*.10), fill=c, width=w)
            draw.rectangle((x - s*.40, y + s*.12, x + s*.40, y + s*.50), outline=c, width=w)
        elif icon == 'protection':
            pts = [(x, y - s*.72), (x + s*.56, y - s*.35), (x + s*.42, y + s*.52), (x, y + s*.76), (x - s*.42, y + s*.52), (x - s*.56, y - s*.35), (x, y - s*.72)]
            draw.line(pts, fill=c, width=w)
        elif icon == 'cash':
            draw.rounded_rectangle((x - s*.72, y - s*.40, x + s*.72, y + s*.40), radius=10, outline=c, width=w)
            draw.text((x - s*.10, y - s*.34), '$', fill=c, font=self._font(int(38 * scale), bold=True))
        else:
            self._draw_line_icon_pil(draw, 'sector', center, color, scale)

    def _draw_pin_pil(self, draw: ImageDraw.ImageDraw, x: float, y: float, size: int) -> None:
        orange = _pil_rgb(ORANGE)
        draw.ellipse((x - size * 0.55, y - size * 1.05, x + size * 0.55, y + size * 0.05), fill=orange)
        draw.polygon([(x, y + size * 0.85), (x - size * 0.42, y - size * 0.05), (x + size * 0.42, y - size * 0.05)], fill=orange)
        draw.ellipse((x - size * 0.2, y - size * 0.70, x + size * 0.2, y - size * 0.30), fill=_pil_rgb(WHITE))

    def _font(self, size: int, bold: bool = False):
        candidates = [
            r'C:\Windows\Fonts\arialbd.ttf' if bold else r'C:\Windows\Fonts\arial.ttf',
            '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf' if bold else '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        ]
        for path in candidates:
            try:
                if Path(path).exists():
                    return ImageFont.truetype(path, size=size)
            except Exception:
                continue
        return ImageFont.load_default()

    def _text_w(self, draw: ImageDraw.ImageDraw, text: str, font) -> int:
        try:
            return int(draw.textbbox((0, 0), text, font=font)[2])
        except Exception:
            return int(draw.textlength(text, font=font))

    def _center_text(self, draw: ImageDraw.ImageDraw, text: str, box: tuple[int, int, int, int], font, fill: str) -> None:
        x1, y1, x2, y2 = box
        tw = self._text_w(draw, text, font)
        th = int(draw.textbbox((0, 0), text, font=font)[3])
        draw.text((x1 + (x2 - x1 - tw) / 2, y1 + (y2 - y1 - th) / 2), text, fill=fill, font=font)

    def _truncate_text(self, text: str, max_chars: int) -> str:
        text = _clean_spaces(text)
        return text if len(text) <= max_chars else text[:max_chars - 1].rstrip() + '…'
