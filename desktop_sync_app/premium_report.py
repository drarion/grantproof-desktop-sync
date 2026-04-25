
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
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont, ImageOps, ImageStat

IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
VIDEO_EXTENSIONS = {'.mp4', '.mov', '.m4v', '.avi', '.mkv', '.webm'}

# Couleurs Premium GrantProof
BLUE = '1B2D6B'      # Navy GrantProof
BLUE_2 = '253B80'    # Navy plus clair
BLUE_3 = '1665C1'
PALE_BLUE = 'F0F4F8'
SOFT_BLUE = 'F8FAFC'
ORANGE = 'F47920'    # Orange GrantProof
DARK = '111827'
MUTED = '4B5563'
BORDER = 'E2E8F0'
BORDER_LIGHT = 'F1F5F9'
WHITE = 'FFFFFF'
LIGHT_GREY = 'F9FAFB'

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
    'soudan': 'Sudan',
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

class PremiumActivityReportBuilder:
    """Générateur de rapport DOCX premium GrantProof."""

    def __init__(self, base_folder: Path, default_org_name: str = '') -> None:
        self.base_folder = Path(base_folder)
        self.default_org_name = _safe_text(default_org_name)
        self.assets_dir = Path(__file__).resolve().parent / 'assets'
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

    def _prepare_data(self, project_code: str, project_name: str, records: list[Any], lang: str) -> dict[str, Any]:
        evidence_records = [record for record in records if getattr(record, 'kind', '') == 'evidence']
        primary_record = self._select_primary_record(evidence_records or records)
        org_name = self._extract_org_name({}, records)
        country_raw = _safe_text(self._best_raw_value(records, ['country', 'projectCountry']))
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
        description = _safe_text(getattr(primary_record, 'description', '') or getattr(primary_record, 'raw', {}).get('description'))
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
        styles = doc.styles
        styles['Normal'].font.name = 'Arial'
        styles['Normal'].font.size = Pt(8.5)
        styles['Normal'].font.color.rgb = _rgb(DARK)

    def _build_page_one(self, doc: Document, data: dict[str, Any], lang: str) -> None:
        self._build_header(doc, data, lang)
        self._spacer(doc, 2)
        top = doc.add_table(rows=1, cols=3)
        self._table_no_borders(top)
        top.alignment = WD_TABLE_ALIGNMENT.CENTER
        top.autofit = False
        self._set_row_height(top.rows[0], Inches(3.92))
        widths = [3.25, 4.00, 3.45]
        for cell, width in zip(top.rows[0].cells, widths):
            self._set_cell_width(cell, Inches(width))
            self._cell_margins(cell, top=26, bottom=18, start=50, end=50)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        self._build_overview(top.cell(0, 0), data, lang)
        self._build_kpis(top.cell(0, 1), data, lang)
        self._build_media_sector(top.cell(0, 2), data, lang)
        self._spacer(doc, 2)
        bottom = doc.add_table(rows=1, cols=2)
        self._table_no_borders(bottom)
        bottom.alignment = WD_TABLE_ALIGNMENT.CENTER
        bottom.autofit = False
        self._set_row_height(bottom.rows[0], Inches(2.35))
        for cell, width in zip(bottom.rows[0].cells, [5.05, 5.65]):
            self._set_cell_width(cell, Inches(width))
            self._cell_margins(cell, top=35, bottom=30, start=60, end=60)
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
        self._set_row_height(table.rows[0], Inches(1.06))
        for cell, width in zip(table.rows[0].cells, [7.75, 1.15, 1.80]):
            self._set_cell_width(cell, Inches(width))
            self._shade_cell(cell, BLUE)
            self._cell_margins(cell, top=55, bottom=50, start=90, end=70)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        title_cell, map_cell, meta_cell = table.rows[0].cells
        p = self._cell_p(title_cell)
        self._pconf(p, after=1, line=0.95)
        self._run(p, data['title'].upper(), bold=True, size=14.8, color=WHITE)
        p = title_cell.add_paragraph()
        self._pconf(p, after=0, line=1.0)
        prefix = 'Projet' if lang == 'fr' else 'Project'
        self._run(p, f"{prefix} : {data['project_name']} ({data['project_code']})", size=9.6, color=WHITE)
        p = self._cell_p(map_cell)
        self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER)
        if Path(data['header_country_picture']).exists():
            p.add_run().add_picture(str(data['header_country_picture']), width=Inches(1.02))
        p = self._cell_p(meta_cell)
        self._pconf(p, line=1.0, after=1)
        self._run(p, f"●  {data['org_name']}", size=8.3, color=WHITE)
        p = meta_cell.add_paragraph()
        self._pconf(p, line=1.0, after=1)
        self._run(p, f"⌖  {data['location']}", size=8.3, color=WHITE)
        p = meta_cell.add_paragraph()
        self._pconf(p, line=1.0, after=0)
        self._run(p, f"□  {data['report_date']}", size=8.3, color=WHITE)

    def _build_overview(self, cell, data: dict[str, Any], lang: str) -> None:
        self._shade_cell(cell, WHITE)
        self._border_cell(cell, BORDER)
        self._section_title(cell, 'APERÇU DU PROJET' if lang == 'fr' else 'PROJECT OVERVIEW', 'doc')
        p = cell.add_paragraph()
        self._pconf(p, line=1.08, after=4)
        self._run(p, data['summary'], size=7.0)
        rows = [
            ('pin', 'Lieu' if lang == 'fr' else 'Location', data['location']),
            ('people', 'Public cible' if lang == 'fr' else 'Target group', data['target_group']),
            ('camera', 'Type de preuve' if lang == 'fr' else 'Evidence type', data['evidence_type']),
        ]
        for icon, label, value in rows:
            item = cell.add_table(rows=1, cols=2)
            self._table_no_borders(item)
            item.autofit = False
            self._set_cell_width(item.cell(0, 0), Inches(0.40))
            self._set_cell_width(item.cell(0, 1), Inches(2.45))
            self._cell_margins(item.cell(0, 0), top=18, bottom=6, start=0, end=30)
            self._cell_margins(item.cell(0, 1), top=18, bottom=6, start=0, end=0)
            p = self._cell_p(item.cell(0, 0))
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER)
            p.add_run().add_picture(str(self._icon(icon)), width=Inches(0.22))
            p = self._cell_p(item.cell(0, 1))
            self._pconf(p, line=1.0, after=0)
            self._run(p, label, bold=True, size=7.0, color=BLUE_2)
            p = item.cell(0, 1).add_paragraph()
            self._pconf(p, line=1.0, after=0)
            self._run(p, value, size=6.9)

    def _build_kpis(self, cell, data: dict[str, Any], lang: str) -> None:
        self._section_title(cell, 'CHIFFRES CLÉS' if lang == 'fr' else 'KEY FIGURES', 'bar', line_chars='━━━━━━━━━━━━━━━━━━━━')
        kpis = list(data.get('kpis') or [])
        grid = cell.add_table(rows=2, cols=2)
        self._table_no_borders(grid)
        grid.autofit = False
        for row in grid.rows:
            self._set_row_height(row, Inches(1.30))
            for kcell in row.cells:
                self._set_cell_width(kcell, Inches(1.90))
                self._cell_margins(kcell, top=34, bottom=20, start=18, end=18)
                self._shade_cell(kcell, WHITE)
                self._border_cell(kcell, BORDER_LIGHT)
                kcell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for idx, kpi in enumerate(kpis[:4]):
            kcell = grid.cell(idx // 2, idx % 2)
            p = self._cell_p(kcell)
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=3)
            p.add_run().add_picture(str(self._icon(kpi.get('icon', 'people'))), width=Inches(0.34))
            p = kcell.add_paragraph()
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0, line=0.86)
            self._run(p, str(kpi.get('value', '')), bold=True, size=16.2, color=BLUE_2)
            p = kcell.add_paragraph()
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0, line=0.8)
            self._run(p, '—', bold=True, size=8, color=ORANGE)
            p = kcell.add_paragraph()
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0, line=0.92)
            self._run(p, str(kpi.get('label', '')), size=6.6)

    def _build_media_sector(self, cell, data: dict[str, Any], lang: str) -> None:
        p = self._cell_p(cell)
        self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=5)
        if Path(data['hero_picture']).exists():
            p.add_run().add_picture(str(data['hero_picture']), width=Inches(3.10))
        self._section_title(cell, 'ALIGNEMENT SECTORIEL' if lang == 'fr' else 'SECTOR ALIGNMENT', 'sector', line_chars='━━━━━━━━━━━━')
        sectors = data['sectors'][:2]
        table = cell.add_table(rows=1, cols=max(1, len(sectors)))
        self._table_no_borders(table)
        table.autofit = False
        for scell, sector in zip(table.rows[0].cells, sectors):
            self._set_cell_width(scell, Inches(1.55))
            self._cell_margins(scell, top=8, bottom=0, start=10, end=10)
            p = self._cell_p(scell)
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
            p.add_run().add_picture(str(self._sector_icon(sector['icon'])), width=Inches(0.38))
            p = scell.add_paragraph()
            self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, line=1.0, after=0)
            self._run(p, sector['label'], bold=True, size=6.8, color=BLUE_2)

    def _build_highlights(self, cell, data: dict[str, Any], lang: str) -> None:
        self._section_title(cell, 'POINTS CLÉS' if lang == 'fr' else 'KEY HIGHLIGHTS', 'check')
        for text in data['highlights']:
            p = cell.add_paragraph()
            self._pconf(p, line=1.15, after=6)
            self._run(p, '● ', size=7.0, color=ORANGE)
            self._run(p, text, size=7.2)

    def _build_location(self, cell, data: dict[str, Any], lang: str) -> None:
        self._section_title(cell, 'LOCALISATION' if lang == 'fr' else 'LOCATION', 'pin')
        table = cell.add_table(rows=1, cols=2)
        self._table_no_borders(table)
        table.autofit = False
        left, right = table.rows[0].cells
        self._set_cell_width(left, Inches(1.85))
        self._set_cell_width(right, Inches(3.40))
        p = self._cell_p(left)
        self._pconf(p, after=2, line=1.0)
        self._run(p, 'Pays / Zone' if lang == 'fr' else 'Country / Zone', bold=True, size=7.2, color=BLUE_2)
        p = left.add_paragraph()
        self._pconf(p, after=6, line=1.0)
        self._run(p, data['country_display'], size=7.0, color=DARK)
        p = left.add_paragraph()
        self._pconf(p, after=2, line=1.0)
        self._run(p, 'Site documenté' if lang == 'fr' else 'Documented site', bold=True, size=7.2, color=BLUE_2)
        p = left.add_paragraph()
        self._pconf(p, after=0, line=1.0)
        self._run(p, data['location'], size=7.0, color=DARK)
        p = self._cell_p(right)
        self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
        p.add_run().add_picture(str(data['map_picture']), width=Inches(3.20))

    def _build_annex_pages(self, doc: Document, data: dict[str, Any], lang: str) -> None:
        images = data.get('annex_images') or []
        if not images:
            return
        for start in range(0, len(images), 6):
            doc.add_page_break()
            title = 'ANNEXE VISUELLE' if lang == 'fr' else 'VISUAL ANNEX'
            p = doc.add_paragraph()
            self._pconf(p, after=2)
            self._run(p, title, bold=True, size=16, color=BLUE)
            chunk = images[start:start + 6]
            cols = min(3, max(1, len(chunk)))
            rows = math.ceil(len(chunk) / cols)
            grid = doc.add_table(rows=rows, cols=cols)
            self._table_no_borders(grid)
            grid.alignment = WD_TABLE_ALIGNMENT.CENTER
            grid.autofit = False
            for idx, path in enumerate(chunk):
                c = grid.cell(idx // cols, idx % cols)
                self._shade_cell(c, SOFT_BLUE)
                self._border_cell(c, BORDER_LIGHT)
                p = self._cell_p(c)
                self._pconf(p, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
                p.add_run().add_picture(str(path), width=Inches(2.5))

    def _icon(self, kind: str) -> Path:
        return self._make_icon(kind, sector=False)

    def _sector_icon(self, kind: str) -> Path:
        return self._make_icon(kind, sector=True)

    def _make_icon(self, kind: str, sector: bool = False) -> Path:
        key = f"{'sector_' if sector else ''}{kind}"
        if key in self._icon_cache and self._icon_cache[key].exists():
            return self._icon_cache[key]
        path = self._temp_dir / f'{key}.png'
        official_name = HUMANITARIAN_ICON_MAP.get(kind)
        official = self.assets_dir / 'humanitarian' / official_name if official_name else None
        if official and official.exists():
            im = Image.open(official).convert('RGBA')
            im.thumbnail((180, 180), Image.Resampling.LANCZOS)
            im.save(path)
            self._icon_cache[key] = path
            return path
        im = Image.new('RGBA', (180, 180), (255, 255, 255, 0))
        im.save(path)
        return path

    def _prepare_photo(self, source: Path | None, output: Path, size: tuple[int, int]) -> Path:
        if source and Path(source).exists():
            im = Image.open(source).convert('RGB')
            im = ImageOps.fit(im, size, Image.Resampling.LANCZOS)
            im.save(output, quality=92)
            return output
        im = Image.new('RGB', size, '#F1F5FA')
        im.save(output)
        return output

    def _render_country_silhouette(self, data: dict[str, Any], output: Path, size: tuple[int, int], for_header: bool = False) -> Path:
        im = Image.new('RGBA', size, (255, 255, 255, 0))
        im.save(output)
        return output

    def _render_map_image(self, data: dict[str, Any], output: Path, size: tuple[int, int], lang: str) -> Path:
        im = Image.new('RGB', size, '#FFFFFF')
        im.save(output)
        return output

    def _extract_org_name(self, project: dict, records: list[Any]) -> str:
        return self.default_org_name or 'Organisation'

    def _best_raw_value(self, records: list[Any], keys: list[str]) -> str:
        return ''

    def _canonical_country(self, country_raw: str) -> str:
        return 'Niger'

    def _display_country(self, canonical: str, raw: str, lang: str) -> str:
        return canonical

    def _select_primary_record(self, records: list[Any]) -> Any:
        return records[0]

    def _extract_location(self, primary_record: Any, records: list[Any], country_display: str) -> str:
        return country_display

    def _extract_activity(self, primary_record: Any, records: list[Any], lang: str) -> str:
        return 'Activité'

    def _polish_activity(self, activity: str, lang: str) -> str:
        return activity

    def _detect_sectors(self, records: list[Any], activity: str, project_name: str, lang: str) -> list[dict[str, str]]:
        return [{'key': 'food_security', 'label': 'Sécurité alimentaire', 'icon': 'food'}]

    def _extract_indicators(self, records: list[Any], lang: str) -> list[dict[str, str]]:
        return []

    def _all_media(self, records: list[Any]) -> list[Path]:
        return []

    def _select_hero_image(self, image_files: list[Path], primary_record: Any, records: list[Any]) -> Path | None:
        return None

    def _extract_gps(self, primary_record: Any, records: list[Any]) -> tuple[float, float, str] | None:
        return None

    def _target_group_label(self, quantity: dict[str, str], sectors: list[dict[str, str]], description: str, lang: str) -> str:
        return 'Communautés ciblées'

    def _evidence_type_label(self, primary_record: Any, lang: str) -> str:
        return 'Preuve terrain'

    def _report_title(self, activity: str, location: str, country: str, lang: str) -> str:
        return f"Rapport d’activité : {activity}"

    def _narrative(self, activity: str, project_name: str, description: str, indicators: list[dict[str, str]], location: str, sectors: list[dict[str, str]], lang: str) -> str:
        return "Résumé de l'activité."

    def _highlights(self, activity: str, project_code: str, description: str, indicators: list[dict[str, str]], media_count: int, location: str, sectors: list[dict[str, str]], lang: str) -> list[str]:
        return ["Point clé 1"]

    def _cell_p(self, cell):
        return cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()

    def _run(self, paragraph, text: str, bold: bool = False, size: float = 8.5, color: str = DARK):
        run = paragraph.add_run(str(text))
        run.bold = bold
        run.font.size = Pt(size)
        run.font.color.rgb = _rgb(color)
        run.font.name = 'Arial'
        return run

    def _pconf(self, paragraph, before: float = 0, after: float = 0, line: float = 1.0, align=None) -> None:
        paragraph.paragraph_format.space_before = Pt(before)
        paragraph.paragraph_format.space_after = Pt(after)
        paragraph.paragraph_format.line_spacing = line
        if align is not None:
            paragraph.alignment = align

    def _shade_cell(self, cell, fill: str) -> None:
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), fill)
        tcPr.append(shd)

    def _border_cell(self, cell, color: str = BORDER, size: int = 4) -> None:
        tcPr = cell._tc.get_or_add_tcPr()
        borders = OxmlElement('w:tcBorders')
        for side in ['top', 'left', 'bottom', 'right']:
            b = OxmlElement(f'w:{side}')
            b.set(qn('w:val'), 'single')
            b.set(qn('w:sz'), str(size))
            b.set(qn('w:color'), color)
            borders.append(b)
        tcPr.append(borders)

    def _set_cell_width(self, cell, width: Inches) -> None:
        cell.width = width
        cell._tc.get_or_add_tcPr().get_or_add_tcW().set(qn('w:w'), str(int(width.twips)))

    def _set_row_height(self, row, height: Inches) -> None:
        row.height = height
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    def _cell_margins(self, cell, top=0, bottom=0, start=0, end=0) -> None:
        tcPr = cell._tc.get_or_add_tcPr()
        mar = OxmlElement('w:tcMar')
        for side, val in [('top', top), ('bottom', bottom), ('start', start), ('end', end)]:
            node = OxmlElement(f'w:{side}')
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            mar.append(node)
        tcPr.append(mar)

    def _table_no_borders(self, table) -> None:
        tbl = table._tbl
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        borders = OxmlElement('w:tblBorders')
        for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            b = OxmlElement(f'w:{side}')
            b.set(qn('w:val'), 'nil')
            borders.append(b)
        tblPr.append(borders)

    def _spacer(self, doc: Document, points: float) -> None:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 0.1
        p.add_run().font.size = Pt(points)

    def _section_title(self, cell, title: str, icon: str | None = None, line_chars: str = '━━━━━━━━━━━━') -> None:
        p = self._cell_p(cell)
        self._run(p, title, bold=True, size=9.8, color=BLUE_2)
        p = cell.add_paragraph()
        self._pconf(p, after=2, line=0.6)
        self._run(p, line_chars, size=4, color=BLUE_2)
