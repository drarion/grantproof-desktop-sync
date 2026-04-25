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
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from PIL import Image, ImageDraw, ImageFont, ImageOps

IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
VIDEO_EXTENSIONS = {'.mp4', '.mov', '.m4v', '.avi', '.mkv', '.webm'}

BLUE = '#064890'
BLUE_2 = '#0A57B5'
BLUE_3 = '#1665C1'
PALE_BLUE = '#EEF5FF'
SOFT_BLUE = '#F6FAFF'
ORANGE = '#F56D28'
DARK = '#1D2D44'
MUTED = '#5D6B7C'
BORDER = '#D9E4F2'
BORDER_LIGHT = '#E8EEF7'
WHITE = '#FFFFFF'

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

# Used only when the activity location is a well-known capital and no GPS is present.
# Otherwise, no point is plotted; the report shows the country map only, as requested.
CAPITAL_COORDS = {
    'Niger': (13.5116, 2.1254, 'Niamey'),
    'Burkina Faso': (12.3714, -1.5197, 'Ouagadougou'),
    'Mali': (12.6392, -8.0029, 'Bamako'),
    'Chad': (12.1348, 15.0557, "N'Djamena"),
    'Nigeria': (9.0765, 7.3986, 'Abuja'),
    'Senegal': (14.7167, -17.4677, 'Dakar'),
    'Cameroon': (3.8480, 11.5021, 'Yaoundé'),
    'S. Sudan': (4.8594, 31.5713, 'Juba'),
    'Sudan': (15.5007, 32.5599, 'Khartoum'),
}

SECTOR_RULES = [
    {
        'key': 'shelter', 'fr': 'Abris', 'en': 'Shelter', 'icon': 'shelter',
        'keywords': ['abri', 'abris', 'shelter', 'logement', 'construction', 'réhabilitation', 'rehabilitation', 'maison', 'habitat', 'ame', 'nfi'],
    },
    {
        'key': 'food_security', 'fr': 'Sécurité alimentaire', 'en': 'Food security', 'icon': 'food',
        'keywords': ['sécurité alimentaire', 'food security', 'vivres', 'food', 'ration', 'maraîcher', 'maraicher', 'semence', 'intrants', 'kit agricole', 'kits agricole', 'agricole', 'cash for food'],
    },
    {
        'key': 'agriculture', 'fr': 'Agriculture', 'en': 'Agriculture', 'icon': 'agriculture',
        'keywords': ['agriculture', 'agricole', 'maraîcher', 'maraicher', 'semence', 'irrigation', 'élevage', 'livelihood', 'moyens d’existence', 'moyens existence', 'agriculteur'],
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


def _rgb(hex_color: str) -> tuple[int, int, int]:
    hex_color = hex_color.strip().lstrip('#')
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


def _rgba(hex_color: str, alpha: int = 255) -> tuple[int, int, int, int]:
    return (*_rgb(hex_color), alpha)


class PremiumActivityReportBuilder:
    """High-fidelity generated DOCX report.

    The first page is rendered as a premium dashboard plate and inserted into a DOCX page.
    This prevents Word/LibreOffice from breaking the complex dashboard layout on different PCs.
    Additional images are added as a second visual annex when needed.
    """

    PAGE_W = 1600
    PAGE_H = 1200

    def __init__(self, base_folder: Path, default_org_name: str = '') -> None:
        self.base_folder = Path(base_folder)
        self.default_org_name = _safe_text(default_org_name)
        self.assets_dir = Path(__file__).resolve().parent / 'assets'
        self.maps_dir = self.assets_dir / 'maps' / 'naturalearth_lowres'
        self._shape_cache: list[Any] | None = None

    def build(self, output: Path, project_code: str, project_name: str, records: list[Any], lang: str) -> None:
        lang = 'en' if str(lang).lower().startswith('en') else 'fr'
        if not records:
            return
        data = self._prepare_data(project_code, project_name, records, lang)
        with tempfile.TemporaryDirectory(prefix='grantproof_premium_report_') as root:
            temp_dir = Path(root)
            page1 = self._render_first_page(data, temp_dir / 'page_1.png', lang)
            annex_pages = self._render_annex_pages(data, temp_dir, lang)
            doc = Document()
            self._configure_docx(doc)
            self._add_full_page_image(doc, page1)
            for page in annex_pages:
                doc.add_page_break()
                self._add_full_page_image(doc, page)
            doc.save(output)

    # ----------------------------- data preparation -----------------------------

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
        sectors = self._detect_sectors(records, activity, project_name, lang)
        quantity = self._extract_primary_quantity(records, sectors, lang)
        media_files = self._all_media(records)
        image_files = [path for path in media_files if Path(path).suffix.lower() in IMAGE_EXTENSIONS]
        hero_image = self._select_hero_image(image_files, primary_record)
        annex_images = [path for path in image_files if hero_image is None or Path(path) != Path(hero_image)]
        gps = self._extract_gps(primary_record, records)
        # If there is no GPS, use a point only for a recognized capital. Otherwise, map is country-only.
        map_point = gps or self._fallback_point(country, location)
        description = _safe_text(getattr(primary_record, 'description', '') or getattr(primary_record, 'raw', {}).get('description'))
        subtype = self._evidence_type_label(primary_record, lang)
        title = self._report_title(activity, location, country_display, lang)
        summary = self._narrative(activity, project_name, description, quantity, location, lang)
        highlights = self._highlights(activity, project_code, description, quantity, len(media_files), lang)
        target_group = self._target_group_label(quantity, lang)
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
            'media_count': len(media_files),
            'evidence_count': len(evidence_records) or len(records),
            'story_count': len(story_records),
            'hero_image': Path(hero_image) if hero_image else None,
            'annex_images': [Path(p) for p in annex_images],
            'map_point': map_point,
            'gps_available': gps is not None,
            'map_location_label': map_point[2] if map_point else country_display,
            'description': description,
            'summary': summary,
            'highlights': highlights,
            'target_group': target_group,
            'evidence_type': subtype,
        }

    # ----------------------------- DOCX container -----------------------------

    def _configure_docx(self, doc: Document) -> None:
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        # 4:3 page ratio, matching the validated visual proposal and avoiding A4-wide stretching.
        section.page_width = Inches(11)
        section.page_height = Inches(8.25)
        section.top_margin = Inches(0.02)
        section.bottom_margin = Inches(0.02)
        section.left_margin = Inches(0.02)
        section.right_margin = Inches(0.02)
        section.header_distance = Inches(0)
        section.footer_distance = Inches(0)
        styles = doc.styles
        styles['Normal'].font.name = 'Arial'
        styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
        styles['Normal'].font.size = Pt(9)

    def _add_full_page_image(self, doc: Document, image_path: Path) -> None:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(image_path), width=Inches(10.55), height=Inches(7.91))

    # ----------------------------- rendering -----------------------------

    def _render_first_page(self, data: dict[str, Any], output: Path, lang: str) -> Path:
        im = Image.new('RGB', (self.PAGE_W, self.PAGE_H), '#FFFFFF')
        draw = ImageDraw.Draw(im)

        # Header
        draw.rectangle((0, 0, self.PAGE_W, 172), fill=BLUE)
        self._draw_subtle_header_texture(draw)
        title_font = self._fit_font(data['title'].upper(), max_width=1120, start_size=42, min_size=30, bold=True, condensed=True)
        subtitle_font = self._font(22, bold=False, condensed=True)
        title_lines = self._wrap_text(draw, data['title'].upper(), title_font, 1120, max_lines=2)
        y = 33
        for line in title_lines:
            draw.text((34, y), line, fill=WHITE, font=title_font)
            y += self._text_h(draw, line, title_font) + 3
        subtitle = f"{'Projet' if lang == 'fr' else 'Project'} : {data['project_name']} ({data['project_code']})"
        draw.text((36, 112 if len(title_lines) == 1 else 132), subtitle, fill=WHITE, font=subtitle_font)
        self._draw_header_country(draw, data, (1160, 14, 1325, 148))
        draw.line((1350, 22, 1350, 150), fill='#D8E8FF', width=2)
        meta_font = self._font(21, bold=False, condensed=True)
        meta_y = 28
        for icon, text in [('person', data['org_name']), ('pin', data['location']), ('calendar', data['report_date'])]:
            self._draw_tiny_header_icon(draw, icon, (1385, meta_y + 3), '#FFFFFF')
            draw.text((1425, meta_y + 2), self._truncate(text, 160, meta_font), fill=WHITE, font=meta_font)
            meta_y += 42

        # Top grid
        self._draw_overview_card(draw, data, (30, 195, 485, 810), lang)
        self._draw_kpi_block(draw, data, (520, 195, 1070, 810), lang)
        self._draw_media_sector_block(draw, data, (1095, 195, 1570, 810), lang)
        draw.line((30, 835, 1570, 835), fill=BORDER_LIGHT, width=2)
        self._draw_highlights_card(draw, data, (30, 850, 765, 1170), lang)
        self._draw_location_card(draw, data, (790, 850, 1570, 1170), lang)

        output.parent.mkdir(parents=True, exist_ok=True)
        im.save(output, quality=95)
        return output

    def _draw_subtle_header_texture(self, draw: ImageDraw.ImageDraw) -> None:
        # Non-branded subtle dots, to avoid a flat block while staying sober.
        for x in range(500, 1180, 18):
            for y in range(20, 152, 18):
                if (x * 17 + y * 11) % 41 == 0:
                    draw.ellipse((x, y, x + 1, y + 1), fill='#1E6FC2')

    def _draw_overview_card(self, draw: ImageDraw.ImageDraw, data: dict[str, Any], box: tuple[int, int, int, int], lang: str) -> None:
        x1, y1, x2, y2 = box
        self._rounded_panel(draw, box, WHITE, BORDER)
        self._section_title(draw, (x1 + 25, y1 + 25), 'APERÇU DU PROJET' if lang == 'fr' else 'PROJECT OVERVIEW', 'doc', width=320)
        body_font = self._font(20)
        lines = self._wrap_text(draw, data['summary'], body_font, x2 - x1 - 55, max_lines=11)
        y = y1 + 92
        for line in lines:
            draw.text((x1 + 25, y), line, fill=DARK, font=body_font)
            y += 26
        draw.line((x1 + 25, y + 4, x2 - 25, y + 4), fill=BORDER, width=1)
        y += 18
        rows = [
            ('pin', 'Lieu' if lang == 'fr' else 'Location', data['location']),
            ('people', 'Public cible' if lang == 'fr' else 'Target group', data['target_group']),
            ('camera', 'Type de preuve' if lang == 'fr' else 'Evidence type', data['evidence_type']),
        ]
        label_font = self._font(18, bold=True)
        value_font = self._font(17)
        for icon, label, value in rows:
            self._draw_line_icon(draw, icon, (x1 + 55, y + 24), BLUE_2, scale=0.82)
            draw.text((x1 + 100, y + 5), label, fill=BLUE_2, font=label_font)
            draw.text((x1 + 100, y + 29), self._truncate(value, 300, value_font), fill=DARK, font=value_font)
            y += 68
            if y < y2 - 30:
                draw.line((x1 + 25, y - 6, x2 - 25, y - 6), fill=BORDER_LIGHT, width=1)

    def _draw_kpi_block(self, draw: ImageDraw.ImageDraw, data: dict[str, Any], box: tuple[int, int, int, int], lang: str) -> None:
        x1, y1, x2, y2 = box
        self._section_title(draw, (x1 + 0, y1 + 25), 'CHIFFRES CLÉS' if lang == 'fr' else 'KEY FIGURES', 'bar', width=500)
        card_w = 160
        gap = 22
        start_x = x1 + 0
        top = y1 + 135
        kpis = [
            ('people', data['quantity']['value'], data['quantity']['label']),
            ('media', str(data['media_count']), 'Médias liés' if lang == 'fr' else 'Linked media'),
            ('check', str(data['evidence_count']), 'Preuve\nconsolidée' if lang == 'fr' else 'Consolidated\nevidence'),
        ]
        for idx, (icon, value, label) in enumerate(kpis):
            cx = start_x + idx * (card_w + gap)
            self._rounded_panel(draw, (cx, top, cx + card_w, top + 440), WHITE, BORDER_LIGHT, radius=12, shadow=True)
            draw.ellipse((cx + 43, top + 47, cx + 117, top + 121), fill=PALE_BLUE)
            self._draw_line_icon(draw, icon, (cx + 80, top + 85), BLUE_2, scale=1.08)
            num_font = self._font(48, bold=True, condensed=True)
            label_font = self._font(19, condensed=True)
            w = self._text_w(draw, value, num_font)
            draw.text((cx + (card_w - w) / 2, top + 188), value, fill=BLUE_2, font=num_font)
            draw.line((cx + 64, top + 277, cx + 96, top + 277), fill=ORANGE, width=3)
            lines = label.split('\n') if '\n' in label else self._wrap_text(draw, label, label_font, card_w - 22, max_lines=2)
            ly = top + 305
            for line in lines:
                lw = self._text_w(draw, line, label_font)
                draw.text((cx + (card_w - lw) / 2, ly), line, fill=DARK, font=label_font)
                ly += 24

    def _draw_media_sector_block(self, draw: ImageDraw.ImageDraw, data: dict[str, Any], box: tuple[int, int, int, int], lang: str) -> None:
        x1, y1, x2, y2 = box
        hero_box = (x1, y1 + 5, x2, y1 + 405)
        self._draw_hero_image(draw, data.get('hero_image'), hero_box, data)
        self._section_title(draw, (x1 + 0, y1 + 430), 'ALIGNEMENT SECTORIEL' if lang == 'fr' else 'SECTOR ALIGNMENT', 'sector', width=405)
        sectors = data['sectors'][:2] or [{'key': 'coordination', 'label': 'Coordination', 'icon': 'coordination'}]
        item_w = (x2 - x1) // max(1, len(sectors))
        for idx, sector in enumerate(sectors):
            sx = x1 + idx * item_w
            if idx > 0:
                draw.line((sx, y1 + 505, sx, y2 - 5), fill=BORDER, width=2)
            self._draw_sector_icon(draw, sector['icon'], (sx + item_w // 2, y1 + 545), BLUE_2, scale=1.10)
            label_font = self._font(18)
            lines = self._wrap_text(draw, sector['label'], label_font, item_w - 35, max_lines=2)
            ly = y1 + 595
            for line in lines:
                lw = self._text_w(draw, line, label_font)
                draw.text((sx + (item_w - lw) / 2, ly), line, fill=DARK, font=label_font)
                ly += 23

    def _draw_highlights_card(self, draw: ImageDraw.ImageDraw, data: dict[str, Any], box: tuple[int, int, int, int], lang: str) -> None:
        x1, y1, x2, y2 = box
        self._rounded_panel(draw, box, SOFT_BLUE, BORDER_LIGHT)
        self._section_title(draw, (x1 + 25, y1 + 25), 'FAITS SAILLANTS' if lang == 'fr' else 'HIGHLIGHTS', 'star', width=450)
        font = self._font(18)
        y = y1 + 105
        for item in data['highlights'][:3]:
            draw.ellipse((x1 + 28, y + 8, x1 + 36, y + 16), fill=BLUE_2)
            lines = self._wrap_text(draw, item, font, x2 - x1 - 80, max_lines=2)
            ly = y
            for line in lines:
                draw.text((x1 + 55, ly), line, fill=DARK, font=font)
                ly += 23
            y = ly + 14

    def _draw_location_card(self, draw: ImageDraw.ImageDraw, data: dict[str, Any], box: tuple[int, int, int, int], lang: str) -> None:
        x1, y1, x2, y2 = box
        self._rounded_panel(draw, box, SOFT_BLUE, BORDER_LIGHT)
        self._section_title(draw, (x1 + 25, y1 + 25), 'LOCALISATION' if lang == 'fr' else 'LOCATION', 'pin_round', width=220)
        font_bold = self._font(17, bold=True)
        draw.text((x1 + 35, y1 + 170), 'Localisation de l’activité' if lang == 'fr' else 'Activity location', fill=BLUE_2, font=font_bold)
        draw.line((x1 + 35, y1 + 203, x1 + 85, y1 + 203), fill=ORANGE, width=3)
        map_img = self._render_map_image(data, (440, 245), lang)
        mx, my = x1 + 330, y1 + 15
        draw.rounded_rectangle((mx - 3, my - 3, mx + 443, my + 248), radius=12, fill='#FFFFFF', outline=BORDER)
        draw.bitmap((mx, my), map_img) if False else im_paste(draw, map_img, (mx, my))
        cap_font = self._font(18, bold=True)
        caption = data['map_location_label'] if data.get('map_point') else data['country_display']
        cw = self._text_w(draw, caption, cap_font)
        draw.text((mx + (440 - cw) / 2, y1 + 280), caption, fill=BLUE_2, font=cap_font)

    def _render_annex_pages(self, data: dict[str, Any], temp_dir: Path, lang: str) -> list[Path]:
        images = data.get('annex_images') or []
        if not images:
            return []
        pages: list[Path] = []
        for page_index, chunk_start in enumerate(range(0, len(images), 6), start=1):
            chunk = images[chunk_start:chunk_start + 6]
            im = Image.new('RGB', (self.PAGE_W, self.PAGE_H), '#FFFFFF')
            draw = ImageDraw.Draw(im)
            draw.rectangle((0, 0, self.PAGE_W, 145), fill=BLUE)
            title = 'ANNEXE VISUELLE' if lang == 'fr' else 'VISUAL ANNEX'
            draw.text((36, 42), title, fill=WHITE, font=self._font(38, bold=True, condensed=True))
            subtitle = 'Images complémentaires liées à l’activité' if lang == 'fr' else 'Additional images linked to the activity'
            draw.text((36, 92), subtitle, fill=WHITE, font=self._font(21, condensed=True))
            cols, rows = 3, 2
            card_w, card_h = 480, 405
            gap_x, gap_y = 35, 35
            start_x, start_y = 35, 185
            for idx, source in enumerate(chunk):
                col = idx % cols
                row = idx // cols
                x = start_x + col * (card_w + gap_x)
                y = start_y + row * (card_h + gap_y)
                self._rounded_panel(draw, (x, y, x + card_w, y + card_h), SOFT_BLUE, BORDER_LIGHT)
                try:
                    photo = self._cover_crop(Image.open(source).convert('RGB'), (card_w - 30, card_h - 75), radius=12)
                    im_paste(draw, photo, (x + 15, y + 15))
                except Exception:
                    draw.text((x + 30, y + 170), Path(source).name, fill=MUTED, font=self._font(18))
                caption = f"Image {chunk_start + idx + 1}"
                draw.text((x + 20, y + card_h - 45), caption, fill=BLUE_2, font=self._font(17, bold=True))
            page_path = temp_dir / f'annex_{page_index}.png'
            im.save(page_path, quality=95)
            pages.append(page_path)
        return pages

    # ----------------------------- visual helpers -----------------------------

    def _rounded_panel(self, draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], fill: str, outline: str, radius: int = 12, shadow: bool = False) -> None:
        x1, y1, x2, y2 = box
        if shadow:
            draw.rounded_rectangle((x1 + 4, y1 + 4, x2 + 4, y2 + 4), radius=radius, fill='#EEF1F5')
        draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=1)

    def _section_title(self, draw: ImageDraw.ImageDraw, xy: tuple[int, int], title: str, icon: str, width: int) -> None:
        x, y = xy
        if icon:
            draw.ellipse((x, y, x + 48, y + 48), fill=BLUE_2)
            self._draw_line_icon(draw, icon, (x + 24, y + 24), WHITE, scale=0.72)
            tx = x + 68
        else:
            tx = x
        font = self._font(24, bold=True, condensed=True)
        draw.text((tx, y + 6), title, fill=BLUE_2, font=font)
        draw.line((tx, y + 45, tx + width, y + 45), fill=BLUE_2, width=2)

    def _draw_hero_image(self, draw: ImageDraw.ImageDraw, hero_path: Path | None, box: tuple[int, int, int, int], data: dict[str, Any]) -> None:
        x1, y1, x2, y2 = box
        if hero_path and hero_path.exists():
            try:
                photo = self._cover_crop(Image.open(hero_path).convert('RGB'), (x2 - x1, y2 - y1), radius=12)
                im_paste(draw, photo, (x1, y1))
                return
            except Exception:
                pass
        # Integrated placeholder, not a blank box.
        draw.rounded_rectangle(box, radius=12, fill='#F1F5FA', outline=BORDER)
        self._draw_sector_icon(draw, data['sectors'][0]['icon'], ((x1 + x2) // 2, (y1 + y2) // 2 - 18), BLUE_2, scale=2.0)
        text = 'Visuel principal à insérer' if data else 'Visual asset'
        font = self._font(20, bold=True)
        tw = self._text_w(draw, text, font)
        draw.text(((x1 + x2 - tw) / 2, (y1 + y2) // 2 + 60), text, fill=MUTED, font=font)

    def _cover_crop(self, img: Image.Image, size: tuple[int, int], radius: int = 12) -> Image.Image:
        tw, th = size
        ratio = tw / th
        img = img.convert('RGB')
        if img.width / img.height > ratio:
            new_w = int(img.height * ratio)
            x = (img.width - new_w) // 2
            img = img.crop((x, 0, x + new_w, img.height))
        else:
            new_h = int(img.width / ratio)
            y = max(0, (img.height - new_h) // 2)
            img = img.crop((0, y, img.width, y + new_h))
        img = img.resize((tw, th), Image.Resampling.LANCZOS)
        mask = Image.new('L', (tw, th), 0)
        d = ImageDraw.Draw(mask)
        d.rounded_rectangle((0, 0, tw, th), radius=radius, fill=255)
        img = img.convert('RGBA')
        img.putalpha(mask)
        return img

    def _draw_header_country(self, draw: ImageDraw.ImageDraw, data: dict[str, Any], box: tuple[int, int, int, int]) -> None:
        country = self._render_country_silhouette(data, (170, 120), for_header=True)
        im_paste(draw, country, (box[0], box[1]))

    def _render_country_silhouette(self, data: dict[str, Any], size: tuple[int, int], for_header: bool = False) -> Image.Image:
        width, height = size
        img = Image.new('RGBA', (width, height), (255, 255, 255, 0))
        d = ImageDraw.Draw(img)
        shape = self._country_shape(data['country'])
        if not shape:
            return img
        polygons, bbox = self._shape_polygons(shape, data['country'])
        if not polygons:
            return img
        mapper = self._map_transform(bbox, width, height, 12)
        for poly in polygons:
            pts = [mapper(lon, lat) for lon, lat in poly]
            if len(pts) >= 3:
                d.polygon(pts, fill=(255, 255, 255, 238), outline=(255, 255, 255, 255))
        point = data.get('map_point')
        if point:
            lat, lon, _ = point
            if self._point_in_bbox(lon, lat, bbox):
                px, py = mapper(lon, lat)
                self._draw_pin(d, px, py, 10)
        return img

    def _render_map_image(self, data: dict[str, Any], size: tuple[int, int], lang: str) -> Image.Image:
        width, height = size
        img = Image.new('RGBA', (width, height), '#FFFFFF')
        draw = ImageDraw.Draw(img)
        target_shape = self._country_shape(data['country'])
        if not target_shape:
            draw.text((width // 2 - 60, height // 2), data['country_display'], fill='#7B8492', font=self._font(18))
            return img
        polygons, bbox = self._shape_polygons(target_shape, data['country'])
        if not polygons:
            return img
        minx, miny, maxx, maxy = bbox
        dx = max(maxx - minx, 0.01)
        dy = max(maxy - miny, 0.01)
        bbox_expanded = (minx - dx * 0.35, miny - dy * 0.30, maxx + dx * 0.35, maxy + dy * 0.30)
        mapper = self._map_transform(bbox_expanded, width, height, 24)

        # neighbors
        target_name = self._shape_country_name(data['country'])
        for sr in self._iter_shapes():
            shape = sr.shape
            sx1, sy1, sx2, sy2 = [float(v) for v in shape.bbox]
            if sx2 < bbox_expanded[0] or sx1 > bbox_expanded[2] or sy2 < bbox_expanded[1] or sy1 > bbox_expanded[3]:
                continue
            name = self._record_name(sr.record)
            is_target = name == target_name
            s_polys, _ = self._shape_polygons(shape, name)
            for poly in s_polys:
                if not poly:
                    continue
                pts = [mapper(lon, lat) for lon, lat in poly]
                fill = '#FFFFFF' if is_target else '#F2F4F7'
                outline = '#D4DAE3' if is_target else '#E1E5EC'
                if len(pts) >= 3:
                    draw.polygon(pts, fill=fill, outline=outline)
        # labels
        label_font = self._font(16)
        country_font = self._font(20, bold=True)
        cx, cy = mapper((minx + maxx) / 2, (miny + maxy) / 2)
        main_label = data['country_display'].upper()
        draw.text((cx - self._text_w(draw, main_label, country_font) / 2, cy - 10), main_label, fill='#6B7280', font=country_font)
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
            label = self._display_country(name, name, lang).upper()
            draw.text((lx - self._text_w(draw, label, label_font) / 2, ly - 6), label, fill='#9AA2AE', font=label_font)
        point = data.get('map_point')
        if point:
            lat, lon, label = point
            if self._point_in_bbox(lon, lat, bbox_expanded):
                px, py = mapper(lon, lat)
                self._draw_pin(draw, px, py, 18)
                draw.text((px + 15, py - 9), label or data['location'], fill=DARK, font=self._font(14, bold=True))
        return img

    def _draw_pin(self, draw: ImageDraw.ImageDraw, x: float, y: float, size: int) -> None:
        draw.ellipse((x - size * 0.55, y - size * 1.05, x + size * 0.55, y + size * 0.05), fill=ORANGE)
        draw.polygon([(x, y + size * 0.85), (x - size * 0.42, y - size * 0.05), (x + size * 0.42, y - size * 0.05)], fill=ORANGE)
        draw.ellipse((x - size * 0.2, y - size * 0.70, x + size * 0.2, y - size * 0.30), fill=WHITE)

    def _draw_tiny_header_icon(self, draw: ImageDraw.ImageDraw, icon: str, xy: tuple[int, int], color: str) -> None:
        x, y = xy
        if icon == 'person':
            draw.ellipse((x + 4, y, x + 22, y + 18), outline=color, width=2)
            draw.arc((x - 2, y + 18, x + 28, y + 46), 200, 340, fill=color, width=2)
        elif icon == 'pin':
            draw.ellipse((x + 2, y, x + 24, y + 22), outline=color, width=2)
            draw.polygon([(x + 13, y + 38), (x + 4, y + 18), (x + 22, y + 18)], outline=color)
        else:
            draw.rounded_rectangle((x + 1, y + 3, x + 27, y + 31), radius=3, outline=color, width=2)
            draw.line((x + 1, y + 11, x + 27, y + 11), fill=color, width=2)

    def _draw_line_icon(self, draw: ImageDraw.ImageDraw, icon: str, center: tuple[int, int], color: str, scale: float = 1.0) -> None:
        x, y = center
        s = 26 * scale
        c = color
        w = max(2, int(3 * scale))
        if icon in {'doc', 'document'}:
            draw.rectangle((x - s * .35, y - s * .55, x + s * .35, y + s * .55), outline=c, width=w)
            for yy in [-.25, 0, .25]:
                draw.line((x - s * .20, y + s * yy, x + s * .20, y + s * yy), fill=c, width=w)
        elif icon in {'bar', 'barchart'}:
            draw.rectangle((x - s*.55, y, x - s*.30, y + s*.55), fill=c)
            draw.rectangle((x - s*.12, y - s*.30, x + s*.13, y + s*.55), fill=c)
            draw.rectangle((x + s*.32, y - s*.65, x + s*.57, y + s*.55), fill=c)
        elif icon in {'people', 'person'}:
            draw.ellipse((x - s*.18, y - s*.52, x + s*.18, y - s*.16), outline=c, width=w)
            draw.arc((x - s*.55, y - s*.12, x + s*.55, y + s*.60), 190, 350, fill=c, width=w)
            draw.ellipse((x - s*.70, y - s*.32, x - s*.42, y - s*.04), outline=c, width=max(1, w-1))
            draw.ellipse((x + s*.42, y - s*.32, x + s*.70, y - s*.04), outline=c, width=max(1, w-1))
        elif icon in {'media', 'camera'}:
            draw.rounded_rectangle((x - s*.50, y - s*.35, x + s*.50, y + s*.35), radius=4, outline=c, width=w)
            if icon == 'camera':
                draw.ellipse((x - s*.20, y - s*.18, x + s*.20, y + s*.22), outline=c, width=w)
            else:
                draw.polygon([(x - s*.08, y - s*.22), (x - s*.08, y + s*.22), (x + s*.28, y)], fill=c)
        elif icon in {'check'}:
            draw.rounded_rectangle((x - s*.42, y - s*.58, x + s*.42, y + s*.55), radius=4, outline=c, width=w)
            draw.line((x - s*.20, y, x - s*.02, y + s*.20, x + s*.30, y - s*.24), fill=c, width=w)
        elif icon in {'pin', 'pin_round'}:
            draw.ellipse((x - s*.30, y - s*.55, x + s*.30, y + s*.05), outline=c, width=w)
            draw.polygon([(x, y + s*.70), (x - s*.26, y), (x + s*.26, y)], outline=c)
            draw.ellipse((x - s*.10, y - s*.36, x + s*.10, y - s*.16), fill=c)
        elif icon == 'star':
            pts = []
            for i in range(10):
                a = -math.pi / 2 + i * math.pi / 5
                rr = s * .55 if i % 2 == 0 else s * .25
                pts.append((x + rr * math.cos(a), y + rr * math.sin(a)))
            draw.line(pts + [pts[0]], fill=c, width=w)
        elif icon == 'sector':
            draw.ellipse((x - s*.50, y - s*.50, x + s*.50, y + s*.50), outline=c, width=w)
            draw.line((x, y - s*.65, x, y + s*.65), fill=c, width=w)
            draw.line((x - s*.65, y, x + s*.65, y), fill=c, width=w)
        else:
            draw.ellipse((x - 5, y - 5, x + 5, y + 5), fill=c)

    def _draw_sector_icon(self, draw: ImageDraw.ImageDraw, icon: str, center: tuple[int, int], color: str, scale: float = 1.0) -> None:
        x, y = center
        s = 36 * scale
        c = color
        w = max(2, int(3 * scale))
        if icon == 'shelter':
            draw.polygon([(x - s*.75, y), (x, y - s*.65), (x + s*.75, y)], outline=c)
            draw.rectangle((x - s*.52, y, x + s*.52, y + s*.58), outline=c, width=w)
            draw.rectangle((x - s*.15, y + s*.25, x + s*.15, y + s*.58), outline=c, width=max(1, w-1))
        elif icon in {'food', 'nutrition'}:
            draw.arc((x - s*.75, y - s*.15, x + s*.75, y + s*.75), 0, 180, fill=c, width=w)
            draw.line((x - s*.70, y + s*.30, x + s*.70, y + s*.30), fill=c, width=w)
            draw.line((x - s*.30, y - s*.55, x - s*.30, y + s*.05), fill=c, width=max(1, w-1))
            draw.ellipse((x - s*.42, y - s*.75, x - s*.18, y - s*.52), outline=c, width=max(1, w-1))
        elif icon == 'agriculture':
            draw.line((x, y + s*.75, x, y - s*.45), fill=c, width=w)
            draw.arc((x - s*.55, y - s*.55, x, y + s*.05), 210, 50, fill=c, width=w)
            draw.arc((x, y - s*.75, x + s*.65, y - s*.05), 130, 320, fill=c, width=w)
            draw.arc((x - s*.65, y - s*.05, x, y + s*.65), 210, 50, fill=c, width=w)
            draw.arc((x, y - s*.05, x + s*.70, y + s*.65), 130, 320, fill=c, width=w)
        elif icon == 'wash':
            draw.polygon([(x, y - s*.70), (x - s*.42, y + s*.05), (x, y + s*.65), (x + s*.42, y + s*.05)], outline=c)
        elif icon == 'health':
            draw.rectangle((x - s*.18, y - s*.65, x + s*.18, y + s*.65), fill=c)
            draw.rectangle((x - s*.65, y - s*.18, x + s*.65, y + s*.18), fill=c)
        elif icon == 'education':
            draw.polygon([(x - s*.75, y - s*.10), (x, y - s*.50), (x + s*.75, y - s*.10), (x, y + s*.28)], outline=c)
            draw.rectangle((x - s*.45, y + s*.10, x + s*.45, y + s*.55), outline=c, width=w)
        elif icon == 'protection':
            draw.polygon([(x, y - s*.75), (x + s*.60, y - s*.35), (x + s*.45, y + s*.55), (x, y + s*.80), (x - s*.45, y + s*.55), (x - s*.60, y - s*.35)], outline=c)
        elif icon == 'cash':
            draw.rounded_rectangle((x - s*.75, y - s*.42, x + s*.75, y + s*.42), radius=6, outline=c, width=w)
            draw.text((x - s*.13, y - s*.35), '$', fill=c, font=self._font(int(30 * scale), bold=True))
        else:
            self._draw_line_icon(draw, 'sector', center, color, scale)

    # ----------------------------- map helpers -----------------------------

    def _shape_file(self) -> Path:
        return self.maps_dir / 'naturalearth_lowres.shp'

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
                # Natural Earth lowres France includes overseas territories; keep metropolitan polygons only.
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

    # ----------------------------- extraction helpers -----------------------------

    def _project_payload(self, records: list[Any]) -> dict[str, Any]:
        for record in records:
            raw = getattr(record, 'raw', {}) or {}
            project = raw.get('project')
            if isinstance(project, dict) and project:
                return project
        return {}

    def _extract_org_name(self, project: dict[str, Any], records: list[Any]) -> str:
        # Broad list because mobile payloads have evolved across builds.
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
        # If the mobile build has not yet sent the organization name, avoid a false logo/name.
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
            project_name,
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
        seen: set[str] = set()
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
            _safe_text(getattr(record, 'description', '')) + ' ' +
            _safe_text(getattr(record, 'title', '')) + ' ' +
            _safe_text((getattr(record, 'raw', {}) or {}).get('activity'))
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
                score = min(w * h, 6_000_000) + (300_000 if w >= h else 0)
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
        # Keep country-only default maps. Plot capitals only when the location explicitly matches the capital.
        if country in CAPITAL_COORDS:
            lat, lon, capital = CAPITAL_COORDS[country]
            if capital.lower() in location.lower():
                return (lat, lon, capital)
        return None

    def _target_group_label(self, quantity: dict[str, str], lang: str) -> str:
        label = quantity['label'].lower()
        if 'abri' in label or 'shelter' in label:
            return 'Ménages / communautés bénéficiaires' if lang == 'fr' else 'Beneficiary households / communities'
        if 'école' in label or 'school' in label:
            return 'Élèves, enseignants et communautés scolaires' if lang == 'fr' else 'Students, teachers and school communities'
        if 'agriculteur' in label or 'farmer' in label:
            return 'Agriculteurs et communautés ciblées' if lang == 'fr' else 'Farmers and targeted communities'
        if 'bénéficiaire' in label or 'participant' in label:
            return 'Bénéficiaires de la formation' if lang == 'fr' else 'Training beneficiaries'
        return 'Communautés ciblées' if lang == 'fr' else 'Targeted communities'

    def _evidence_type_label(self, primary_record: Any, lang: str) -> str:
        subtype = _safe_text(getattr(primary_record, 'subtype', '') or getattr(primary_record, 'raw', {}).get('type')).lower()
        if 'photo' in subtype or 'image' in subtype:
            return 'Preuve photo' if lang == 'fr' else 'Photo evidence'
        if 'video' in subtype or 'vidéo' in subtype:
            return 'Preuve vidéo' if lang == 'fr' else 'Video evidence'
        if 'document' in subtype or 'doc' in subtype:
            return 'Document source' if lang == 'fr' else 'Source document'
        return 'Preuve terrain' if lang == 'fr' else 'Field evidence'

    def _narrative(self, activity: str, project_name: str, description: str, quantity: dict[str, str], location: str, lang: str) -> str:
        value = quantity['value']
        label = quantity['label'].lower()
        activity_clean = activity.lower()
        if lang == 'fr':
            intro = f"Cette activité de {activity_clean} a été documentée à {location} dans le cadre du projet {project_name}."
            if description:
                detail = f"La description source indique : {description.strip()}"
                detail = _sentence(detail)
            else:
                detail = ''
            metric = ''
            if value and label and value not in (intro + detail):
                metric = f"Le dossier met en évidence {value} {label}."
            support = "Les éléments collectés permettent d’illustrer la mise en œuvre de l’activité et de soutenir le reporting bailleur, sans aller au-delà des informations disponibles."
            return ' '.join([part for part in [intro, detail, metric, support] if part])
        intro = f"This {activity_clean} activity was documented in {location} under the project {project_name}."
        detail = _sentence(f"The source description states: {description.strip()}") if description else ''
        metric = f"The file highlights {value} {label}." if value and label else ''
        support = 'The collected elements illustrate implementation and support donor reporting without going beyond the available information.'
        return ' '.join([part for part in [intro, detail, metric, support] if part])

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

    # ----------------------------- text/layout helpers -----------------------------

    def _font(self, size: int, bold: bool = False, condensed: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
        candidates: list[str] = []
        if condensed:
            candidates += [
                r'C:\Windows\Fonts\arialnbi.ttf' if bold else r'C:\Windows\Fonts\arialn.ttf',
                r'C:\Windows\Fonts\ArialNarrow.ttf',
                '/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed-Bold.ttf' if bold else '/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf',
            ]
        candidates += [
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

    def _fit_font(self, text: str, max_width: int, start_size: int, min_size: int, bold: bool, condensed: bool = False):
        dummy = Image.new('RGB', (10, 10))
        d = ImageDraw.Draw(dummy)
        for size in range(start_size, min_size - 1, -1):
            font = self._font(size, bold=bold, condensed=condensed)
            if self._text_w(d, text, font) <= max_width:
                return font
        return self._font(min_size, bold=bold, condensed=condensed)

    def _wrap_text(self, draw: ImageDraw.ImageDraw, text: str, font, max_width: int, max_lines: int | None = None) -> list[str]:
        words = _clean_spaces(text).split()
        lines: list[str] = []
        current = ''
        for word in words:
            candidate = f'{current} {word}'.strip()
            if self._text_w(draw, candidate, font) <= max_width or not current:
                current = candidate
            else:
                lines.append(current)
                current = word
                if max_lines and len(lines) >= max_lines:
                    break
        if current and (not max_lines or len(lines) < max_lines):
            lines.append(current)
        if max_lines and len(lines) == max_lines and len(' '.join(words)) > len(' '.join(lines)):
            # ellipsis last line if needed
            last = lines[-1]
            while last and self._text_w(draw, last + '…', font) > max_width:
                last = last[:-1].rstrip()
            lines[-1] = last + '…'
        return lines

    def _truncate(self, text: str, max_width: int, font) -> str:
        dummy = Image.new('RGB', (10, 10))
        d = ImageDraw.Draw(dummy)
        if self._text_w(d, text, font) <= max_width:
            return text
        out = text
        while out and self._text_w(d, out + '…', font) > max_width:
            out = out[:-1]
        return out.rstrip() + '…'

    def _text_w(self, draw: ImageDraw.ImageDraw, text: str, font) -> int:
        try:
            return int(draw.textbbox((0, 0), text, font=font)[2])
        except Exception:
            return int(draw.textlength(text, font=font))

    def _text_h(self, draw: ImageDraw.ImageDraw, text: str, font) -> int:
        try:
            b = draw.textbbox((0, 0), text, font=font)
            return int(b[3] - b[1])
        except Exception:
            return 20


def im_paste(draw: ImageDraw.ImageDraw, source: Image.Image, xy: tuple[int, int]) -> None:
    # ImageDraw has no paste method; this helper uses the underlying image.
    target = getattr(draw, '_image', None)
    if target is not None:
        if source.mode == 'RGBA':
            target.paste(source, xy, source)
        else:
            target.paste(source, xy)
