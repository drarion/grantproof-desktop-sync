from __future__ import annotations

import json
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
VIDEO_EXTENSIONS = {'.mp4', '.mov', '.m4v', '.avi', '.mkv', '.webm'}

BLUE = '173B72'
ORANGE = 'F28C28'
SOFT_BLUE = 'EAF1FB'
SOFT_ORANGE = 'FFF2E2'
SOFT_GREEN = 'ECFDF3'
SOFT_GREY = 'F6F8FC'
MID_GREY = '6B7280'
DARK = '1F2937'

THIN_GREY_BORDER = Border(
    left=Side(style='thin', color='D1D9E6'),
    right=Side(style='thin', color='D1D9E6'),
    top=Side(style='thin', color='D1D9E6'),
    bottom=Side(style='thin', color='D1D9E6'),
)


def _safe_text(value: object, fallback: str = '') -> str:
    if value is None:
        return fallback
    text = str(value).strip()
    return text if text else fallback


def _safe_date(value: object) -> str:
    text = _safe_text(value)
    if not text:
        return ''
    try:
        return datetime.fromisoformat(text.replace('Z', '+00:00')).strftime('%Y-%m-%d %H:%M')
    except Exception:
        return text


def _parse_iso(value: object) -> datetime | None:
    text = _safe_text(value)
    if not text:
        return None
    try:
        return datetime.fromisoformat(text.replace('Z', '+00:00'))
    except Exception:
        return None


def _title_case(value: str) -> str:
    cleaned = _safe_text(value)
    return cleaned[:1].upper() + cleaned[1:] if cleaned else ''


def _sentence(value: str, fallback: str = '') -> str:
    text = ' '.join(_safe_text(value, fallback).replace('\n', ' ').split())
    if not text:
        return fallback
    if text[-1] not in '.!?':
        text += '.'
    return text


def _comma_join(items: list[str], fallback: str = '-') -> str:
    clean = [item for item in (_safe_text(i) for i in items) if item]
    return ', '.join(clean) if clean else fallback


@dataclass
class ItemRecord:
    kind: str
    subtype: str
    project_code: str
    project_name: str
    title: str
    description: str
    created_at: str
    location: str
    relative_folder: str
    metadata_path: Path
    media_files: list[Path]
    raw: dict


class GrantProofAIWriter:
    """Local, deterministic writing layer for donor-friendly narratives."""

    def executive_summary(self, project_name: str, donor_name: str, country: str, records: list[ItemRecord]) -> list[str]:
        evidence_records = [record for record in records if record.kind == 'evidence']
        story_records = [record for record in records if record.kind == 'story']
        outputs = [self._output_label(record) for record in evidence_records if self._output_label(record)]
        activities = [self._activity_label(record) for record in evidence_records if self._activity_label(record)]
        locations = [self._location_label(record) for record in records if self._location_label(record)]
        media_count = sum(len(record.media_files) for record in records)

        top_output = self._top_label(outputs, 'the expected project outputs')
        top_activity = self._top_label(activities, 'implementation activities')
        top_location = self._top_label(locations, 'field locations')

        donor_clause = f' under {donor_name}' if donor_name else ''
        country_clause = f' in {country}' if country else ''
        para1 = (
            f'This GrantProof reporting pack consolidates the latest field evidence collected for {project_name}{country_clause}{donor_clause}. '
            f'The current reporting window documents {len(evidence_records)} evidence item(s), {len(story_records)} story item(s), '
            f'and {media_count} supporting media asset(s), allowing the team to move from raw field capture to a structured donor-facing narrative.'
        )
        para2 = (
            f'The strongest concentration of documented activity relates to {top_activity}, with the clearest output signal around {top_output}. '
            f'Field material points to consistent implementation momentum and a diversified evidence base rather than isolated updates.'
        )
        para3 = (
            f'Across the available material, the most visible implementation footprint appears in {top_location}. '
            f'The overall portfolio suggests a credible chain between activity delivery, observable outputs, and beneficiary-facing storytelling that can be reused in donor reporting, internal briefs, and audit preparation.'
        )
        return [_sentence(para1), _sentence(para2), _sentence(para3)]

    def key_highlights(self, records: list[ItemRecord]) -> list[str]:
        evidence_records = [record for record in records if record.kind == 'evidence']
        story_records = [record for record in records if record.kind == 'story']
        activity_counts = Counter(self._activity_label(record) for record in evidence_records if self._activity_label(record))
        output_counts = Counter(self._output_label(record) for record in evidence_records if self._output_label(record))
        media_rich = sum(1 for record in records if record.media_files)
        gps_enabled = sum(1 for record in evidence_records if record.raw.get('latitude') is not None and record.raw.get('longitude') is not None)

        highlights = [
            f'{len(evidence_records)} evidence item(s) and {len(story_records)} story item(s) are already packaged for reuse in donor reporting.',
            f'{media_rich} captured item(s) include supporting media, strengthening visual verification and annex quality.',
        ]
        if activity_counts:
            activity, count = activity_counts.most_common(1)[0]
            highlights.append(f'The most documented activity stream is {activity} with {count} linked evidence item(s).')
        if output_counts:
            output, count = output_counts.most_common(1)[0]
            highlights.append(f'The strongest documented output signal is {output} with {count} linked evidence item(s).')
        if gps_enabled:
            highlights.append(f'{gps_enabled} evidence item(s) include GPS coordinates, which strengthens traceability for verification and audit purposes.')
        return [_sentence(item) for item in highlights[:5]]

    def implementation_summary(self, activity: str, records: list[ItemRecord]) -> str:
        outputs = [self._output_label(record) for record in records if self._output_label(record)]
        locations = [self._location_label(record) for record in records if self._location_label(record)]
        media_count = sum(len(record.media_files) for record in records)
        top_output = self._top_label(outputs, 'linked outputs')
        top_location = self._top_label(locations, 'documented field locations')
        return _sentence(
            f'The evidence grouped under {activity} shows {len(records)} field record(s) and {media_count} supporting media asset(s). '
            f'The material most often points toward {top_output}, with the clearest implementation footprint in {top_location}. '
            f'As a donor-facing narrative, this activity cluster demonstrates continuity between delivery on the ground and verifiable supporting material.'
        )

    def evidence_narrative(self, record: ItemRecord) -> str:
        activity = self._activity_label(record) or 'the planned activity stream'
        output = self._output_label(record) or 'the intended output'
        location = self._location_label(record) or 'the documented field site'
        description = _safe_text(record.description)
        subtitle = self._subtype_phrase(record)
        media_count = len(record.media_files)
        media_clause = f' and is supported by {media_count} media asset(s)' if media_count else ''
        if description:
            return _sentence(
                f'{subtitle} captured under {activity} in {location}. {description} '
                f'In donor-friendly terms, this evidence helps demonstrate concrete progress toward {output}{media_clause}.'
            )
        return _sentence(
            f'{subtitle} captured under {activity} in {location}. '
            f'Even with limited free-text detail, the item contributes a verifiable implementation signal linked to {output}{media_clause}.'
        )

    def story_narrative(self, record: ItemRecord) -> str:
        beneficiary = self._location_label(record) or 'the documented participant'
        summary = _safe_text(record.description)
        quote = _safe_text(record.raw.get('quote'))
        quote_clause = f' The recorded quote — “{quote}” — reinforces the human impact signal.' if quote else ''
        if summary:
            return _sentence(
                f'This beneficiary story centers on {beneficiary}. {summary} '
                f'For donor reporting, it provides qualitative evidence that complements the operational record and makes the project results more tangible.{quote_clause}'
            )
        return _sentence(
            f'This beneficiary story centers on {beneficiary}. '
            f'It adds a qualitative layer to the evidence base and helps translate project implementation into human-facing narrative.{quote_clause}'
        )

    def item_takeaways(self, record: ItemRecord) -> list[str]:
        bullets: list[str] = []
        if record.kind == 'evidence':
            activity = self._activity_label(record)
            output = self._output_label(record)
            if activity:
                bullets.append(f'Linked activity: {activity}.')
            if output:
                bullets.append(f'Illustrated output: {output}.')
            if record.raw.get('latitude') is not None and record.raw.get('longitude') is not None:
                bullets.append('GPS coordinates are available for traceability.')
        else:
            beneficiary = self._location_label(record)
            if beneficiary:
                bullets.append(f'Beneficiary perspective documented: {beneficiary}.')
            if record.raw.get('consentGiven'):
                bullets.append('Consent was confirmed at the time of capture.')
        if record.media_files:
            bullets.append(f'{len(record.media_files)} media asset(s) support this entry.')
        if not bullets:
            bullets.append('This item remains available as a documented reporting input.')
        return [_sentence(item) for item in bullets]

    def media_caption(self, record: ItemRecord, file_path: Path) -> str:
        base = file_path.stem.replace('_', ' ').replace('-', ' ').strip() or file_path.name
        if file_path.suffix.lower() in IMAGE_EXTENSIONS:
            return _sentence(f'Illustrative image linked to {record.title}: {base}.')
        if file_path.suffix.lower() in VIDEO_EXTENSIONS:
            return _sentence(f'Video evidence linked to {record.title}: {base}.')
        return _sentence(f'Linked supporting file for {record.title}: {base}.')

    def _activity_label(self, record: ItemRecord) -> str:
        return _safe_text(record.raw.get('activity'))

    def _output_label(self, record: ItemRecord) -> str:
        return _safe_text(record.raw.get('output'))

    def _location_label(self, record: ItemRecord) -> str:
        return _safe_text(record.raw.get('locationLabel') or record.raw.get('beneficiaryAlias'))

    def _top_label(self, values: list[str], fallback: str) -> str:
        clean = [value for value in values if value]
        if not clean:
            return fallback
        return Counter(clean).most_common(1)[0][0]

    def _subtype_phrase(self, record: ItemRecord) -> str:
        subtype = _safe_text(record.subtype)
        if subtype == 'attendance':
            return 'Attendance-focused evidence'
        if subtype == 'document':
            return 'Documentary evidence'
        if subtype == 'receipt':
            return 'Supporting receipt or justification'
        if subtype == 'note':
            return 'Narrative field note'
        return 'Field evidence'


class GrantProofReportEngine:
    def __init__(self, base_folder: Path) -> None:
        self.base_folder = Path(base_folder)
        self.projects_root = self.base_folder / 'projects'
        self.ai_writer = GrantProofAIWriter()

    def ensure_root_files(self) -> None:
        self.base_folder.mkdir(parents=True, exist_ok=True)
        readme = self.base_folder / '_README_GrantProof.txt'
        readme.write_text(
            '\n'.join([
                'GrantProof export folder',
                '',
                'This folder is maintained by GrantProof Desktop Sync.',
                'Projects contain evidence, stories and ready-to-share report outputs generated from mobile sync.',
                'Premium reports are available in each project under the reports folder.',
                '',
                'Main outputs:',
                '- Project_Register.xlsx',
                '- Project_Report.docx',
                '- evidence.docx / story.docx in each synced item folder',
            ]),
            encoding='utf-8',
        )
        self.rebuild_global_index()

    def rebuild_global_index(self) -> None:
        records = self._collect_all_records()
        wb = Workbook()
        ws = wb.active
        ws.title = 'Projects'
        ws.freeze_panes = 'A2'
        ws.append(['Project code', 'Project name', 'Evidence', 'Stories', 'Media assets', 'Last update'])
        self._header_style(ws[1], fill=BLUE)

        projects: dict[str, dict[str, object]] = {}
        for record in records:
            stats = projects.setdefault(record.project_code, {
                'name': record.project_name,
                'evidence': 0,
                'stories': 0,
                'media': 0,
                'last': '',
            })
            if record.kind == 'evidence':
                stats['evidence'] = int(stats['evidence']) + 1
            else:
                stats['stories'] = int(stats['stories']) + 1
            stats['media'] = int(stats['media']) + len(record.media_files)
            last_value = _safe_date(record.raw.get('syncedAt') or record.raw.get('createdAt'))
            if last_value and (not stats['last'] or str(last_value) > str(stats['last'])):
                stats['last'] = last_value

        for code in sorted(projects.keys()):
            stats = projects[code]
            ws.append([code, stats['name'], stats['evidence'], stats['stories'], stats['media'], stats['last']])

        details = wb.create_sheet('Items')
        details.freeze_panes = 'A2'
        details.append(['Kind', 'Subtype', 'Project code', 'Project name', 'Title', 'Created at', 'Location', 'Media', 'Folder'])
        self._header_style(details[1], fill=ORANGE)
        for record in records:
            details.append([
                record.kind,
                record.subtype,
                record.project_code,
                record.project_name,
                record.title,
                record.created_at,
                record.location,
                len(record.media_files),
                record.relative_folder,
            ])

        summary = wb.create_sheet('Summary')
        summary.append(['Metric', 'Value'])
        self._header_style(summary[1], fill=BLUE)
        summary.append(['Projects tracked', len(projects)])
        summary.append(['Evidence total', sum(int(item['evidence']) for item in projects.values())])
        summary.append(['Stories total', sum(int(item['stories']) for item in projects.values())])
        summary.append(['Media total', sum(int(item['media']) for item in projects.values())])
        summary.append(['Generated at', datetime.now().strftime('%Y-%m-%d %H:%M')])

        for sheet in (ws, details, summary):
            self._autosize(sheet)
            self._apply_sheet_polish(sheet)

        wb.save(self.base_folder / '_INDEX_GrantProof.xlsx')

    def rebuild_project(self, project_code: str) -> None:
        project_dir = self.projects_root / project_code
        if not project_dir.exists():
            return
        records = self._collect_project_records(project_code)
        if not records:
            return

        reports_dir = project_dir / 'reports'
        reports_dir.mkdir(parents=True, exist_ok=True)
        project_name = next((r.project_name for r in records if r.project_name), project_code)

        self._build_project_register(reports_dir / 'Project_Register.xlsx', project_code, project_name, records)
        self._build_project_report(reports_dir / 'Project_Report.docx', project_code, project_name, records)
        for record in records:
            if record.kind == 'evidence':
                self._build_item_doc(record, 'evidence.docx')
            else:
                self._build_item_doc(record, 'story.docx')

        self.rebuild_global_index()

    def rebuild_all(self) -> None:
        self.ensure_root_files()
        if not self.projects_root.exists():
            return
        for project_dir in self.projects_root.iterdir():
            if project_dir.is_dir():
                self.rebuild_project(project_dir.name)

    def _collect_all_records(self) -> list[ItemRecord]:
        records: list[ItemRecord] = []
        if not self.projects_root.exists():
            return records
        for project_dir in self.projects_root.iterdir():
            if project_dir.is_dir():
                records.extend(self._collect_project_records(project_dir.name))
        return sorted(records, key=lambda item: (item.project_code, item.created_at, item.title))

    def _collect_project_records(self, project_code: str) -> list[ItemRecord]:
        project_dir = self.projects_root / project_code
        if not project_dir.exists():
            return []
        records: list[ItemRecord] = []
        for kind in ('evidence', 'stories'):
            kind_dir = project_dir / kind
            if not kind_dir.exists():
                continue
            for metadata_path in kind_dir.rglob('*.json'):
                if metadata_path.name not in {'evidence.json', 'story.json'}:
                    continue
                try:
                    raw = json.loads(metadata_path.read_text(encoding='utf-8'))
                except Exception:
                    continue
                project = raw.get('project') or {}
                created_at = _safe_date(raw.get('createdAt'))
                title = _safe_text(raw.get('title'), metadata_path.parent.name)
                if kind == 'evidence':
                    description = _safe_text(raw.get('description'), 'No description')
                    location = _safe_text(raw.get('locationLabel'), '-')
                    item_kind = 'evidence'
                    subtype = _safe_text(raw.get('type'), 'note')
                else:
                    description = _safe_text(raw.get('summary'), 'No summary')
                    location = _safe_text(raw.get('beneficiaryAlias'), '-')
                    item_kind = 'story'
                    subtype = 'story'
                media_dir = metadata_path.parent / 'media'
                media_files = sorted([path for path in media_dir.iterdir() if path.is_file()]) if media_dir.exists() else []
                records.append(
                    ItemRecord(
                        kind=item_kind,
                        subtype=subtype,
                        project_code=project_code,
                        project_name=_safe_text(project.get('name'), project_code),
                        title=title,
                        description=description,
                        created_at=created_at,
                        location=location,
                        relative_folder=str(metadata_path.parent.relative_to(self.base_folder)),
                        metadata_path=metadata_path,
                        media_files=media_files,
                        raw=raw,
                    )
                )
        return sorted(records, key=lambda item: (item.created_at, item.title))

    def _build_project_register(self, output: Path, project_code: str, project_name: str, records: list[ItemRecord]) -> None:
        project_meta = self._project_meta(records)
        wb = Workbook()
        overview = wb.active
        overview.title = 'Overview'
        overview.append(['Metric', 'Value'])
        self._header_style(overview[1], fill=BLUE)
        overview.append(['Project code', project_code])
        overview.append(['Project name', project_name])
        overview.append(['Donor', project_meta['donor_name']])
        overview.append(['Country', project_meta['country']])
        overview.append(['Reporting window', project_meta['reporting_window']])
        overview.append(['Generated at', datetime.now().strftime('%Y-%m-%d %H:%M')])
        overview.append(['Evidence count', sum(1 for r in records if r.kind == 'evidence')])
        overview.append(['Story count', sum(1 for r in records if r.kind == 'story')])
        overview.append(['Media asset count', sum(len(r.media_files) for r in records)])
        overview.append(['Top activity', self._top_value([_safe_text(r.raw.get('activity')) for r in records if r.kind == 'evidence'])])
        overview.append(['Top output', self._top_value([_safe_text(r.raw.get('output')) for r in records if r.kind == 'evidence'])])

        dashboard = wb.create_sheet('Dashboard')
        dashboard.append(['Dimension', 'Value'])
        self._header_style(dashboard[1], fill=ORANGE)
        for label, value in self._dashboard_rows(records):
            dashboard.append([label, value])

        timeline = wb.create_sheet('Timeline')
        timeline.freeze_panes = 'A2'
        timeline.append(['Date', 'Kind', 'Subtype', 'Title', 'Activity', 'Output / Quote', 'Location / Beneficiary', 'Media', 'Folder'])
        self._header_style(timeline[1], fill=BLUE)
        for record in records:
            timeline.append([
                record.created_at,
                record.kind,
                record.subtype,
                record.title,
                _safe_text(record.raw.get('activity')),
                _safe_text(record.raw.get('output') if record.kind == 'evidence' else record.raw.get('quote')),
                record.location,
                len(record.media_files),
                record.relative_folder,
            ])

        evidence_sheet = wb.create_sheet('Evidence')
        evidence_sheet.freeze_panes = 'A2'
        evidence_sheet.append(['Title', 'Created at', 'Type', 'Activity', 'Output', 'Location', 'Description', 'Media', 'GPS', 'Folder'])
        self._header_style(evidence_sheet[1], fill=BLUE)

        story_sheet = wb.create_sheet('Stories')
        story_sheet.freeze_panes = 'A2'
        story_sheet.append(['Title', 'Created at', 'Beneficiary', 'Summary', 'Quote', 'Consent', 'Media', 'Folder'])
        self._header_style(story_sheet[1], fill=ORANGE)

        media_sheet = wb.create_sheet('Media')
        media_sheet.freeze_panes = 'A2'
        media_sheet.append(['Linked item', 'Kind', 'File name', 'Type', 'Caption', 'Folder'])
        self._header_style(media_sheet[1], fill=BLUE)

        for record in records:
            if record.kind == 'evidence':
                evidence_sheet.append([
                    record.title,
                    record.created_at,
                    record.subtype,
                    _safe_text(record.raw.get('activity')),
                    _safe_text(record.raw.get('output')),
                    record.location,
                    record.description,
                    len(record.media_files),
                    self._gps_value(record),
                    record.relative_folder,
                ])
            else:
                story_sheet.append([
                    record.title,
                    record.created_at,
                    record.location,
                    record.description,
                    _safe_text(record.raw.get('quote')),
                    'Yes' if record.raw.get('consentGiven') else 'No',
                    len(record.media_files),
                    record.relative_folder,
                ])
            for media_file in record.media_files:
                media_sheet.append([
                    record.title,
                    record.kind,
                    media_file.name,
                    'Image' if media_file.suffix.lower() in IMAGE_EXTENSIONS else ('Video' if media_file.suffix.lower() in VIDEO_EXTENSIONS else 'File'),
                    self.ai_writer.media_caption(record, media_file),
                    record.relative_folder,
                ])

        for sheet in (overview, dashboard, timeline, evidence_sheet, story_sheet, media_sheet):
            self._autosize(sheet)
            self._apply_sheet_polish(sheet)

        wb.save(output)

    def _build_project_report(self, output: Path, project_code: str, project_name: str, records: list[ItemRecord]) -> None:
        document = Document()
        self._configure_document(document)
        meta = self._project_meta(records)
        evidence_records = [record for record in records if record.kind == 'evidence']
        story_records = [record for record in records if record.kind == 'story']

        self._cover_block(document, project_code, project_name, meta, records)

        document.add_page_break()
        self._add_section_title(document, 'Executive summary')
        for paragraph in self.ai_writer.executive_summary(project_name, meta['donor_name'], meta['country'], records):
            self._body_paragraph(document, paragraph)

        self._add_section_title(document, 'Portfolio snapshot')
        summary_table = document.add_table(rows=0, cols=2)
        self._set_table_style(summary_table)
        for label, value in [
            ('Donor', meta['donor_name'] or 'Not specified'),
            ('Country', meta['country'] or 'Not specified'),
            ('Reporting window', meta['reporting_window']),
            ('Evidence items', str(len(evidence_records))),
            ('Story items', str(len(story_records))),
            ('Media assets', str(sum(len(r.media_files) for r in records))),
            ('Top activity', self._top_value([_safe_text(r.raw.get('activity')) for r in evidence_records])),
            ('Top output', self._top_value([_safe_text(r.raw.get('output')) for r in evidence_records])),
        ]:
            row = summary_table.add_row().cells
            row[0].text = label
            row[1].text = value
        self._shade_first_column(summary_table, SOFT_BLUE)

        self._add_section_title(document, 'Key highlights for donor reporting')
        for bullet in self.ai_writer.key_highlights(records):
            self._bullet_paragraph(document, bullet)

        self._add_section_title(document, 'Implementation narrative by activity')
        activity_groups: dict[str, list[ItemRecord]] = defaultdict(list)
        for record in evidence_records:
            activity_groups[_safe_text(record.raw.get('activity'), 'Unspecified activity')].append(record)
        if not activity_groups:
            self._body_paragraph(document, 'No evidence has been synchronized yet for this project.')
        for activity, grouped_records in sorted(activity_groups.items(), key=lambda item: (-len(item[1]), item[0])):
            self._subheading(document, activity)
            self._body_paragraph(document, self.ai_writer.implementation_summary(activity, grouped_records))
            outputs = sorted({value for value in (_safe_text(r.raw.get('output')) for r in grouped_records) if value})
            locations = sorted({value for value in (_safe_text(r.raw.get('locationLabel')) for r in grouped_records) if value})
            detail_table = document.add_table(rows=0, cols=2)
            self._set_table_style(detail_table)
            for label, value in [
                ('Evidence items', str(len(grouped_records))),
                ('Supporting media', str(sum(len(r.media_files) for r in grouped_records))),
                ('Outputs covered', _comma_join(outputs)),
                ('Locations', _comma_join(locations)),
            ]:
                row = detail_table.add_row().cells
                row[0].text = label
                row[1].text = value
            self._shade_first_column(detail_table, SOFT_ORANGE)
            document.add_paragraph('')

        self._add_section_title(document, 'Evidence highlights')
        if not evidence_records:
            self._body_paragraph(document, 'No evidence items are available yet.')
        for index, record in enumerate(self._prioritized_records(evidence_records)[:8], start=1):
            self._record_showcase(document, record, index=index)

        self._add_section_title(document, 'Beneficiary stories and qualitative signals')
        if not story_records:
            self._body_paragraph(document, 'No beneficiary stories are available yet.')
        for index, record in enumerate(self._prioritized_records(story_records)[:6], start=1):
            self._record_showcase(document, record, index=index)

        self._add_section_title(document, 'Media annex')
        media_table = document.add_table(rows=1, cols=4)
        self._set_table_style(media_table)
        headers = ['Item', 'Kind', 'Media file', 'Interpretive caption']
        for cell, header in zip(media_table.rows[0].cells, headers):
            cell.text = header
        self._shade_row(media_table.rows[0], BLUE, 'FFFFFF')
        media_rows = 0
        for record in self._prioritized_records(records):
            for media_file in record.media_files:
                row = media_table.add_row().cells
                row[0].text = record.title
                row[1].text = record.kind
                row[2].text = media_file.name
                row[3].text = self.ai_writer.media_caption(record, media_file)
                media_rows += 1
        if media_rows == 0:
            row = media_table.add_row().cells
            row[0].text = 'No media annex available'
            row[1].text = '-'
            row[2].text = '-'
            row[3].text = 'No synced media available for this reporting window.'

        document.save(output)

    def _build_item_doc(self, record: ItemRecord, filename: str) -> None:
        output = record.metadata_path.parent / filename
        document = Document()
        self._configure_document(document)
        self._cover_title(document, record.title)
        self._subtitle_line(document, f'{record.project_name} ({record.project_code})')
        self._meta_strip(document, [
            ('Created', record.created_at or '-'),
            ('Kind', _title_case(record.kind)),
            ('Type', _title_case(record.subtype)),
            ('Linked folder', record.relative_folder),
        ])

        self._add_section_title(document, 'Donor-facing narrative')
        narrative = self.ai_writer.evidence_narrative(record) if record.kind == 'evidence' else self.ai_writer.story_narrative(record)
        self._body_paragraph(document, narrative)

        self._add_section_title(document, 'Key takeaways')
        for item in self.ai_writer.item_takeaways(record):
            self._bullet_paragraph(document, item)

        self._add_section_title(document, 'Source details')
        source_table = document.add_table(rows=0, cols=2)
        self._set_table_style(source_table)
        details = [
            ('Project', f'{record.project_name} ({record.project_code})'),
            ('Created', record.created_at or '-'),
            ('Description / Summary', record.description or '-'),
        ]
        if record.kind == 'evidence':
            details.extend([
                ('Activity', _safe_text(record.raw.get('activity'), '-')),
                ('Output', _safe_text(record.raw.get('output'), '-')),
                ('Location', _safe_text(record.raw.get('locationLabel'), '-')),
                ('GPS', self._gps_value(record)),
            ])
        else:
            details.extend([
                ('Beneficiary alias', _safe_text(record.raw.get('beneficiaryAlias'), '-')),
                ('Consent confirmed', 'Yes' if record.raw.get('consentGiven') else 'No'),
                ('Quote', _safe_text(record.raw.get('quote'), '-')),
            ])
        for label, value in details:
            row = source_table.add_row().cells
            row[0].text = label
            row[1].text = value
        self._shade_first_column(source_table, SOFT_BLUE)

        image_files = [path for path in record.media_files if path.suffix.lower() in IMAGE_EXTENSIONS]
        video_files = [path for path in record.media_files if path.suffix.lower() in VIDEO_EXTENSIONS]

        if image_files:
            self._add_section_title(document, 'Illustrative images')
            for image_path in image_files[:4]:
                try:
                    document.add_picture(str(image_path), width=Inches(5.7))
                    caption = document.add_paragraph()
                    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = caption.add_run(self.ai_writer.media_caption(record, image_path))
                    run.italic = True
                except Exception:
                    self._body_paragraph(document, image_path.name)

        if video_files:
            self._add_section_title(document, 'Linked videos')
            for video_path in video_files:
                self._bullet_paragraph(document, self.ai_writer.media_caption(record, video_path))

        document.save(output)

    def _record_showcase(self, document: Document, record: ItemRecord, index: int) -> None:
        label = 'Evidence' if record.kind == 'evidence' else 'Story'
        self._subheading(document, f'{label} {index:02d} — {record.title}')
        meta_table = document.add_table(rows=0, cols=2)
        self._set_table_style(meta_table)
        rows = [
            ('Created', record.created_at or '-'),
            ('Location / Beneficiary', record.location or '-'),
            ('Media assets', str(len(record.media_files))),
        ]
        if record.kind == 'evidence':
            rows.extend([
                ('Activity', _safe_text(record.raw.get('activity'), '-')),
                ('Output', _safe_text(record.raw.get('output'), '-')),
            ])
        else:
            rows.append(('Consent', 'Yes' if record.raw.get('consentGiven') else 'No'))
        for label_text, value in rows:
            row = meta_table.add_row().cells
            row[0].text = label_text
            row[1].text = value
        self._shade_first_column(meta_table, SOFT_GREY)
        self._body_paragraph(document, self.ai_writer.evidence_narrative(record) if record.kind == 'evidence' else self.ai_writer.story_narrative(record))
        for bullet in self.ai_writer.item_takeaways(record):
            self._bullet_paragraph(document, bullet)
        image_files = [path for path in record.media_files if path.suffix.lower() in IMAGE_EXTENSIONS]
        if image_files:
            try:
                document.add_picture(str(image_files[0]), width=Inches(5.3))
            except Exception:
                pass
        quote = _safe_text(record.raw.get('quote'))
        if quote:
            quote_para = document.add_paragraph()
            quote_para.style = document.styles['Quote']
            quote_para.add_run(f'“{quote}”')
        document.add_paragraph('')

    def _project_meta(self, records: list[ItemRecord]) -> dict[str, str]:
        project = {}
        for record in records:
            candidate = record.raw.get('project')
            if isinstance(candidate, dict) and candidate:
                project = candidate
                break
        start_values = [dt for dt in (_parse_iso(r.raw.get('createdAt')) for r in records) if dt is not None]
        reporting_window = '-'
        if start_values:
            reporting_window = f'{min(start_values).strftime("%Y-%m-%d")} to {max(start_values).strftime("%Y-%m-%d")}'
        return {
            'donor_name': _safe_text(project.get('donorName')),
            'country': _safe_text(project.get('country')),
            'reporting_window': reporting_window,
        }

    def _dashboard_rows(self, records: list[ItemRecord]) -> list[tuple[str, str]]:
        evidence_records = [record for record in records if record.kind == 'evidence']
        story_records = [record for record in records if record.kind == 'story']
        activity_counts = Counter(_safe_text(record.raw.get('activity')) for record in evidence_records if _safe_text(record.raw.get('activity')))
        output_counts = Counter(_safe_text(record.raw.get('output')) for record in evidence_records if _safe_text(record.raw.get('output')))
        location_counts = Counter(record.location for record in records if record.location and record.location != '-')
        rows = [
            ('Evidence items', str(len(evidence_records))),
            ('Story items', str(len(story_records))),
            ('Media assets', str(sum(len(record.media_files) for record in records))),
            ('GPS-enabled evidence', str(sum(1 for record in evidence_records if record.raw.get('latitude') is not None and record.raw.get('longitude') is not None))),
        ]
        for activity, count in activity_counts.most_common(5):
            rows.append((f'Activity · {activity}', str(count)))
        for output, count in output_counts.most_common(5):
            rows.append((f'Output · {output}', str(count)))
        for location, count in location_counts.most_common(5):
            rows.append((f'Location · {location}', str(count)))
        return rows

    def _prioritized_records(self, records: list[ItemRecord]) -> list[ItemRecord]:
        def sort_key(record: ItemRecord):
            parsed = _parse_iso(record.raw.get('createdAt'))
            timestamp = parsed.timestamp() if parsed else 0
            return (len(record.media_files), timestamp, len(_safe_text(record.description)), record.title.lower())
        return sorted(records, key=sort_key, reverse=True)

    def _gps_value(self, record: ItemRecord) -> str:
        lat = record.raw.get('latitude')
        lon = record.raw.get('longitude')
        if lat is None or lon is None:
            return '-'
        return f'{lat}, {lon}'

    def _top_value(self, values: list[str], fallback: str = 'Not specified') -> str:
        clean = [value for value in values if value]
        if not clean:
            return fallback
        return Counter(clean).most_common(1)[0][0]

    def _configure_document(self, document: Document) -> None:
        section = document.sections[0]
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        styles = document.styles
        styles['Normal'].font.name = 'Aptos'
        styles['Normal'].font.size = Pt(10.5)
        styles['Heading 1'].font.name = 'Aptos Display'
        styles['Heading 1'].font.size = Pt(18)
        styles['Heading 1'].font.bold = True
        styles['Heading 1'].font.color.rgb = RGBColor.from_string(BLUE)
        styles['Heading 2'].font.name = 'Aptos Display'
        styles['Heading 2'].font.size = Pt(13)
        styles['Heading 2'].font.bold = True
        styles['Heading 2'].font.color.rgb = RGBColor.from_string(BLUE)
        styles['Quote'].font.italic = True
        styles['Quote'].font.color.rgb = RGBColor.from_string(DARK)

    def _cover_block(self, document: Document, project_code: str, project_name: str, meta: dict[str, str], records: list[ItemRecord]) -> None:
        self._cover_title(document, 'GrantProof Premium Project Report')
        self._subtitle_line(document, project_name)
        self._subtitle_line(document, f'Project code: {project_code}')
        self._subtitle_line(document, f'Donor: {meta["donor_name"] or "Not specified"} • Country: {meta["country"] or "Not specified"}')
        self._subtitle_line(document, f'Reporting window: {meta["reporting_window"]}')
        self._subtitle_line(document, f'Generated on {datetime.now().strftime("%Y-%m-%d %H:%M")}.')
        document.add_paragraph('')

        hero_table = document.add_table(rows=2, cols=3)
        self._set_table_style(hero_table)
        hero_values = [
            ('Evidence items', str(sum(1 for r in records if r.kind == 'evidence'))),
            ('Stories', str(sum(1 for r in records if r.kind == 'story'))),
            ('Media assets', str(sum(len(r.media_files) for r in records))),
            ('Top activity', self._top_value([_safe_text(r.raw.get('activity')) for r in records if r.kind == 'evidence'])),
            ('Top output', self._top_value([_safe_text(r.raw.get('output')) for r in records if r.kind == 'evidence'])),
            ('Local archive', 'GrantProof / projects / reports'),
        ]
        for index, (label, value) in enumerate(hero_values):
            cell = hero_table.cell(index // 3, index % 3)
            cell.text = f'{label}\n{value}'
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for row in hero_table.rows:
            for cell in row.cells:
                self._shade_cell(cell, SOFT_BLUE)
        document.add_paragraph('')

    def _cover_title(self, document: Document, text: str) -> None:
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(text)
        run.bold = True
        run.font.size = Pt(24)
        run.font.color.rgb = RGBColor.from_string(BLUE)

    def _subtitle_line(self, document: Document, text: str) -> None:
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(text)
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor.from_string(DARK)

    def _meta_strip(self, document: Document, pairs: list[tuple[str, str]]) -> None:
        table = document.add_table(rows=0, cols=2)
        self._set_table_style(table)
        for label, value in pairs:
            row = table.add_row().cells
            row[0].text = label
            row[1].text = value
        self._shade_first_column(table, SOFT_BLUE)
        document.add_paragraph('')

    def _add_section_title(self, document: Document, text: str) -> None:
        document.add_heading(text, level=1)

    def _subheading(self, document: Document, text: str) -> None:
        document.add_heading(text, level=2)

    def _body_paragraph(self, document: Document, text: str) -> None:
        paragraph = document.add_paragraph(text)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.space_after = Pt(6)

    def _bullet_paragraph(self, document: Document, text: str) -> None:
        paragraph = document.add_paragraph(style='List Bullet')
        paragraph.add_run(text)

    def _set_table_style(self, table) -> None:
        table.style = 'Table Grid'
        table.autofit = True

    def _shade_first_column(self, table, fill_hex: str) -> None:
        for row in table.rows:
            self._shade_cell(row.cells[0], fill_hex)
            self._bold_paragraphs(row.cells[0])

    def _shade_row(self, row, fill_hex: str, text_hex: str) -> None:
        for cell in row.cells:
            self._shade_cell(cell, fill_hex)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.color.rgb = RGBColor.from_string(text_hex)
                    run.bold = True

    def _shade_cell(self, cell, fill_hex: str) -> None:
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), fill_hex)
        tc_pr.append(shd)

    def _bold_paragraphs(self, cell) -> None:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.color.rgb = RGBColor.from_string(BLUE)

    def _header_style(self, cells: Iterable, fill: str = BLUE) -> None:
        for cell in cells:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor=fill)
            cell.alignment = Alignment(vertical='top', wrap_text=True)
            cell.border = THIN_GREY_BORDER

    def _autosize(self, worksheet) -> None:
        for column_cells in worksheet.columns:
            max_length = 0
            for cell in column_cells:
                value = _safe_text(cell.value)
                for line in value.splitlines() or ['']:
                    max_length = max(max_length, len(line))
            worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = min(max(max_length + 2, 14), 42)

    def _apply_sheet_polish(self, worksheet) -> None:
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical='top', wrap_text=True)
                cell.border = THIN_GREY_BORDER
        worksheet.sheet_view.showGridLines = False
