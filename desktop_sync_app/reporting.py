from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.shared import Inches
from openpyxl import Workbook
from openpyxl.styles import Font

IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
VIDEO_EXTENSIONS = {'.mp4', '.mov', '.m4v', '.avi', '.mkv', '.webm'}


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


@dataclass
class ItemRecord:
    kind: str
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


class GrantProofReportEngine:
    def __init__(self, base_folder: Path) -> None:
        self.base_folder = Path(base_folder)
        self.projects_root = self.base_folder / 'projects'

    def ensure_root_files(self) -> None:
        self.base_folder.mkdir(parents=True, exist_ok=True)
        readme = self.base_folder / '_README_GrantProof.txt'
        readme.write_text(
            '\n'.join([
                'GrantProof export folder',
                '',
                'This folder is maintained by GrantProof Desktop Sync.',
                'Projects contain evidence and stories captured from the mobile application.',
                'Human-readable reports are available in each project under the reports folder.',
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
        ws.append(['Project code', 'Project name', 'Evidence', 'Stories', 'Last update'])
        self._header_style(ws[1])

        projects: dict[str, dict[str, object]] = {}
        for record in records:
            stats = projects.setdefault(record.project_code, {
                'name': record.project_name,
                'evidence': 0,
                'stories': 0,
                'last': '',
            })
            if record.kind == 'evidence':
                stats['evidence'] = int(stats['evidence']) + 1
            else:
                stats['stories'] = int(stats['stories']) + 1
            last_value = _safe_date(record.raw.get('syncedAt') or record.raw.get('createdAt'))
            if last_value and (not stats['last'] or str(last_value) > str(stats['last'])):
                stats['last'] = last_value

        for code in sorted(projects.keys()):
            stats = projects[code]
            ws.append([code, stats['name'], stats['evidence'], stats['stories'], stats['last']])

        details = wb.create_sheet('Items')
        details.append(['Kind', 'Project code', 'Project name', 'Title', 'Created at', 'Location', 'Folder'])
        self._header_style(details[1])
        for record in records:
            details.append([
                record.kind,
                record.project_code,
                record.project_name,
                record.title,
                record.created_at,
                record.location,
                record.relative_folder,
            ])

        self._autosize(ws)
        self._autosize(details)
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
                else:
                    description = _safe_text(raw.get('summary'), 'No summary')
                    location = _safe_text(raw.get('beneficiaryAlias'), '-')
                    item_kind = 'story'
                media_dir = metadata_path.parent / 'media'
                media_files = sorted([path for path in media_dir.iterdir() if path.is_file()]) if media_dir.exists() else []
                records.append(
                    ItemRecord(
                        kind=item_kind,
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
        wb = Workbook()
        overview = wb.active
        overview.title = 'Overview'
        overview.append(['Project code', project_code])
        overview.append(['Project name', project_name])
        overview.append(['Generated at', datetime.now().strftime('%Y-%m-%d %H:%M')])
        overview.append(['Evidence count', sum(1 for r in records if r.kind == 'evidence')])
        overview.append(['Story count', sum(1 for r in records if r.kind == 'story')])

        evidence_sheet = wb.create_sheet('Evidence')
        story_sheet = wb.create_sheet('Stories')
        for sheet in (evidence_sheet, story_sheet):
            sheet.append(['Title', 'Created at', 'Activity / Quote', 'Location / Beneficiary', 'Description / Summary', 'Folder'])
            self._header_style(sheet[1])

        for record in records:
            target = evidence_sheet if record.kind == 'evidence' else story_sheet
            extra = _safe_text(record.raw.get('activity') if record.kind == 'evidence' else record.raw.get('quote'))
            target.append([
                record.title,
                record.created_at,
                extra,
                record.location,
                record.description,
                record.relative_folder,
            ])

        self._autosize(overview)
        self._autosize(evidence_sheet)
        self._autosize(story_sheet)
        wb.save(output)

    def _build_project_report(self, output: Path, project_code: str, project_name: str, records: list[ItemRecord]) -> None:
        document = Document()
        document.add_heading(f'GrantProof Project Report – {project_code}', level=0)
        document.add_paragraph(project_name)
        document.add_paragraph(f'Generated on {datetime.now().strftime("%Y-%m-%d %H:%M")}.')

        evidence_records = [record for record in records if record.kind == 'evidence']
        story_records = [record for record in records if record.kind == 'story']

        document.add_heading('Summary', level=1)
        document.add_paragraph(f'Evidence items: {len(evidence_records)}')
        document.add_paragraph(f'Stories: {len(story_records)}')

        document.add_heading('Evidence', level=1)
        if not evidence_records:
            document.add_paragraph('No evidence synced yet.')
        for record in evidence_records:
            self._append_record_block(document, record)

        document.add_heading('Stories', level=1)
        if not story_records:
            document.add_paragraph('No stories synced yet.')
        for record in story_records:
            self._append_record_block(document, record)

        document.save(output)

    def _build_item_doc(self, record: ItemRecord, filename: str) -> None:
        output = record.metadata_path.parent / filename
        document = Document()
        document.add_heading(record.title, level=0)
        document.add_paragraph(f'Project: {record.project_name} ({record.project_code})')
        document.add_paragraph(f'Created: {record.created_at or "-"}')
        if record.kind == 'evidence':
            document.add_paragraph(f'Activity: {_safe_text(record.raw.get("activity"), "-")}')
            document.add_paragraph(f'Output: {_safe_text(record.raw.get("output"), "-")}')
            document.add_paragraph(f'Location: {_safe_text(record.raw.get("locationLabel"), "-")}')
        else:
            document.add_paragraph(f'Beneficiary: {_safe_text(record.raw.get("beneficiaryAlias"), "-")}')
            document.add_paragraph(f'Consent confirmed: {"Yes" if record.raw.get("consentGiven") else "No"}')
            if _safe_text(record.raw.get('quote')):
                document.add_paragraph(f'Quote: “{_safe_text(record.raw.get("quote"))}”')

        document.add_heading('Narrative', level=1)
        document.add_paragraph(record.description)

        image_files = [path for path in record.media_files if path.suffix.lower() in IMAGE_EXTENSIONS]
        video_files = [path for path in record.media_files if path.suffix.lower() in VIDEO_EXTENSIONS]

        if image_files:
            document.add_heading('Images', level=1)
            for image_path in image_files[:4]:
                try:
                    document.add_picture(str(image_path), width=Inches(5.5))
                except Exception:
                    document.add_paragraph(image_path.name)

        if video_files:
            document.add_heading('Videos', level=1)
            for video_path in video_files:
                document.add_paragraph(video_path.name, style='List Bullet')

        document.add_paragraph(f'Folder: {record.relative_folder}')
        document.save(output)

    def _append_record_block(self, document: Document, record: ItemRecord) -> None:
        document.add_heading(record.title, level=2)
        document.add_paragraph(f'Created: {record.created_at or "-"}')
        document.add_paragraph(record.description)
        if record.kind == 'evidence':
            document.add_paragraph(f'Activity: {_safe_text(record.raw.get("activity"), "-")}')
            document.add_paragraph(f'Output: {_safe_text(record.raw.get("output"), "-")}')
            document.add_paragraph(f'Location: {_safe_text(record.raw.get("locationLabel"), "-")}')
        else:
            document.add_paragraph(f'Beneficiary: {_safe_text(record.raw.get("beneficiaryAlias"), "-")}')
            if _safe_text(record.raw.get('quote')):
                document.add_paragraph(f'Quote: “{_safe_text(record.raw.get("quote"))}”')
        if record.media_files:
            document.add_paragraph(f'Media files: {len(record.media_files)}')
        document.add_paragraph(f'Folder: {record.relative_folder}')

    def _header_style(self, cells: Iterable) -> None:
        for cell in cells:
            cell.font = Font(bold=True)

    def _autosize(self, worksheet) -> None:
        for column_cells in worksheet.columns:
            length = max(len(_safe_text(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(length + 2, 12), 42)
