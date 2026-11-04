"""Private, read-only Streamlit UI service; no real workbook update capability.

Source Excel bytes are written only to a temporary directory and immediately
removed after checking. Original ChannelWise workbook files are read, never saved.
"""
from __future__ import annotations

from datetime import date
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile, is_zipfile
import json

from src.real_business_dry_run import audit, BRANDS
from scripts.export_verified_channelwise_copies import export_copies

MAX_UPLOAD_BYTES = 40 * 1024 * 1024


def validate_month(month: str) -> str:
    if not isinstance(month, str) or len(month) != 7 or month[4] != '-':
        raise ValueError('Reporting month must be YYYY-MM')
    try:
        date.fromisoformat(month + '-01')
    except ValueError as exc:
        raise ValueError('Invalid reporting month') from exc
    return month


def _check_xlsx(data: bytes, description: str) -> None:
    if not data or len(data) > MAX_UPLOAD_BYTES:
        raise ValueError(f'{description} file is empty or exceeds 40 MB')
    with BytesIO(data) as buff:
        if not is_zipfile(buff):
            raise ValueError(f'{description} file is not an XLSX ZIP file')
        with ZipFile(buff) as f:
            if '[Content_Types].xml' not in f.namelist() or 'xl/workbook.xml' not in f.namelist():
                raise ValueError(f'{description} is not a valid Excel workbook package')


def _template_dir(value: str | Path) -> Path:
    root = Path(value).expanduser().resolve(strict=True)
    if not root.is_dir():
        raise NotADirectoryError(f'Template folder not found: {root}')
    if any(not (root / conf['file']).is_file() for conf in BRANDS.values()):
        raise FileNotFoundError('Template folder must contain all five original ChannelWise workbooks')
    return root


def input_fingerprint(sales: bytes, returns: bytes, month: str, template_folder: str | Path) -> str:
    validate_month(month)
    h = sha256()
    for block in (sales, returns, month.encode(), str(Path(template_folder).resolve()).encode()):
        h.update(len(block).to_bytes(8, 'big'))
        h.update(block)
    return h.hexdigest()


def run_read_only_review(sales: bytes, returns: bytes, month: str, template_folder: str | Path) -> dict:
    """Read-only check: returns aggregate status; never writes back to input XLSX."""
    validate_month(month)
    _check_xlsx(sales, 'Sales')
    _check_xlsx(returns, 'Returns')
    root = _template_dir(template_folder)
    with TemporaryDirectory(prefix='channelwise_ui_review_') as temp:
        tmp = Path(temp)
        sales_file = tmp / 'sales_uploaded.xlsx'
        returns_file = tmp / 'returns_uploaded.xlsx'
        sales_file.write_bytes(sales)
        returns_file.write_bytes(returns)
        report = audit(sales_file, returns_file, root, month)
    report['ui_input_fingerprint'] = input_fingerprint(sales, returns, month, root)
    return report


def review_rows(report: dict) -> list[dict]:
    rows = []
    for business, value in report['businesses'].items():
        s, r = value['sales'], value['returns']
        rows.append({
            'Business': business,
            'Sales source': s['source_records'],
            'Sales missing': s['missing_from_template'],
            'Sales differences': s['full_field_different_matched_rows'],
            'Returns source': r['source_records'],
            'Returns missing': r['missing_from_template'],
            'Returns differences': r['full_field_different_matched_rows'],
            'Net amount': value['source_net_amount'],
        })
    return rows


def can_export_verified_copies(report: dict) -> bool:
    return (
        report.get('status') == 'NO_INSERTS_NEEDED_REVIEW_EXTRAS'
        and report.get('inputs_unchanged_sha256_verified') is True
        and report.get('financial_partition_check') is True
        and all(
            not values[kind]['missing_from_template']
            and not values[kind]['full_field_different_matched_rows']
            for values in report['businesses'].values()
            for kind in ('sales', 'returns')
        )
    )


def create_verified_copies_zip(
    sales: bytes, returns: bytes, month: str, template_folder: str | Path,
    previously_reviewed: dict,
) -> bytes:
    """Re-audits CURRENT inputs and exports identical snapshots only if fully matched.

    Does not insert transactions. Blocks future-month missing transactions.
    """
    if previously_reviewed.get('ui_input_fingerprint') != input_fingerprint(sales, returns, month, template_folder):
        raise ValueError('Inputs changed since review. Run validation again')
    report = run_read_only_review(sales, returns, month, template_folder)
    if not can_export_verified_copies(report):
        raise ValueError('Missing or conflicting transactions; cannot generate verified copies. Review required.')
    # Audit has been re-run so template changes between review and export are caught.
    if report['input_sha256'] != previously_reviewed.get('input_sha256'):
        raise ValueError('One or more source templates changed after review. Validate again')
    with TemporaryDirectory(prefix='channelwise_ui_export_') as temp:
        dest = Path(temp) / 'verified_copies'
        export_copies(_template_dir(template_folder), dest, month)
        blob = BytesIO()
        with ZipFile(blob, 'w', compression=ZIP_DEFLATED) as z:
            for item in sorted(dest.iterdir()):
                if item.is_file():
                    z.write(item, arcname=item.name)
            # Aggregate audit only; includes confidential hashes and business totals.
            z.writestr('PRIVATE_READ_ONLY_AUDIT.json', json.dumps(report, indent=2, default=str))
        return blob.getvalue()
