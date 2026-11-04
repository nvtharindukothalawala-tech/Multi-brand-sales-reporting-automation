"""Create private, checksum-verified copies of existing ChannelWise reports.

This is a snapshot/export step, NOT a workbook update or recalculation.
No originals are changed. Copies only; real audit remains a separate required check.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import shutil
from datetime import datetime, timezone
from pathlib import Path

FILES = (
    'BeDapper_ChannelWise_Report.xlsx',
    'BeDesi_ChannelWise_Report.xlsx',
    'Giggles_ChannelWise_Report.xlsx',
    'Protecto_ChannelWise_Report.xlsx',
    'SafetySure_ChannelWise_Report.xlsx',
)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as src:
        for part in iter(lambda: src.read(1024 * 1024), b''):
            h.update(part)
    return h.hexdigest()


def export_copies(source_dir: Path, dest_dir: Path, month: str) -> None:
    source_dir = source_dir.resolve(strict=True)
    dest_dir = dest_dir.resolve()
    try:
        datetime.strptime(month, '%Y-%m')
    except ValueError as exc:
        raise ValueError('Month must be YYYY-MM') from exc
    if len(month) != 7 or month[4] != '-':
        raise ValueError('Month must be YYYY-MM')
    if dest_dir.exists():
        raise FileExistsError(f'Output already exists; use a fresh directory: {dest_dir}')
    sources = [source_dir / file for file in FILES]
    for src in sources:
        if not src.is_file():
            raise FileNotFoundError(f'Missing report: {src}')
    source_hashes = {src.name: sha256(src) for src in sources}
    if any(src.resolve() == (dest_dir / src.name).resolve() for src in sources):
        raise ValueError('Destination must differ from all source files')
    dest_dir.mkdir(parents=True, exist_ok=False)
    rows = []
    for src in sources:
        dst = dest_dir / src.name
        with src.open('rb') as reader, dst.open('xb') as writer:
            shutil.copyfileobj(reader, writer, 1024 * 1024)
        copied_hash = sha256(dst)
        if copied_hash != source_hashes[src.name]:
            raise RuntimeError(f'Copy integrity failed: {src.name}')
        rows.append((src.name, source_hashes[src.name], copied_hash, 'IDENTICAL_COPY'))
    for src in sources:
        if sha256(src) != source_hashes[src.name]:
            raise RuntimeError(f'Source file changed during export: {src.name}')
    with (dest_dir / 'COPY_VERIFICATION.csv').open('x', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['filename', 'source_sha256', 'copy_sha256', 'result'])
        writer.writerows(rows)
    (dest_dir / 'READ_ME_FIRST.txt').write_text(
        f'PRIVATE CHANNELWISE SNAPSHOT | {month}\n'
        'These are byte-identical copies of the existing company Excel files.\n'
        'No transactions were added; no formulas or pivots recalculated.\n'
        'Before accepting an October report, run the separate real-data read-only audit.\n'
        'KEEP PRIVATE. DO NOT UPLOAD XLSX FILES TO GITHUB.\n',
        encoding='utf-8',
    )
    print(f'EXPORT COMPLETED: {len(rows)} copied workbooks')
    print(f'Location: {dest_dir}')
    print('Original files unchanged: SHA-256 verified')
    print('No transaction updates performed')


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--templates-dir', type=Path, required=True)
    parser.add_argument('--output-dir', type=Path, required=True)
    parser.add_argument('--month', default='2026-10')
    args = parser.parse_args()
    export_copies(args.templates_dir, args.output_dir, args.month)


if __name__ == '__main__':
    main()
