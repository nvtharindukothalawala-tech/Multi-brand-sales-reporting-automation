"""Synthetic/isolated tests: no private business data used."""
from io import BytesIO
from zipfile import ZipFile
import pytest
from src import reporting_ui_service as svc


def test_month_validation():
    assert svc.validate_month('2026-10') == '2026-10'
    for bad in ('2026-13', '2026-00', '2026-1', 'any'):
        with pytest.raises(ValueError):
            svc.validate_month(bad)


def test_fingerprint_is_bound_to_both_files_month_and_folder():
    f = svc.input_fingerprint(b'a', b'b', '2026-10', 'first')
    assert f != svc.input_fingerprint(b'A', b'b', '2026-10', 'first')
    assert f != svc.input_fingerprint(b'a', b'b', '2026-11', 'first')
    assert f != svc.input_fingerprint(b'a', b'b', '2026-10', 'second')


def test_validate_xlsx_rejects_invalid_bytes():
    with pytest.raises(ValueError):
        svc._check_xlsx(b'not excel', 'Sales')


def test_validate_xlsx_accepts_workbook():
    from openpyxl import Workbook
    book=Workbook()
    out=BytesIO()
    book.save(out)
    svc._check_xlsx(out.getvalue(), 'Sales')


def sample_report(missing=0, drift=0):
    return {'status': 'NO_INSERTS_NEEDED_REVIEW_EXTRAS',
        'inputs_unchanged_sha256_verified': True, 'financial_partition_check': True,
        'businesses': {'Sample': {'sales': {'missing_from_template':missing,'full_field_different_matched_rows':drift},
                                  'returns': {'missing_from_template':0,'full_field_different_matched_rows':0}}}}


def test_no_new_records_allows_verified_snapshots_only():
    assert svc.can_export_verified_copies(sample_report())


def test_missing_records_blocks_export():
    assert not svc.can_export_verified_copies(sample_report(missing=1))


def test_drift_blocks_export():
    assert not svc.can_export_verified_copies(sample_report(drift=1))


def test_review_required_blocks_export():
    report=sample_report()
    report['status']='REVIEW_REQUIRED'
    assert not svc.can_export_verified_copies(report)


def test_report_unverified_hashes_blocks_export():
    report=sample_report()
    report['inputs_unchanged_sha256_verified']=False
    assert not svc.can_export_verified_copies(report)


def test_export_rejects_changed_user_input_without_touching_workbooks(tmp_path):
    with pytest.raises(ValueError,match='Inputs changed'):
        svc.create_verified_copies_zip(b'new',b'returns','2026-10',tmp_path,{'ui_input_fingerprint':'old'})


def test_uploaded_file_name_never_used_as_disk_path():
    import inspect
    code=inspect.getsource(svc.run_read_only_review)
    assert "tmp / 'sales_uploaded.xlsx'" in code
    assert "tmp / 'returns_uploaded.xlsx'" in code
