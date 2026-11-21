"""Fixed synthetic Excel-export check used by the source and frozen build tests."""
import json
from pathlib import Path


def check_excel_exports(application, destination):
    directory = Path(destination).absolute()
    try:
        directory.mkdir(parents=True, exist_ok=False)
    except OSError:
        return 2  # Never use an existing directory or replace user files.
    result = dict(status='failed', excel_available=application.OPENPYXL_AVAILABLE,
                  numpy_fallback=application.OPENPYXL_NUMPY_FALLBACK, reports=[])
    try:
        if not application.OPENPYXL_AVAILABLE:
            raise application._openpyxl_unavailable_error('Excel support unavailable')
        records = [dict(si=1, name='synthetic.txt', size_bytes=1536, path='synthetic.txt'),
                   dict(si=2, name='=1+1.txt', size_bytes=2097152, path='=1+1.txt')]
        folders = [dict(folder='SYNTHETIC', files=records, total_size=2098688)]
        expected = {'auto': ['1.50 KB', '2.00 MB'], 'kb': ['1.50 KB', '2048.00 KB'],
                    'mb': ['0.00 MB', '2.00 MB']}
        for unit, sizes in expected.items():
            workbook_path = directory / f'inventory-{unit}.xlsx'
            html_path = directory / f'inventory-{unit}.html'
            for generator, path in ((application.ExcelReportGenerator, workbook_path),
                                    (application.HtmlReportGenerator, html_path)):
                generator.generate(path, folders, ['si', 'name', 'size'], 'Synthetic Excel Check',
                                   2, 2098688, size_unit=unit)
            workbook = application.openpyxl.load_workbook(workbook_path)
            try:
                rows = [row for row in workbook.active.iter_rows() if isinstance(row[0].value, int)]
                assert [row[2].value for row in rows] == sizes
                assert rows[1][1].value == '=1+1.txt' and rows[1][1].data_type == 's'
            finally:
                workbook.close()
            rendered = html_path.read_text(encoding='utf-8')
            assert all(size in rendered for size in sizes)
            result['reports'].extend((workbook_path.name, html_path.name))
        hashes_path = directory / 'hash-verification.xlsx'
        application.HashReportExporter.export_excel(hashes_path, [dict(
            name='synthetic.txt', path='synthetic.txt', size=1536, algorithm='SHA-256', result='CALCULATED',
            calculated_hash='a' * 64)])
        workbook = application.openpyxl.load_workbook(hashes_path)
        try:
            assert any(cell.value == 'synthetic.txt' for row in workbook.active for cell in row)
        finally:
            workbook.close()
        result['reports'].append(hashes_path.name)
        result['status'] = 'passed'
    except Exception as error:
        result['error'] = f'{type(error).__name__}: {error}'
    with (directory / 'result.json').open('x', encoding='utf-8') as output:
        json.dump(result, output, indent=2)
    return 0 if result['status'] == 'passed' else 1
