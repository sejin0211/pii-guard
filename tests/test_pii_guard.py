"""pii_guard.py 테스트 (가짜 정보만 사용).

실행:  python -m unittest discover tests
필요:  pip install python-docx openpyxl pypdf
"""
import os, re, sys, tempfile, unittest, zipfile
from pathlib import Path

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import pii_guard as g

# (입력, 기대 결과)
CASES = [
    ('900101-1234567', '[주민등록번호_1]'),
    ('900101-5234567', '[주민등록번호_1]'),            # 외국인등록번호
    ('1234-5678-9012-3456', '[카드번호_1]'),
    ('010-1234-5678', '[휴대전화_1]'),
    ('02-2260-3912', '[전화번호_1]'),
    ('(02)2260-3912', '[전화번호_1]'),
    ('07012345678', '[전화번호_1]'),
    ('+82 10 1234 5678', '[전화번호_1]'),
    ('gildong.hong@example.com', '[이메일_1]'),
    ('123-45-67890', '[사업자번호_1]'),
    ('M12345678', '[여권번호_1]'),
    ('IP는 192.168.0.15.', 'IP는 [IP주소_1].'),
    ('110-123-456789', '[계좌번호_1]'),
    ('2026-09-29', '2026-09-29'),                       # 날짜는 그대로
    ('버전 1.2.3.4.5', '버전 1.2.3.4.5'),                 # IP 아님
]


class Rules(unittest.TestCase):
    def test_cases(self):
        for src, want in CASES:
            with self.subTest(src=src):
                self.assertEqual(g.Masker().text(src), want)

    def test_same_value_same_token_and_words(self):
        M = g.Masker(['홍길동'])
        self.assertEqual(M.text('홍길동 010-1234-5678 / 홍길동 010-1234-5678'),
                         '[지정어_1] [휴대전화_1] / [지정어_1] [휴대전화_1]')
        self.assertEqual(M.map['[휴대전화_1]'], '010-1234-5678')

    def test_rules_match_web(self):
        """웹(index.html)과 파이썬의 규칙이 똑같은지 확인 → 한쪽만 고치는 실수 방지."""
        page = Path(ROOT, 'index.html').read_text(encoding='utf-8')
        block = page[page.index('const RULES=['):]
        block = block[:block.index('];')]
        web = dict(re.findall(r"\['([^']+)',/(.+?)/g\]", block))
        self.assertEqual(web, {label: rx.pattern for label, rx in g.RULES})


class Files(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.p = lambda name: os.path.join(self.dir.name, name)

    def tearDown(self):
        self.dir.cleanup()

    def test_docx_keeps_format_and_masks_hidden_text(self):
        import docx
        from docx.oxml import parse_xml
        ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        d = docx.Document()
        p = d.add_paragraph()
        p.add_run('중요:').bold = True
        p.add_run(' 담당 010-1111-2222 확인 ')
        p.add_run('끝').italic = True
        p = d.add_paragraph('연락처 ')   # 변경 내용 추적으로 지운 번호, 여러 조각으로 나뉜 필드 코드
        p._p.append(parse_xml(f'<w:del {ns} w:id="1" w:author="a"><w:r><w:delText>010-9876-5432</w:delText></w:r></w:del>'))
        for piece in (' HYPERLINK "mailto:kim.yh', '@example.com" '):
            p._p.append(parse_xml(f'<w:r {ns}><w:instrText xml:space="preserve">{piece}</w:instrText></w:r>'))
        d.save(self.p('a.docx'))
        g.do_docx(self.p('a.docx'), self.p('a_masked.docx'), g.Masker())

        runs = [(r.text, r.bold, r.italic) for r in docx.Document(self.p('a_masked.docx')).paragraphs[0].runs]
        self.assertEqual(runs, [('중요:', True, None), (' 담당 [휴대전화_1] 확인 ', None, None), ('끝', None, True)])
        with zipfile.ZipFile(self.p('a_masked.docx')) as z:
            xml = z.read('word/document.xml').decode()
        for s in ('010-1111-2222', '010-9876-5432', 'kim.yh'):
            self.assertNotIn(s, xml)

    def test_xlsx_leading_zero_and_recheck(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 1012345678                      # 서식으로 010-1234-5678처럼 보이는 셀
        ws['A1'].number_format = '000-0000-0000'
        ws['B1'] = 501014123456                    # 0501014123456 (앞자리 0이 지워진 주민번호)
        ws.title = '010-1234-5678'                 # 시트 이름은 가리지 않음 → 재검사가 알려야 함
        wb.save(self.p('a.xlsx'))
        g.do_xlsx(self.p('a.xlsx'), self.p('a_masked.xlsx'), g.Masker())

        ws = openpyxl.load_workbook(self.p('a_masked.xlsx')).active
        self.assertEqual((ws['A1'].value, ws['B1'].value), ('[휴대전화_1]', '[주민등록번호_1]'))
        self.assertEqual(g.leftover(self.p('a_masked.xlsx'), []), {'휴대전화': 1})

    def test_text_encodings(self):
        for enc in ('cp949', 'utf-16'):
            with self.subTest(enc=enc):
                with open(self.p('a.csv'), 'wb') as f:
                    f.write('이름,연락처\r\n홍길동,010-1234-5678\r\n'.encode(enc))
                g.do_text(self.p('a.csv'), self.p('a_masked.csv'), g.Masker(['홍길동']))
                raw = Path(self.p('a_masked.csv')).read_bytes()
                self.assertTrue(raw.startswith(b'\xef\xbb\xbf'))   # 엑셀용 BOM
                self.assertEqual(raw[3:].decode('utf-8'), '이름,연락처\r\n[지정어_1],[휴대전화_1]\r\n')


if __name__ == '__main__':
    unittest.main()
