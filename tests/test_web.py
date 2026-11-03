"""웹 도구(index.html) 브라우저 테스트: Edge·Chrome에서 실제 버튼을 눌러 처리하고, 결과 파일을 열어 확인 (가짜 정보만).

실행:  pip install playwright python-docx openpyxl pypdf reportlab olefile pillow
       python -m unittest tests.test_web
브라우저: 환경 변수 PII_BROWSER=msedge(기본)|chrome
OCR(스캔본 PDF)까지 확인하려면 저장소 폴더를 http로 띄우고 그 주소를 PII_URL로 주세요.
       python -m http.server 8765 --bind 127.0.0.1   →   PII_URL=http://127.0.0.1:8765/index.html
"""
import io, os, random, re, sys, tempfile, unittest, zipfile, zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import pii_guard as g

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    sync_playwright = None

URL = os.environ.get('PII_URL') or ROOT.joinpath('index.html').as_uri()
BROWSER = os.environ.get('PII_BROWSER', 'msedge')
FONT = Path(r'C:\Windows\Fonts\malgun.ttf')
flat = lambda s: ''.join(s.split())   # OCR이 넣는 빈칸 무시


def cfb(streams):
    """규격대로 만든 작은 OLE 복합 파일(.hwp 테스트용): 512바이트 섹터, 4096바이트 미만은 미니 스트림."""
    END, FREE, FATS, NOS = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF
    u32 = lambda x: x.to_bytes(4, 'little')
    sec, fat = [], []

    def alloc(data, pad=b'\0'):
        n = max(1, -(-len(data) // 512)); s0 = len(sec)
        data = data.ljust(n * 512, pad)
        for i in range(n):
            sec.append(data[i * 512:(i + 1) * 512]); fat.append(s0 + i + 1 if i < n - 1 else END)
        return s0
    mini, minifat, start = b'', [], {}
    for k, v in streams.items():
        if len(v) >= 4096:
            start[k] = alloc(v)
        else:
            n = max(1, -(-len(v) // 64)); start[k] = len(minifat)
            minifat += [len(minifat) + i + 1 for i in range(n - 1)] + [END]; mini += v.ljust(n * 64, b'\0')
    mstart = alloc(mini) if mini else END
    mfstart = alloc(b''.join(map(u32, minifat)), b'\xff') if minifat else END
    ents = [['Root Entry', 5, mstart, len(mini), []]]   # 이름, 종류(5 루트·1 폴더·2 스트림), 시작, 크기, 자식
    for path in streams:
        cur, parts = 0, path.split('/')
        for i, name in enumerate(parts):
            hit = next((j for j in ents[cur][4] if ents[j][0] == name), None)
            if hit is None:
                leaf = i == len(parts) - 1
                ents.append([name, 2 if leaf else 1, start[path] if leaf else 0, len(streams[path]) if leaf else 0, []])
                hit = len(ents) - 1; ents[cur][4].append(hit)
            cur = hit
    right = {a: b for e in ents for a, b in zip(e[4], e[4][1:])}   # 형제는 right로만 이음
    d = b''
    for i, (name, typ, st, size, kids) in enumerate(ents):
        nm = name.encode('utf-16le') + b'\0\0'
        d += (nm.ljust(64, b'\0') + len(nm).to_bytes(2, 'little') + bytes([typ, 1]) + u32(NOS) + u32(right.get(i, NOS))
              + u32(kids[0] if kids else NOS) + bytes(36) + u32(st) + size.to_bytes(8, 'little'))
    dstart = alloc(d)
    nfat = -(-(len(sec) + 2) * 4 // 512)
    ids = list(range(len(sec), len(sec) + nfat)); fat += [FATS] * nfat
    hdr = (bytes.fromhex('D0CF11E0A1B11AE1') + bytes(16) + (0x3E).to_bytes(2, 'little') + (3).to_bytes(2, 'little')
           + (0xFFFE).to_bytes(2, 'little') + (9).to_bytes(2, 'little') + (6).to_bytes(2, 'little') + bytes(10) + u32(nfat)
           + u32(dstart) + bytes(4) + u32(4096) + u32(mfstart) + u32(1 if minifat else 0) + u32(END) + bytes(4)
           + b''.join(map(u32, ids)) + u32(FREE) * (109 - nfat))
    return hdr + b''.join(sec) + b''.join(map(u32, fat)).ljust(nfat * 512, b'\xff')


def hwp(lines, flags=1):
    """한글(.hwp) 5.0 모양: FileHeader + 압축한 BodyText/Section0 (문단 글자 레코드)."""
    rec = lambda tag, body: (tag | len(body) << 20).to_bytes(4, 'little') + body
    sec = b''.join(rec(66, bytes(22)) + rec(67, s.encode('utf-16le') + b'\x0d\x00') for s in lines)
    if flags & 1:
        c = zlib.compressobj(wbits=-15); sec = c.compress(sec) + c.flush()
    head = b'HWP Document File'.ljust(32, b'\0') + bytes([0, 0, 1, 5]) + flags.to_bytes(4, 'little') + bytes(216)
    return cfb({'FileHeader': head, 'BodyText/Section0': sec, 'PrvText': '<미리보기>'.encode('utf-16le')})


@unittest.skipIf(sync_playwright is None, 'playwright가 없어요 (pip install playwright)')
class Web(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pw = sync_playwright().start()
        cls.browser = cls.pw.chromium.launch(channel=BROWSER)
        cls.dir = tempfile.TemporaryDirectory()

    @classmethod
    def tearDownClass(cls):
        cls.browser.close(); cls.pw.stop(); cls.dir.cleanup()

    def setUp(self):
        self.page = self.browser.new_page(accept_downloads=True)
        self.errors, self.outside = [], []
        self.page.on('pageerror', lambda e: self.errors.append(str(e)))
        host = URL.split('/')[2] if URL.startswith('http') else None
        self.page.on('request', lambda r: self.outside.append(r.url) if r.url.startswith('http') and r.url.split('/')[2] != host else None)
        self.page.goto(URL)

    def tearDown(self):
        self.assertEqual(self.errors, [], 'JS 오류')
        self.assertEqual(self.outside, [], '외부(http/https) 요청')   # 개인정보는 물론 아무 데이터도 밖으로 나가면 안 됨
        self.page.close()

    def p(self, name):
        return os.path.join(self.dir.name, name)

    def text(self, src, mode='token', words=''):
        pg = self.page
        pg.fill('#src', src); pg.fill('#words', words); pg.select_option('#mode', mode); pg.click('#go')
        return pg.input_value('#out'), pg.inner_text('#mm')

    def files(self, paths, mode='token', pdf='txt', words='홍길동'):
        """파일 탭에서 처리하고 내려받은 결과: [(파일 경로 또는 None, 결과 문구)]"""
        pg = self.page
        pg.click('.tabs button[data-t="f"]')
        pg.evaluate("document.getElementById('fres').replaceChildren()")
        pg.set_input_files('#ff', paths); pg.fill('#fwords', words); pg.select_option('#fmode', mode); pg.select_option('#fpdf', pdf)
        pg.click('#fgo')
        pg.wait_for_function(f"!document.getElementById('fgo').disabled && document.querySelectorAll('#fres .res').length==={len(paths)}", timeout=300000)
        out = []
        for r in pg.query_selector_all('#fres .res'):
            info, btn = r.inner_text().replace('\n', ' / '), r.query_selector('button')
            if not btn:
                out.append((None, info)); continue
            with pg.expect_download() as dl:
                btn.click()
            path = self.p('web_' + dl.value.suggested_filename); dl.value.save_as(path); out.append((path, info))
        return out

    # ---------- 글 마스킹 ----------
    def test_text_mask_restore_partial(self):
        src = ('홍길동님(900101-1234567) 연락처 010-1234-5678, 메일 gildong.hong@example.com\n'
               '서버 IP는 192.168.0.15. 접수 2026-10-21, 버전 1.2.3.4.5\n'
               '환불 계좌 4428-010-108-0107 / 입금 계좌 44280101080107 / 주문번호 44280101080108\n'
               '주소: 서울시 송파구 위례송파로 123, 1706동 804호')
        out, msg = self.text(src, words='홍길동')
        for s in ('홍길동', '900101-1234567', '010-1234-5678', 'gildong', '192.168.0.15', '4428-010-108-0107', '44280101080107', '위례송파로 123'):
            self.assertNotIn(s, out)
        for s in ('2026-10-21', '1.2.3.4.5', '주문번호 44280101080108'):
            self.assertIn(s, out)
        self.assertIn('✅ 재검사', msg)
        pg = self.page
        pg.fill('#ans', out); pg.click('#rs')
        self.assertEqual(pg.input_value('#rout'), src)   # 이름표 → 원래 정보
        out, _ = self.text(src, 'partial', '홍길동')
        for s in ('홍*동', '900101-1******', '010-****-5678', '192.168.*.*.', '4428-***-***-0107', '442*******0107', '위례송파로 ***, ****동 ***호'):
            self.assertIn(s, out)

    def test_chip_off_and_formats_table(self):
        pg = self.page
        pg.click('#cats-t label:has(input[value="이메일"])')   # 이메일 끄기
        out, msg = self.text('메일 kim.yh@example.com 연락처 010-1234-5678')
        self.assertIn('kim.yh@example.com', out); self.assertNotIn('010-1234-5678', out); self.assertIn('ℹ️', msg)
        bad = pg.evaluate("""() => { const M = makeMasker('token', '', new Set(CATS)), bad = [];
            for (const kind of ['ok', 'no']) for (const c of document.querySelectorAll('#formats td.' + kind + ' code')) {
              const x = c.textContent; if ((kind === 'ok') !== (M.fn(x) !== x)) bad.push(kind + ': ' + x); } return bad; }""")
        self.assertEqual(bad, [], '화면의 「찾는 모양」 표가 실제 동작과 다름')

    def test_text_file_euckr(self):
        Path(self.p('a.csv')).write_bytes('이름,연락처\r\n홍길동,010-1234-5678\r\n'.encode('cp949'))
        self.page.set_input_files('#tf', self.p('a.csv'))
        self.page.wait_for_function("document.getElementById('src').value.length > 0")
        self.assertIn('홍길동', self.page.input_value('#src'))   # 글자가 깨지지 않음

    def test_web_and_python_agree(self):
        """같은 문장을 웹·파이썬에 넣으면 세 가지 방식 모두 글자 하나까지 같아야 함 (무작위 문장)."""
        rnd = random.Random(20261103)
        words = ['계좌', '입금', '은행', '주문번호', '연락처', '홍길동', '서울시', '경기도', '송파구', '구리시', '위례송파로', '갈매중앙로',
                 '대왕판교로', '번길', '동', '호', '층', 'account', 'kim.yh@example.com', 'M', '+82 ', '(', ')']
        def num():
            d = ''.join(rnd.choice('0123456789') for _ in range(rnd.choice([1, 2, 3, 4, 4, 6, 7, 10, 11, 13, 14, 16])))
            return ('01' + rnd.choice('016789') + d[3:]) if rnd.random() < .15 else d
        cases = list(dict.fromkeys(''.join((num() if rnd.random() < .5 else rnd.choice(words)) + rnd.choice(['-', '.', ' ', '', ',', '/', ' - '])
                                           for _ in range(rnd.randint(1, 8))) for _ in range(4000)))
        web = self.page.evaluate("""xs => xs.map(s => ['token', 'partial', 'star'].map(mode => {
            SES.seen = {}; SES.cnt = {}; SES.map = {}; return makeMasker(mode, '홍길동', new Set(CATS)).fn(s); }))""", cases)
        diff = [(s, mode, py, w) for s, ws in zip(cases, web) for mode, py, w in zip(
            ('token', 'partial', 'star'), (g.Masker(['홍길동']).text(s), g.Masker(['홍길동'], part=True).text(s), g.Masker(['홍길동'], star=True).text(s)), ws) if py != w]
        self.assertEqual(diff[:5], [])

    # ---------- 파일 마스킹 ----------
    def test_docx(self):
        import docx
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.oxml import parse_xml
        from docx.shared import Pt
        ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        d = docx.Document()
        p = d.add_paragraph(); p.add_run('중요:').bold = True; p.add_run(' 담당 010-1111-2222 확인 '); p.add_run('끝').italic = True
        p = d.add_paragraph('연락처: '); p.add_run('010-').bold = True; p.add_run('1234-'); p.add_run('5678').font.size = Pt(14)
        p = d.add_paragraph('메일: ')
        rid = p.part.relate_to('mailto:gildong.hong@example.com', RT.HYPERLINK, is_external=True)
        p._p.append(parse_xml(f'<w:hyperlink {ns} r:id="{rid}"><w:r><w:t>gildong.hong@example.com</w:t></w:r></w:hyperlink>'))
        p._p.append(parse_xml(f'<w:r {ns} xmlns:v="urn:schemas-microsoft-com:vml"><w:pict><v:shape><v:textbox><w:txbxContent>'
                              f'<w:p><w:r><w:t>상자 속 010-9876-5432</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>'))
        d.sections[0].header.paragraphs[0].text = '담당 02-345-6789'
        d.save(self.p('a.docx'))
        [(path, info)] = self.files([self.p('a.docx')])
        self.assertIn('✅ 재검사', info)
        with zipfile.ZipFile(path) as z:
            xml = '\n'.join(z.read(n).decode('utf-8', 'replace') for n in z.namelist() if n.endswith(('.xml', '.rels')))
            parts = sorted(n for n in z.namelist() if n.startswith(('word/header', 'word/footer')))
        for s in ('010-1111-2222', '010-9876-5432', 'gildong', '02-345-6789', '1234-'):
            self.assertNotIn(s, xml)
        self.assertEqual(parts, ['word/header1.xml'])   # 쓰지 않던 머리글·바닥글이 생기지 않음
        runs = [(r.text, r.bold, r.italic) for r in docx.Document(path).paragraphs[0].runs]
        self.assertEqual(runs, [('중요:', True, None), (' 담당 [휴대전화_1] 확인 ', None, None), ('끝', None, True)])   # 서식 유지

    def test_xlsx(self):
        import openpyxl
        from openpyxl.comments import Comment
        wb = openpyxl.Workbook(); ws = wb.active
        ws.append(['2026 환불 신청자 명단'])
        ws.append(['성명', '계좌', '주문번호', '금액', '연락처', '비고'])
        ws.append(['홍길동', 44280101080107, 44280101080108, 250000000, 1012345678, '홍길동 확인'])
        ws.append(['Kim Young Hee', '4428-010-108-0107', 'A-1', 3112345678, 501014123456, ''])
        ws.append(['합계', '', '', '', '', ''])
        ws['F4'].comment = Comment('예비 연락처 010-9876-5432', '작성자')
        ws['F4'].value = '메일'; ws['F4'].hyperlink = 'mailto:kim.yh@example.com'
        form = wb.create_sheet('양식'); form.append(['예금주', '김영희']); form.append(['입금계좌', '110123456789'])
        wb.save(self.p('a.xlsx'))
        with zipfile.ZipFile(self.p('a.xlsx'), 'a') as z:   # 최신 Excel의 스레드 메모
            z.writestr('xl/threadedComments/threadedComment1.xml', '<?xml version="1.0" encoding="UTF-8"?><ThreadedComments '
                       'xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="A1">'
                       '<text>담당자 lee.sj@example.com</text></threadedComment></ThreadedComments>')
        [(path, info)] = self.files([self.p('a.xlsx')])
        self.assertIn('✅ 재검사', info)
        w = openpyxl.load_workbook(path); s1, s2 = w.worksheets
        row3, row4 = [c.value for c in s1[3]], [c.value for c in s1[4]]
        self.assertTrue(str(row3[0]).startswith('[이름_') and str(row3[1]).startswith('[계좌번호_') and str(row4[0]).startswith('[이름_'))
        self.assertEqual((row3[2], row3[3], row4[3]), (44280101080108, 250000000, 3112345678))   # 주문번호·금액은 그대로
        self.assertTrue(row3[4].startswith('[휴대전화_') and row4[4].startswith('[주민등록번호_'))   # 앞자리 0이 지워진 번호
        self.assertTrue(str(row3[5]).endswith(' 확인') and '홍길동' not in str(row3[5]))   # 찾은 이름은 다른 칸에서도
        self.assertEqual((s1['A2'].value, s1['A5'].value), ('성명', '합계'))
        self.assertTrue(str(s2['B1'].value).startswith('[이름_') and str(s2['B2'].value).startswith('[계좌번호_'))
        with zipfile.ZipFile(path) as z:
            xml = '\n'.join(z.read(n).decode('utf-8', 'replace') for n in z.namelist() if n.endswith(('.xml', '.rels')))
        for s in ('010-9876-5432', 'kim.yh', 'lee.sj', '홍길동', '44280101080107'):
            self.assertNotIn(s, xml)

    def test_text_files(self):
        Path(self.p('a.csv')).write_bytes('이름,연락처\r\n홍길동,010-1234-5678\r\n'.encode('cp949'))
        Path(self.p('b.txt')).write_bytes('연락처 010-1234-5678\r\n'.encode('utf-16'))
        (a, ia), (b, ib) = self.files([self.p('a.csv'), self.p('b.txt')])
        raw = Path(a).read_bytes()
        self.assertTrue(raw.startswith(b'\xef\xbb\xbf'))   # 엑셀용 BOM
        self.assertEqual(raw[3:].decode(), '이름,연락처\r\n[지정어_1],[휴대전화_1]\r\n')
        self.assertNotIn('010-1234-5678', Path(b).read_text(encoding='utf-8'))

    @unittest.skipUnless(FONT.exists(), '맑은 고딕 글꼴이 없어요')
    def test_pdf_masked_pdf_and_txt(self):
        from pypdf import PdfReader
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.pdfgen import canvas
        pdfmetrics.registerFont(TTFont('Malgun', str(FONT)))
        cv = canvas.Canvas(self.p('a.pdf')); cv.setFont('Malgun', 13)
        for i, s in enumerate(['성명: 홍길동   주민등록번호: 901023-1234567', '휴대전화: 010-9249-5149 (평일)', '비고: 2026-10-21 접수']):
            cv.drawString(60, 760 - i * 28, s)
        cv.save()
        [(pdf, info)] = self.files([self.p('a.pdf')], 'partial', 'pdf')
        r = PdfReader(pdf)
        self.assertTrue(pdf.endswith('_masked.pdf') and '✅ 재검사' in info)
        self.assertEqual(''.join(x.extract_text() or '' for x in r.pages).strip(), '')   # 글자층 없음 → 복사해도 안 나옴
        [(txt, _)] = self.files([self.p('a.pdf')], 'partial', 'txt')
        t = Path(txt).read_text(encoding='utf-8')
        for s in ('홍*동', '901023-1******', '010-****-5149', '2026-10-21'):
            self.assertIn(s, t)

    def test_hwpx_and_hwp(self):
        hp = 'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"'
        sec = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec {hp} xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section">'
               '<hp:p><hp:run charPrIDRef="1"><hp:t>담당 010-</hp:t></hp:run><hp:run charPrIDRef="2"><hp:t>1234-5678 확인</hp:t></hp:run></hp:p>'
               '<hp:p><hp:run><hp:t>메일<hp:tab width="4000"/>kim.yh@example.com</hp:t></hp:run></hp:p>'
               '<hp:p><hp:run><hp:tbl><hp:tr><hp:tc><hp:subList><hp:p><hp:run><hp:t>900101-1234567</hp:t></hp:run></hp:p></hp:subList></hp:tc></hp:tr></hp:tbl></hp:run></hp:p>'
               '<hp:p><hp:run><hp:ctrl><hp:fieldBegin type="HYPERLINK"><hp:parameters><hp:stringParam name="Command">mailto\\:lee.sj@example.com;1;0;0;</hp:stringParam>'
               '</hp:parameters></hp:fieldBegin></hp:ctrl><hp:t>링크</hp:t></hp:run></hp:p></hs:sec>')
        with zipfile.ZipFile(self.p('a.hwpx'), 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr(zipfile.ZipInfo('mimetype'), 'application/hwp+zip')
            z.writestr('Contents/section0.xml', sec)
            z.writestr('Contents/content.hpf', '<?xml version="1.0" encoding="UTF-8"?><opf:package xmlns:opf="http://www.idpf.org/2007/opf/"><opf:metadata>'
                       '<opf:meta name="lastsaveby" content="text">홍길동</opf:meta></opf:metadata></opf:package>')
            z.writestr('Preview/PrvText.txt', '담당 010-1234-5678\r\n')
            z.writestr('Preview/PrvImage.png', b'\x89PNG first page')
        Path(self.p('b.hwp')).write_bytes(hwp(['성명: 홍길동', '주민등록번호: 900101-1234567', '이메일: gildong.hong@example.com']))
        Path(self.p('big.hwp')).write_bytes(hwp([f'{i}번째 연락처 010-{1000 + i}-5678' for i in range(200)], flags=0))   # 압축 안 함, 일반 섹터
        Path(self.p('locked.hwp')).write_bytes(hwp(['x'], flags=3))   # 비밀번호 문서
        (x, ix), (h, ih), (big, _), (lk, il) = self.files([self.p(n) for n in ('a.hwpx', 'b.hwp', 'big.hwp', 'locked.hwp')])
        with zipfile.ZipFile(self.p('a.hwpx')) as zi, zipfile.ZipFile(x) as zo:
            self.assertEqual(zi.namelist(), zo.namelist())   # 파일 목록·순서 그대로
            self.assertEqual(zo.infolist()[0].compress_type, zipfile.ZIP_STORED)   # mimetype 무압축 맨 앞
            xml = zo.read('Contents/section0.xml').decode()
            self.assertEqual(zo.read('Preview/PrvImage.png'), g.BLANK_PNG)
            rest = xml + zo.read('Contents/content.hpf').decode() + zo.read('Preview/PrvText.txt').decode()
        self.assertIn('charPrIDRef="2"><hp:t> 확인</hp:t>', xml)   # 나뉜 run 서식 유지
        for s in ('010-1234', '900101-1234567', 'kim.yh', 'lee.sj', '홍길동'):
            self.assertNotIn(s, rest)
        self.assertRegex(Path(h).read_text(encoding='utf-8'), r'^성명: \[지정어_\d+\]\n주민등록번호: \[주민등록번호_\d+\]\n이메일: \[이메일_\d+\]\n$')
        self.assertEqual(Path(big).read_text(encoding='utf-8').count('[휴대전화_'), 200)
        self.assertIsNone(lk); self.assertIn('비밀번호', il)
        for name in ('b.hwp', 'big.hwp'):   # 웹과 파이썬의 .hwp 글자 추출이 같아야 함
            import olefile
            data = Path(self.p(name)).read_bytes()
            web = self.page.evaluate('a => hwpText(new Uint8Array(a).buffer)', list(data))
            with olefile.OleFileIO(self.p(name)) as f:
                self.assertEqual(web, g.hwp_text(lambda n: f.openstream(n).read() if f.exists(n) else None))

    @unittest.skipUnless(URL.startswith('http') and FONT.exists(), 'OCR은 index.html을 http로 띄워야 해요 (PII_URL)')
    def test_ocr_scanned_pdf(self):
        from PIL import Image, ImageDraw, ImageFilter, ImageFont
        font = ImageFont.truetype(str(FONT), 40)
        def page(lines, blur):
            im = Image.new('L', (1654, 2339), 255); dr = ImageDraw.Draw(im)
            for i, s in enumerate(lines):
                dr.text((150, 200 + i * 90), s, fill=0, font=font)
            im = im.rotate(0.4, fillcolor=255)   # 스캐너처럼 살짝 기울임
            return im.filter(ImageFilter.GaussianBlur(blur)) if blur else im
        lines = ['개인정보 수집 동의서', '성명: 홍길동', '주민등록번호: 900101-1234567', '휴대전화: 010-9249-5149',
                 '이메일: gildong.hong@example.com', '주소: 경기도 구리시 갈매중앙로 16, 709-702']
        page(lines, 0).save(self.p('scan.pdf'), 'PDF', resolution=200, save_all=True,
                            append_images=[page(['보호자 연락처 02-2260-3912', '계좌번호 110-123-456789'], .8)])
        secrets = ['홍길동', '900101-1234567', '010-9249-5149', 'gildong.hong@example.com', '갈매중앙로16', '02-2260-3912', '110-123-456789']
        (txt, info), = self.files([self.p('scan.pdf')], 'token', 'txt')
        self.assertIn('2쪽은 OCR로 읽음', info)
        self.assertEqual([s for s in secrets if flat(s) in flat(Path(txt).read_text(encoding='utf-8'))], [])
        (pdf, _), = self.files([self.p('scan.pdf')], 'token', 'pdf')
        (again, _), = self.files([pdf], 'token', 'txt')   # 가린 PDF를 다시 OCR로 읽어도 개인정보가 안 보여야 함
        self.assertEqual([s for s in secrets if flat(s) in flat(Path(again).read_text(encoding='utf-8'))], [])


if __name__ == '__main__':
    unittest.main()
