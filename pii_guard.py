#!/usr/bin/env python3
"""AI에 넘기기 전 개인정보 마스킹 (Word / Excel / PDF / 텍스트) + 선택적 비밀번호 설정.

설치:  pip install python-docx openpyxl pypdf cryptography msoffcrypto-tool
사용:  python pii_guard.py 파일1 [파일2 ...] [--words 홍길동,김영희] [--star | --partial] [--map] [--lock | --lock-only]

  --words    자동으로 못 찾는 이름·회사명 등 (쉼표 구분)
  --star     토큰 대신 별표(*)로 가림 (복원 불가)
  --partial  일부만 가림 (예: 010-****-5678, 901023-1******) — 업무에 다시 쓰기 좋음, 복원 불가
  --map      토큰→원본 대응표(_map.json) 저장. 이 파일에는 개인정보가 들어 있으니 따로 보관하세요.
  --lock       원본 파일(개인정보 포함)에 열기 비밀번호를 걸어 '이름_locked.확장자'로 함께 저장 (docx, xlsx, pdf)
  --lock-only  마스킹은 하지 않고 비밀번호만 설정
결과: 원본은 건드리지 않고 '이름_masked.확장자' 로 저장합니다. PDF는 '이름_pdf_masked.txt'(텍스트만 추출).
"""
import argparse, getpass, html, json, os, re, sys, zipfile
from urllib.parse import unquote

N, X = r'(?<!\d)', r'(?!\d)'
C = lambda p: re.compile(p, re.ASCII)   # \d는 0-9만 (웹의 자바스크립트와 같게). 전각 숫자는 FW로 따로 처리
RULES = [   # index.html의 RULES와 정규식이 같아야 함 (tests/test_pii_guard.py가 확인)
    # 앞 6자리는 실제 날짜(월 01~12, 일 01~31)일 때만, 뒷자리 5~8은 외국인등록번호
    ('주민등록번호', C(N + r'\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])[- ]?[1-8]\d{6}' + X)),
    ('카드번호',     C(N + r'\d{4}[- ]?\d{4}[- ]?\d{4}[- ]?\d{4}' + X)),
    ('휴대전화',     C(N + r'(?:\(01[016789]\)|01[016789])[-. ]?\d{3,4}[-. ]?\d{4}' + X)),
    # (02)123-4567, 070·080, +82 2-123-4567 / +82 (0)2-123-4567 / +82 10-1234-5678 형식 포함
    ('전화번호',     C(r'(?:(?<!\d)(?:\(0(?:2|[3-6]\d|70|80)\)|0(?:2|[3-6]\d|70|80))|\+82[-. ]?(?:\(0\)[-. ]?)?(?:2|[3-6]\d|70|80|1[016789]))'
                       r'[-. ]?\d{3,4}[-. ]?\d{4}' + X)),
    ('이메일',       C(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+')),
    ('사업자번호',   C(N + r'\d{3}-\d{2}-\d{5}' + X)),
    ('여권번호',     C(r'(?<![A-Za-z0-9])[MSROD](?:\d{8}|\d{3}[A-Z]\d{4})(?![A-Za-z0-9])')),  # 신형(2021~): M123A4567
    ('IP주소',       C(r'(?<![\d.])(?:\d{1,3}\.){3}\d{1,3}(?!\d|\.\d)')),  # 문장 끝 마침표는 허용
    ('계좌번호',     C(N + r'\d{2,6}-\d{2,6}-\d{2,8}(?:-\d{1,4})?' + X)),
    # 도로명·지번 주소. 시·도와 시·군·구가 둘 다 있으면 그대로, 하나만 있으면 동·층·호(709-702 포함)까지 있어야 함 ("경기 결과로 3배" 같은 오탐 방지)
    ('주소',         C(r'(?<![가-힣])(?:(?:서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충청북|충청남|충북|충남|전라북|전라남|전북|전남|경상북|경상남|경북|경남|제주)'
                       r'(?:특별시|광역시|특별자치시|특별자치도|도|시)?\s+(?:[가-힣]+(?:시|군|구)\s+){1,2}(?:[가-힣]+(?:읍|면)\s+)?[가-힣0-9]+(?:로|'
                       r'길|동|리|가)\s*\d+(?:-\d+)?(?:번지)?(?!\d)(?:,?\s*[0-9A-Za-z]+동(?:\s*\d+층)?(?:\s*[0-9A-Za-z-]+호)'
                       r'?|,?\s*\d+층(?:\s*[0-9A-Za-z-]+호)?|,?\s*[0-9A-Za-z-]+호|,?\s*\d{1,4}-\d{1,4}(?![\d-]))?|(?:(?:서울|'
                       r'부산|대구|인천|광주|대전|울산|세종|경기|강원|충청북|충청남|충북|충남|전라북|전라남|전북|전남|경상북|경상남|경북|경남|제주)(?:특별시|광역시|특별자치시|특별자치도|'
                       r'도|시)?\s+|(?:[가-힣]+(?:시|군|구)\s+){1,2})(?:[가-힣]+(?:읍|면)\s+)?[가-힣0-9]+(?:로|길|동|리|가)\s*\d+(?:-\d+)'
                       r'?(?:번지)?(?!\d)(?:,?\s*[0-9A-Za-z]+동(?:\s*\d+층)?(?:\s*[0-9A-Za-z-]+호)?|,?\s*\d+층(?:\s*[0-9A-Za-z-]+호)'
                       r'?|,?\s*[0-9A-Za-z-]+호|,?\s*\d{1,4}-\d{1,4}(?![\d-])))')),
]
DATE = C(r'(19|20)\d{2}-\d{2}-\d{2}')  # 날짜를 계좌번호로 오인하지 않도록
# 숫자 셀 앞에 0을 붙여 볼 때는 이 둘만 (250000000원 같은 금액이 02-5000-0000 전화번호로 가려지지 않게).
# shortcut: 10억대 금액 중 10·11·16~19로 시작하는 값은 여전히 휴대전화로 보임, 셀 표시 형식까지 볼 필요가 생기면 그때 바꿈
ZERO = [rx for label, rx in RULES if label in ('휴대전화', '주민등록번호')]
FW = str.maketrans('０１２３４５６７８９', '0123456789')   # 전각 숫자 → 일반 숫자 (글자 수가 같아 위치가 그대로)
PH_OPEN, PH_CLOSE = chr(0xE000), chr(0xE001)   # 찾은 값을 잠시 바꿔 두는 보호 문자 (사용자 영역 문자)
PH = re.compile(PH_OPEN + '([' + chr(0xE010) + '-' + chr(0xE019) + ']+)' + PH_CLOSE)


def partial(label, m):
    """일부만 가리기. 웹(index.html)의 partial()과 같은 결과 (예: 010-****-5678, 901023-1******)."""
    m = m.translate(FW)
    if label == '지정어':
        return m[0] + '*' * (len(m) - 1) if len(m) <= 2 else m[0] + '*' * (len(m) - 2) + m[-1]
    if label == '이메일':
        i = m.index('@')
        k = 2 if i > 2 else 1
        return m[:k] + '*' * max(1, i - k) + m[i:]
    if label == 'IP주소':
        p = m.split('.')
        return f'{p[0]}.{p[1]}.*.*'
    if label == '여권번호':
        return m[0] + '****' + m[-4:]
    if label == '주소':   # 지역·도로명은 남기고 번지·동·호수만 가림
        return re.sub(r'[0-9]', '*', m)
    if label == '주민등록번호':
        d, sep = re.sub(r'[^0-9]', '', m), re.search(r'[- ]', m)
        return d[:6] + (sep.group(0) if sep else '') + d[6] + '******'
    n = len(re.split(r'[-. ]', m))
    if n >= 3:   # 첫 묶음과 마지막 묶음(최대 4자리)만 남김
        k = -1

        def grp(x):
            nonlocal k
            k, s = k + 1, x.group(0)
            if k == 0:
                return s
            if k == n - 1:
                return '*' * (len(s) - 4) + s[-4:] if len(s) > 4 else s
            return '*' * len(s)
        return re.sub(r'[0-9]+', grp, m)
    first, ds = 4 if label == '카드번호' else 3, [i for i, c in enumerate(m) if c in '0123456789']
    hide = set(ds[first:len(ds) - 4])
    return ''.join('*' if i in hide else c for i, c in enumerate(m))


class Masker:
    def __init__(self, words=(), star=False, part=False):
        self.star, self.part, self.map, self.seen, self.cnt, self.found = star, part, {}, {}, {}, {}
        self.rules = list(RULES)
        ws = sorted({w.strip().translate(FW) for w in words if w.strip()}, key=len, reverse=True)
        if ws:
            self.rules.insert(0, ('지정어', re.compile('|'.join(map(re.escape, ws)))))

    def text(self, s):
        return ''.join(p[-1] for p in self.parts(s))

    def parts(self, s):
        """원문을 [그대로인 조각] 또는 [찾은 값, 바꿀 값]의 목록으로 나눔 (Word 서식 유지에 사용).
        찾은 값은 잠시 보호 문자로 바꿔 두어 뒤의 규칙이 다시 건드리지 않게 함 (웹 도구와 같은 방식).
        검사는 전각 숫자를 일반 숫자로 바꾼 글(s)에서 하고, 결과에는 원래 글자(o)를 씀."""
        keep = []
        o, s = s, s.translate(FW)
        for label, rx in self.rules:
            ns, no, last = [], [], 0
            for m in rx.finditer(s):
                if label == '계좌번호' and DATE.fullmatch(m.group(0)):   # 날짜는 그대로
                    continue
                v = o[m.start():m.end()]
                keep.append([v, self._sub(label, v)])
                ph = PH_OPEN + ''.join(chr(0xE010 + int(d)) for d in str(len(keep) - 1)) + PH_CLOSE
                ns += [s[last:m.start()], ph]
                no += [o[last:m.start()], ph]
                last = m.end()
            s, o = ''.join(ns) + s[last:], ''.join(no) + o[last:]
        out = []
        for i, seg in enumerate(PH.split(o)):
            if i % 2:
                out.append(keep[int(''.join(str(ord(c) - 0xE010) for c in seg))])
            elif seg:
                out.append([seg])
        return out

    def _sub(self, label, m):
        self.found[label] = self.found.get(label, 0) + 1
        if self.part:
            return partial(label, m)
        if self.star:
            return re.sub(r'[0-9０-９A-Za-z가-힣]', '*', m)
        key = (label, m)
        if key not in self.seen:
            n = self.cnt[label] = self.cnt.get(label, 0) + 1
            self.seen[key] = f'[{label}_{n}]'
            self.map[self.seen[key]] = m
        return self.seen[key]


W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
DOCX_PARTS = re.compile(r'/word/(document|header\d*|footer\d*|footnotes|endnotes|comments|people)\.xml')   # people: 댓글·변경 추적 작성자 목록
ATTRS = ('instr', 'author', 'userId')   # 속성 속 글자: 필드 코드, 작성자 이름, 작성자 계정(S::이메일::…)


def mask_group(ts, M):
    """번호가 여러 조각(run)으로 나뉘어 있어도 잡도록 한 문단의 글자를 이어 붙여 검사하고,
    바꿀 값은 그 값이 시작된 조각에 넣어 나머지 글자의 서식(굵게·크기 등)은 그대로 둠."""
    o = [t.text or '' for t in ts]
    parts = M.parts(''.join(o))
    if all(len(p) == 1 for p in parts):
        return
    owner = [i for i, s in enumerate(o) for _ in s]
    out, pos = [''] * len(o), 0
    for p in parts:
        if len(p) == 1:
            for k, c in enumerate(p[0]):
                out[owner[pos + k]] += c
        else:
            out[owner[pos]] += p[1]
        pos += len(p[0])
    for t, before, after in zip(ts, o, out):
        if after != before:
            t.text = after
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


def nearest_p(e):
    e = e.getparent()
    while e is not None and e.tag != W + 'p':
        e = e.getparent()
    return e


def check_office(path):
    """Word·Excel 파일을 열기 전에 확인해, 알기 쉬운 오류 메시지를 냄."""
    with open(path, 'rb') as i:
        head = i.read(4)
    if head == b'\xd0\xcf\x11\xe0':   # 암호화된 Office 파일, 또는 확장자만 바꾼 .doc·.xls
        raise ValueError('비밀번호가 걸린 파일이거나 옛 형식(.doc, .xls)이에요. 비밀번호를 풀거나 .docx·.xlsx로 다시 저장한 뒤 넣어 주세요.')
    if not zipfile.is_zipfile(path):
        raise ValueError('파일을 열 수 없어요. 파일이 손상됐거나 Word·Excel 파일이 아닐 수 있어요.')


def do_docx(src, dst, M):
    import docx
    from lxml import etree
    check_office(src)
    d = docx.Document(src)
    for part in d.part.package.iter_parts():
        if DOCX_PARTS.fullmatch(str(part.partname)):
            xml = hasattr(part, '_element')
            root = part._element if xml else etree.fromstring(part.blob)  # 각주 등은 python-docx가 XML로 읽지 않음
            # 하이퍼링크·텍스트 상자 안의 글자까지 포함 (텍스트 상자 속 문단은 따로 처리)
            # delText: 변경 내용 추적으로 지운 글자, instrText: 필드 코드 (HYPERLINK "mailto:..." 등, 여러 조각으로 나뉘기도 함)
            for p in root.iter(W + 'p'):
                for tag in ('t', 'delText', 'instrText'):
                    mask_group([t for t in p.iter(W + tag) if nearest_p(t) is p], M)
            for e in root.iter():   # w:author, w15:author처럼 네임스페이스가 달라도 이름으로 찾음
                for k, v in e.attrib.items():
                    if k.rsplit('}', 1)[-1] in ATTRS:
                        e.set(k, M.text(v))
            if not xml:
                part._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        for rel in part.rels.values():   # 하이퍼링크 주소(mailto: 등)
            if rel.is_external:
                rel._target = M.text(rel.target_ref)
    d.save(dst)


def do_xlsx(src, dst, M):
    import openpyxl
    check_office(src)
    wb = openpyxl.load_workbook(src)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and not v.startswith('='):
                    c.value = M.text(v)
                elif isinstance(v, int) and not isinstance(v, bool):   # 숫자로 저장된 주민번호·카드번호 등
                    s, z = str(v), '0' + str(v)
                    if M.text(s) != s:
                        c.value = M.text(s)
                    elif any(rx.fullmatch(z) for rx in ZERO):   # 엑셀이 앞자리 0을 지운 휴대전화·주민번호
                        c.value = M.text(z)
                if c.hyperlink:   # 하이퍼링크 주소(mailto: 등)
                    h = c.hyperlink
                    for k in ('target', 'location', 'display', 'tooltip'):
                        if getattr(h, k):
                            setattr(h, k, M.text(getattr(h, k)))
                if c.comment:
                    c.comment.text = M.text(c.comment.text)
    wb.save(dst)


def do_pdf(src, dst, M):
    from pypdf import PdfReader
    try:
        r = PdfReader(src)
        locked = r.is_encrypted and not r.decrypt('')   # 복사·인쇄 제한만 걸린 PDF는 빈 비밀번호로 열림
    except Exception:
        raise ValueError('PDF 파일을 열 수 없어요. 파일이 손상됐을 수 있어요.')
    if locked:
        raise ValueError('비밀번호가 걸린 PDF예요. 비밀번호를 푼 뒤 다시 넣어 주세요.')
    text = '\n\n'.join((p.extract_text() or '') for p in r.pages)
    if not text.strip():
        raise ValueError('텍스트를 읽을 수 없는 PDF입니다(스캔본이면 OCR이 필요합니다).')
    # 줄바꿈으로 끊긴 이메일·번호를 이어 붙여서 놓치지 않게 함
    text = re.sub(r'(?<=[-@])\n(?=[A-Za-z0-9])|(?<=\.)\n(?=[a-z])|(?<=[A-Za-z0-9])\n(?=@)', '', text)
    with open(dst, 'w', encoding='utf-8') as o:
        o.write(M.text(text))


def do_text(src, dst, M):
    with open(src, 'rb') as i:
        raw = i.read()
    if raw[:2] in (b'\xff\xfe', b'\xfe\xff'):   # 엑셀 '유니코드 텍스트' 등 UTF-16 파일
        text = raw.decode('utf-16')
    else:
        try:
            text = raw.decode('utf-8-sig')
        except UnicodeDecodeError:
            text = raw.decode('cp949')   # 엑셀에서 저장한 CSV 등 EUC-KR(CP949) 파일 → UTF-8로 저장
    # CSV는 BOM을 붙여야 한국어 Windows의 엑셀에서 한글이 깨지지 않음. 원본 줄바꿈(\r\n 등)은 그대로 유지
    enc = 'utf-8-sig' if dst.lower().endswith('.csv') else 'utf-8'
    with open(dst, 'w', encoding=enc, newline='') as o:
        o.write(M.text(text))


def leftover(dst, words):
    """결과 파일 전체를 같은 규칙으로 다시 검사해 남은 항목 수를 돌려줌 (웹 도구의 재검사와 같은 범위).
    Word·Excel은 모든 XML 글자, 링크 주소, 시트 이름·그림 대체 텍스트 등 속성 값까지 확인."""
    if dst.endswith(('.docx', '.xlsx')):
        parts = []
        with zipfile.ZipFile(dst) as z:
            for n in z.namelist():
                if not n.endswith(('.xml', '.rels')):
                    continue
                x = z.read(n).decode('utf-8', 'replace')
                parts.append(re.sub(r'<[^>]+>', ' ', x))
                parts += [unquote(v) for v in re.findall(r'Target="([^"]*)"', x)]
                parts += re.findall(r'[\s:](?:name|descr|title|alias|tag|tooltip|display|instr|author|userId)="([^"]*)"', x)
        text = html.unescape('\n'.join(parts))
    else:
        with open(dst, encoding='utf-8-sig') as i:
            text = i.read()
    R = Masker(words)
    R.text(text)
    return R.found


def lock(path, pw):
    """파일 자체에 열기 비밀번호를 설정 (Word/Excel/PDF 프로그램에서 바로 열림)."""
    base, ext = os.path.splitext(path)
    ext = ext.lower()
    dst = base + '_locked' + ext
    if ext == '.pdf':
        from pypdf import PdfReader, PdfWriter
        w = PdfWriter(clone_from=PdfReader(path))
        w.encrypt(pw, algorithm='AES-256')
        with open(dst, 'wb') as o:
            w.write(o)
    elif ext in ('.docx', '.xlsx'):
        from msoffcrypto.format.ooxml import OOXMLFile
        with open(path, 'rb') as i, open(dst, 'wb') as o:
            OOXMLFile(i).encrypt(pw, o)
    else:
        raise ValueError('비밀번호 설정은 docx, xlsx, pdf만 지원합니다.')
    return dst


def ask_pw():
    pw = getpass.getpass('파일 열기 비밀번호(8자 이상): ')
    if len(pw) < 8:
        raise SystemExit('비밀번호는 8자 이상이어야 합니다.')
    if pw != getpass.getpass('비밀번호 확인: '):
        raise SystemExit('비밀번호가 서로 다릅니다.')
    return pw


def main():
    ap = argparse.ArgumentParser(description='개인정보 마스킹 + 파일 비밀번호 설정')
    ap.add_argument('files', nargs='+')
    ap.add_argument('--words', default='')
    how = ap.add_mutually_exclusive_group()
    how.add_argument('--star', action='store_true')
    how.add_argument('--partial', action='store_true')
    ap.add_argument('--map', action='store_true')
    ap.add_argument('--lock', action='store_true')
    ap.add_argument('--lock-only', action='store_true')
    a = ap.parse_args()

    pw = ask_pw() if (a.lock or a.lock_only) else None
    fails = 0   # 하나라도 실패하면 종료 코드 1 (자동화에서 알 수 있게)

    for f in a.files:
        if a.lock_only:
            try:
                print(f'{f} -> {lock(f, pw)}')
            except Exception as e:
                print(f'실패: {f} - {e}'); fails += 1
            continue
        base, ext = os.path.splitext(f)
        ext = ext.lower()
        M = Masker(a.words.split(','), a.star, a.partial)
        try:
            if ext == '.docx':
                dst = base + '_masked.docx'; do_docx(f, dst, M)
            elif ext == '.xlsx':
                dst = base + '_masked.xlsx'; do_xlsx(f, dst, M)
            elif ext == '.pdf':
                dst = base + '_pdf_masked.txt'; do_pdf(f, dst, M)
            elif ext in ('.txt', '.csv', '.md', '.json', '.log', '.tsv'):
                dst = base + '_masked' + ext; do_text(f, dst, M)
            else:
                print(f'건너뜀(지원하지 않는 형식): {f}'); continue
        except Exception as e:
            print(f'실패: {f} - {e}'); fails += 1; continue
        found = ', '.join(f'{k} {v}건' for k, v in M.found.items()) or '찾은 항목 없음'
        print(f'{f} -> {dst}  [{found}]')
        left = leftover(dst, a.words.split(','))
        if left:
            print('  ⚠️ 재검사: 아직 남은 항목이 있어요 —', ', '.join(f'{k} {v}건' for k, v in left.items()),
                  '(시트 이름·수식 등). 결과 파일을 열어 직접 확인하세요.')
        else:
            print('  ✅ 재검사: 남은 개인정보가 없습니다.')
        if a.map and M.map:
            with open(base + '_map.json', 'w', encoding='utf-8') as o:
                json.dump(M.map, o, ensure_ascii=False, indent=1)
        if pw:
            try:
                print('  원본 잠금 ->', lock(f, pw))
            except Exception as e:
                print('  잠금 실패:', e); fails += 1
    return 1 if fails else 0


if __name__ == '__main__':
    sys.exit(main())
