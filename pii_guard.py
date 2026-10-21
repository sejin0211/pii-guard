#!/usr/bin/env python3
"""AI에 넘기기 전 개인정보 마스킹 (Word / Excel / PDF / 텍스트) + 선택적 비밀번호 설정.

설치:  pip install python-docx openpyxl pypdf cryptography msoffcrypto-tool
사용:  python pii_guard.py 파일1 [파일2 ...] [--words 홍길동,김영희] [--star] [--map] [--lock | --lock-only]

  --words    자동으로 못 찾는 이름·회사명 등 (쉼표 구분)
  --star     토큰 대신 별표(*)로 가림 (복원 불가)
  --map      토큰→원본 대응표(_map.json) 저장. 이 파일에는 개인정보가 들어 있으니 따로 보관하세요.
  --lock       원본 파일(개인정보 포함)에 열기 비밀번호를 걸어 '이름_locked.확장자'로 함께 저장 (docx, xlsx, pdf)
  --lock-only  마스킹은 하지 않고 비밀번호만 설정
결과: 원본은 건드리지 않고 '이름_masked.확장자' 로 저장합니다. PDF는 '이름_pdf_masked.txt'(텍스트만 추출).
"""
import argparse, getpass, json, os, re, time

N, X = r'(?<!\d)', r'(?!\d)'
RULES = [
    ('주민등록번호', re.compile(N + r'\d{6}[- ]?[1-8]\d{6}' + X)),  # 5~8: 외국인등록번호
    ('카드번호',     re.compile(N + r'\d{4}[- ]?\d{4}[- ]?\d{4}[- ]?\d{4}' + X)),
    ('휴대전화',     re.compile(N + r'01[016789][-. ]?\d{3,4}[-. ]?\d{4}' + X)),
    ('전화번호',     re.compile(N + r'0(?:2|[3-6]\d)[-. ]?\d{3,4}[-. ]?\d{4}' + X)),
    ('이메일',       re.compile(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+')),
    ('사업자번호',   re.compile(N + r'\d{3}-\d{2}-\d{5}' + X)),
    ('여권번호',     re.compile(r'(?<![A-Za-z0-9])[MSROD]\d{8}(?!\d)')),
    ('IP주소',       re.compile(r'(?<![\d.])(?:\d{1,3}\.){3}\d{1,3}(?!\d|\.\d)')),  # 문장 끝 마침표는 허용
    ('계좌번호',     re.compile(N + r'\d{2,6}-\d{2,6}-\d{2,8}(?:-\d{1,4})?' + X)),
]
DATE = re.compile(r'(19|20)\d{2}-\d{2}-\d{2}')  # 날짜를 계좌번호로 오인하지 않도록


class Masker:
    def __init__(self, words=(), star=False):
        self.star, self.map, self.seen, self.cnt, self.found = star, {}, {}, {}, {}
        self.rules = list(RULES)
        ws = sorted({w.strip() for w in words if w.strip()}, key=len, reverse=True)
        if ws:
            self.rules.insert(0, ('지정어', re.compile('|'.join(map(re.escape, ws)))))

    def text(self, s):
        return ''.join(p[-1] for p in self.parts(s))

    def parts(self, s):
        """원문을 [그대로인 조각] 또는 [찾은 값, 바꿀 값]의 목록으로 나눔 (Word 서식 유지에 사용).
        찾은 값은 잠시 보호 문자로 바꿔 두어 뒤의 규칙이 다시 건드리지 않게 함 (웹 도구와 같은 방식)."""
        keep = []

        def hit(label, v):
            r = self._sub(label, v)
            if label == '계좌번호' and r == v:   # 날짜는 그대로
                return v
            keep.append([v, r])
            return '\ue000' + ''.join(chr(0xE010 + int(d)) for d in str(len(keep) - 1)) + '\ue001'

        for label, rx in self.rules:
            s = rx.sub(lambda m, l=label: hit(l, m.group(0)), s)
        out = []
        for i, seg in enumerate(re.split('\ue000([\ue010-\ue019]+)\ue001', s)):
            if i % 2:
                out.append(keep[int(''.join(str(ord(c) - 0xE010) for c in seg))])
            elif seg:
                out.append([seg])
        return out

    def _sub(self, label, m):
        if label == '계좌번호' and DATE.fullmatch(m):
            return m
        self.found[label] = self.found.get(label, 0) + 1
        if self.star:
            return re.sub(r'[0-9A-Za-z가-힣]', '*', m)
        key = (label, m)
        if key not in self.seen:
            n = self.cnt[label] = self.cnt.get(label, 0) + 1
            self.seen[key] = f'[{label}_{n}]'
            self.map[self.seen[key]] = m
        return self.seen[key]


W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
DOCX_PARTS = re.compile(r'/word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml')


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


def do_docx(src, dst, M):
    import docx
    from lxml import etree
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
            if not xml:
                part._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        for rel in part.rels.values():   # 하이퍼링크 주소(mailto: 등)
            if rel.is_external:
                rel._target = M.text(rel.target_ref)
    d.save(dst)


def do_xlsx(src, dst, M):
    import openpyxl
    wb = openpyxl.load_workbook(src)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and not v.startswith('='):
                    c.value = M.text(v)
                elif isinstance(v, int) and not isinstance(v, bool):
                    # 숫자로 저장된 주민번호·카드번호 등. 엑셀이 앞자리 0을 지운 번호(010…, 2000년대생 주민번호)도 확인
                    for s in (str(v), '0' + str(v)):
                        m = M.text(s)
                        if m != s:
                            c.value = m
                            break
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
    text = '\n\n'.join((p.extract_text() or '') for p in PdfReader(src).pages)
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
    with open(dst, 'w', encoding='utf-8', newline='') as o:   # 원본 줄바꿈(\r\n 등) 그대로 유지
        o.write(M.text(text))


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
    ap.add_argument('--star', action='store_true')
    ap.add_argument('--map', action='store_true')
    ap.add_argument('--lock', action='store_true')
    ap.add_argument('--lock-only', action='store_true')
    a = ap.parse_args()

    pw = ask_pw() if (a.lock or a.lock_only) else None

    for f in a.files:
        if a.lock_only:
            try:
                print(f'{f} -> {lock(f, pw)}')
            except Exception as e:
                print(f'실패: {f} - {e}')
            continue
        base, ext = os.path.splitext(f)
        ext = ext.lower()
        M = Masker(a.words.split(','), a.star)
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
            print(f'실패: {f} - {e}'); continue
        found = ', '.join(f'{k} {v}건' for k, v in M.found.items()) or '찾은 항목 없음'
        print(f'{f} -> {dst}  [{found}]')
        if a.map and M.map:
            with open(base + '_map.json', 'w', encoding='utf-8') as o:
                json.dump(M.map, o, ensure_ascii=False, indent=1)
        if pw:
            try:
                print('  원본 잠금 ->', lock(f, pw))
            except Exception as e:
                print('  잠금 실패:', e)


if __name__ == '__main__':
    main()
