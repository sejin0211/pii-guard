# ocr/ — 스캔본 PDF용 OCR 엔진

웹 도구(`index.html`)가 **글자가 없는(스캔본) PDF 쪽**을 만났을 때만 이 폴더의 파일을 불러와 브라우저 안에서 글자를 읽습니다.
같은 사이트에서 불러오므로 외부 서버로 나가는 요청은 없고, 브라우저 저장소(IndexedDB)에도 저장하지 않습니다(`cacheMethod: 'none'`).

> `index.html`을 파일로 바로 열면(`file://`) 브라우저 보안 때문에 이 파일들을 불러올 수 없어 OCR이 동작하지 않습니다.
> 배포된 페이지(https://sejin0211.github.io/pii-guard/)에서 쓰세요.

## 출처와 무결성

모두 npm 공식 저장소의 패키지에서 꺼낸 파일을 **수정하지 않고** 그대로 넣었습니다. 패키지를 받을 때 npm이 공개한 `sha512` 무결성 값과 일치하는 것을 확인했습니다.

| 파일 | 패키지 | 라이선스 | SHA-256 |
|---|---|---|---|
| `tesseract.min.js` | tesseract.js@7.0.0 `dist/` | Apache-2.0 | `000c27d9cd0def655f77b36c72a389c0ab13793aa31cb4d7aab56d09c0afbc7e` |
| `worker.min.js` | tesseract.js@7.0.0 `dist/` | Apache-2.0 | `576b7df7e3393e137e51849357c9adb53fe7ac1bb69bfa06cf3d61520f182c6d` |
| `tesseract-core-simd-lstm.wasm.js` | tesseract.js-core@7.0.0 | Apache-2.0 | `c58b46a4c796c0b8afccf77591d5b875b6896b45d402bbce8caa6f5362447b38` |
| `kor.traineddata.gz` | @tesseract.js-data/kor@1.0.0 `4.0.0_best_int/` | Apache-2.0 (Tesseract tessdata) | `78c21276ab14c9bb734d83be1055d9fe5469a4e7e977c51ad385be5737e61126` |
| `eng.traineddata.gz` | @tesseract.js-data/eng@1.0.0 `4.0.0_best_int/` | Apache-2.0 (Tesseract tessdata) | `45b4cb346724ac1774f1c36f42f182b887bcdb28ebe63e6fff90ac41f3fcff91` |

`*.LICENSE.txt`, `LICENSE-*` 파일은 각 패키지에 들어 있던 라이선스 문서입니다.

영어 데이터(`eng`)는 이메일처럼 영문이 섞인 정보를 읽기 위해 함께 씁니다. 한국어 데이터만 쓰면 `gildong.hong@example.com`이 엉뚱한 글자로 읽혀 가려지지 않았습니다.

## 확인하는 방법

```bash
sha256sum ocr/*
```

## 버전을 올릴 때

1. 새 버전 패키지에서 같은 파일을 꺼내 바꿉니다(`npm pack tesseract.js@버전` 등).
2. 위 표의 버전과 SHA-256을 고칩니다.
3. 웹 도구에서 스캔본 PDF를 넣어 OCR이 되는지, 개발자 도구 Network 탭에 외부 요청이 없는지 확인합니다.
