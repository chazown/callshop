# 🛒 쇼핑몰 상품 크롤러

이 프로젝트는 Python과 Selenium을 사용하여 쇼핑몰에서 특정 상품을 검색하고,  
상품명 / 가격 / 링크 / 리뷰 수 / 이미지 등을 수집하여 **Excel 파일로 저장**하는 자동화 스크립트입니다.


## 📁 파일 구성

| 파일명               | 설명 |
|---------------------|------|
| `naver_shopping.py` | 네이버 쇼핑에서 상품 정보 수집 후 `네이버쇼핑.xlsx` 파일로 저장 |
| `coupang_scraper.py`| 쿠팡에서 상품 정보 수집 후 `쿠팡정보.xlsx` 파일로 저장 |
| `requirements.txt`  | 프로젝트에 필요한 파이썬 패키지 목록 |
| `LICENSE`           | MIT 오픈소스 라이선스 |
| `README.md`         | 프로젝트 설명 문서 (본 문서) |

---

## ⚙️ 설치 방법

1. Python 3.8 이상 설치
2. 크롬 브라우저 설치
3. ChromeDriver 버전 확인 후 설치 [(다운로드 링크)](https://chromedriver.chromium.org/downloads)
4. 가상 환경 설정 (선택)

```bash
# 프로젝트 클론 후 디렉토리 이동
git clone https://github.com/your-username/shopping-crawler.git
cd shopping-crawler

# 가상환경 생성 및 활성화 (선택)
python -m venv venv
source venv/bin/activate  # Windows는 venv\Scripts\activate

# 필요한 패키지 설치
pip install -r requirements.txt
