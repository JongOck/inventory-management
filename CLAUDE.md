# 재고수불현황 관리 시스템 (Inventory Management)

## 프로젝트 개요
- 담당자: 김종옥 선임 (경영지원본부 시스템관리팀)
- 원본: Tkinter 데스크탑 앱(inventory_management_07_07.py) → FastAPI + AG Grid 웹앱으로 전환
- GitHub: https://github.com/JongOck/inventory-management.git

## 기술 스택
- Backend: FastAPI + pg8000 (pure Python PostgreSQL driver)
- Frontend: AG Grid Community v31 + 순수 HTML/CSS/JS
- Font: Noto Sans KR (Google Fonts)
- DB: PostgreSQL @ foodall.co.kr:5432, database: ecommerce

## 로컬 개발 서버 실행
```bash
cd C:/Users/Foodallmarket/Documents/inventory-management
uvicorn main:app --reload --port 8002
```
브라우저: http://localhost:8002

## DB 연결 정보 (.env 파일, git에 포함 안됨)
```
DB_NAME=ecommerce
DB_USER=postgres
DB_PASSWORD=!rjtkd4279
DB_HOST=foodall.co.kr
DB_PORT=5432
```

## 주요 파일 구조
```
inventory-management/
├── main.py              # FastAPI 앱 진입점, 9개 라우터 등록
├── database.py          # pg8000 DB 연결, query() 함수
├── .env                 # DB 비밀번호 (git 제외)
├── requirements.txt     # 의존성
├── routers/
│   ├── inventory.py     # 재고수불  → mds_basic_data
│   ├── warehouse.py     # 창고재고실사 → master
│   ├── outgoing.py      # 양품출고  → mds_shipment_status
│   ├── incoming.py      # 입고현황  → mds_purchase_receipt_status
│   ├── shipment.py      # 출하현황  → mds_shipment_status
│   ├── purchase.py      # 매입영수증 → mds_purchase_receipt_status
│   ├── ledger.py        # 재고원장  → mds_monthly_inventory_status
│   ├── evaluation.py    # 재고평가  → mds_inventory_evaluation
│   └── incentive.py     # 인센티브  → mds_incentive_result
└── static/
    ├── index.html       # 9탭 레이아웃
    ├── css/style.css    # Noto Sans KR, 블루 테마 (#3b82f6)
    └── js/app.js        # AG Grid 초기화, 한글 컬럼명 매핑
```

## DB 테이블 날짜 형식
- 모든 테이블의 reference_month는 **YYYY/MM** 형식 (예: 2025/06)
- 프론트에서 YYYYMM → YYYY/MM 변환: `toSlashMonth()` 함수 사용
- mds_basic_data의 reference_year는 '2025', '2025/06', '2025-06' 혼재 → IN (%s,%s,%s) 쿼리 사용

## API 진단 엔드포인트
- GET /api/health → 서버 상태
- GET /api/dbtest → DB 연결 테스트

## 배포 계획
- 최종 목표: foodall.co.kr 서버에 직접 배포 (SSH 접근 필요 → 이거상 팀장님 확인 필요)
- Railway 배포는 DB 방화벽 문제로 불가 (foodall.co.kr:5432 접근 차단됨)

## 진행 중인 작업
- [ ] UI 개선: 인트라넷(intra.foodall.co.kr) 스타일처럼 부드럽게
- [ ] 그룹 컬럼 헤더 추가 (이월재고/입고내역/출고내역 등)
- [ ] 액션 버튼 추가 (원스탑작업, 불러오기, 현재고계산 등)
- [ ] 합계 행 상단 고정
- [ ] foodall.co.kr 서버 최종 배포
