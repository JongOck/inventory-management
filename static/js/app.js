// 현재 활성 탭
let currentTab = 0;
let grids = {};

// 탭별 API 엔드포인트
const TAB_CONFIG = {
  0: { api: '/api/inventory',  label: '재고수불'    },
  1: { api: '/api/warehouse',  label: '창고재고실사' },
  2: { api: '/api/outgoing',   label: '양품출고'    },
  3: { api: '/api/incoming',   label: '입고현황'    },
  4: { api: '/api/shipment',   label: '출하현황'    },
  5: { api: '/api/purchase',   label: '매입영수증'  },
  6: { api: '/api/ledger',     label: '재고원장'    },
  7: { api: '/api/evaluation', label: '재고평가'    },
  8: { api: '/api/incentive',  label: '인센티브'    },
};

// 숫자 포맷 (천단위 쉼표)
function numFmt(params) {
  if (params.value == null || params.value === '') return '';
  const n = parseFloat(params.value);
  if (isNaN(n)) return params.value;
  return Number.isInteger(n) ? n.toLocaleString() : n.toLocaleString(undefined, { maximumFractionDigits: 5 });
}

// AG Grid 초기화
function initGrid(tabIdx, colDefs, rowData) {
  const el = document.getElementById('grid' + tabIdx);
  if (!el) return;

  if (grids[tabIdx]) {
    grids[tabIdx].destroy();
  }

  const gridOptions = {
    columnDefs: colDefs,
    rowData: rowData,
    defaultColDef: {
      sortable: true,
      filter: true,
      resizable: true,
      minWidth: 80,
    },
    rowSelection: 'single',
    suppressMovableColumns: false,
    enableCellTextSelection: true,
    onGridReady: (params) => {
      updateRowCount(tabIdx, rowData.length);
    },
  };

  grids[tabIdx] = agGrid.createGrid(el, gridOptions);
}

// 컬럼 자동 생성 (첫 번째 row 기준)
function autoColDefs(data) {
  if (!data || data.length === 0) return [];
  return Object.keys(data[0]).map((key, i) => {
    const isNum = typeof data[0][key] === 'number';
    return {
      field: key,
      headerName: key,
      width: isNum ? 110 : 130,
      valueFormatter: isNum ? numFmt : undefined,
      type: isNum ? 'numericColumn' : undefined,
      pinned: i === 0 ? 'left' : undefined,
    };
  });
}

// 데이터 로드
async function loadTab(tabIdx) {
  const config = TAB_CONFIG[tabIdx];
  if (!config) return;

  const month = document.getElementById('work-month').value.replace('-', '');
  const url = `${config.api}?work_month=${month}`;

  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();

    const colDefs = autoColDefs(data);
    initGrid(tabIdx, colDefs, data);
  } catch (e) {
    console.error(`탭 ${tabIdx} 로드 오류:`, e);
    alert(`데이터 로드 실패: ${e.message}`);
  }
}

// 현재 탭 조회
function loadCurrentTab() {
  loadTab(currentTab);
}

// 탭 전환
function showTab(tabId) {
  const idx = parseInt(tabId.replace('tab', ''));
  currentTab = idx;

  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));

  document.getElementById(tabId).classList.add('active');
  document.querySelectorAll('.tab-btn')[idx].classList.add('active');

  // 그리드가 없으면 자동 로드
  if (!grids[idx]) loadTab(idx);
}

// 검색 필터
function filterGrid() {
  const keyword = document.getElementById('search-input').value;
  if (grids[currentTab]) {
    grids[currentTab].setGridOption('quickFilterText', keyword);
  }
}

// 행 수 표시
function updateRowCount(tabIdx, count) {
  if (tabIdx === currentTab) {
    document.getElementById('row-count').textContent = `총 ${count.toLocaleString()}건`;
  }
}

// Excel 다운로드
function exportExcel() {
  if (grids[currentTab]) {
    grids[currentTab].exportDataAsCsv({
      fileName: `재고수불_${TAB_CONFIG[currentTab].label}_${document.getElementById('work-month').value}.csv`
    });
  }
}

// 초기화: 오늘 기준 월 설정 후 첫 탭 로드
window.onload = () => {
  const today = new Date();
  const ym = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
  document.getElementById('work-month').value = ym;
  loadTab(0);
};
