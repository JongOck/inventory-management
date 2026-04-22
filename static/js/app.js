let currentTab = 0;
let grids = {};

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

// 컬럼 한글명 매핑
const COL_NAMES = {
  item_code: '품목코드', item_name: '품목명', specification: '규격', unit: '단위',
  category: '구분', reference_year: '기준년도', reference_month: '기준월',
  major_category: '대구분', minor_category: '중구분', main_supplier_name: '주거래처',
  purchase_vat: '매입부가세', last_updated: '최종업데이트',
  // 재고수불
  beginning_unit_price: '기초단가', beginning_quantity: '기초수량', beginning_amount: '기초금액',
  incoming_unit_price: '입고단가', incoming_quantity: '입고수량', incoming_amount: '입고금액',
  outgoing_unit_price: '출고단가', outgoing_quantity: '출고수량', outgoing_amount: '출고금액',
  current_unit_price: '현재단가', current_quantity: '현재수량', current_amount: '현재금액',
  misc_profit_amount: '기타이익금액', incentive_amount: '인센티브금액',
  // 출하/입고
  shipment_quantity: '출하수량', unit_price: '단가', amount: '금액',
  won_amount_shipment: '원화금액(출하)', vat_shipment: '부가세(출하)',
  won_amount_sales: '원화금액(매출)', vat_sales: '부가세(매출)',
  total_amount_shipment: '총금액(출하)', total_amount_sales: '총금액(매출)',
  total_amount: '총금액', weight: '중량',
  management_quantity: '입고수량', won_amount: '금액',
  supplier_name: '거래처명', supplier_code: '거래처코드',
  vat: '부가세',
  // 재고평가
  receipt_quantity: '입고수량', receipt_amount: '입고금액',
  substitution_quantity: '대체수량', substitution_amount: '대체금액',
  inventory_quantity: '재고수량', inventory_unit_price: '재고단가', inventory_amount: '재고금액',
  // 인센티브
  no: 'NO', supplier_code: '거래처코드', supplier_name: '거래처명',
  sum_won_amount: '원화금액합계', sum_vat_amount: '부가세합계', sum_total_amount: '총금액합계',
  ratio: '비율', incentive: '인센티브',
  // 대체
  quantity: '수량', substitution_type: '대체유형', warehouse: '창고',
  output_number: '출고번호', request_number: '요청번호', department: '부서',
  manager: '담당자', customer_code: '거래처코드', customer_name: '거래처명',
  foreign_currency_amount: '외화금액', weight_unit: '중량단위',
  account_type: '계정유형', requesting_department: '요청부서',
};

// 숫자 컬럼 판별
const NUM_KEYWORDS = ['quantity','amount','price','vat','weight','ratio','incentive','unit_price','won_amount'];
function isNumeric(field) {
  return NUM_KEYWORDS.some(k => field.includes(k));
}

function numFmt(params) {
  if (params.value == null || params.value === '') return '';
  const n = parseFloat(params.value);
  if (isNaN(n)) return params.value;
  return Number.isInteger(n)
    ? n.toLocaleString('ko-KR')
    : parseFloat(n.toFixed(5)).toLocaleString('ko-KR', { maximumFractionDigits: 5 });
}

// work_month → YYYY/MM 형식
function toSlashMonth(ym) {
  if (!ym) return '';
  const clean = ym.replace('-', '');
  return clean.length >= 6 ? `${clean.slice(0,4)}/${clean.slice(4,6)}` : clean;
}

function buildColDefs(data) {
  if (!data || data.length === 0) return [];
  return Object.keys(data[0]).map((key, i) => {
    const num = isNumeric(key);
    return {
      field: key,
      headerName: COL_NAMES[key] || key,
      width: num ? 120 : (key === 'item_name' ? 200 : 130),
      valueFormatter: num ? numFmt : undefined,
      type: num ? 'numericColumn' : undefined,
      pinned: i < 2 ? 'left' : undefined,
    };
  });
}

function initGrid(tabIdx, colDefs, rowData) {
  const el = document.getElementById('grid' + tabIdx);
  if (!el) return;
  if (grids[tabIdx]) { grids[tabIdx].destroy(); }

  grids[tabIdx] = agGrid.createGrid(el, {
    columnDefs: colDefs,
    rowData: rowData,
    defaultColDef: { sortable: true, filter: true, resizable: true, minWidth: 80 },
    rowSelection: 'single',
    enableCellTextSelection: true,
    onGridReady: () => updateRowCount(tabIdx, rowData.length),
  });
}

async function loadTab(tabIdx) {
  const config = TAB_CONFIG[tabIdx];
  if (!config) return;

  const rawMonth = document.getElementById('work-month').value.replace('-', '');
  const slashMonth = toSlashMonth(rawMonth);
  const url = `${config.api}?work_month=${rawMonth}`;

  document.getElementById('row-count').textContent = '로딩 중...';
  try {
    const res = await fetch(url);
    if (!res.ok) {
      const err = await res.text();
      throw new Error(`HTTP ${res.status}: ${err}`);
    }
    const data = await res.json();
    initGrid(tabIdx, buildColDefs(data), data);
  } catch (e) {
    console.error(`탭 ${tabIdx} 오류:`, e);
    alert(`데이터 로드 실패: ${e.message}`);
    document.getElementById('row-count').textContent = '';
  }
}

function loadCurrentTab() { loadTab(currentTab); }

function showTab(tabId) {
  const idx = parseInt(tabId.replace('tab', ''));
  currentTab = idx;
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  document.getElementById(tabId).classList.add('active');
  document.querySelectorAll('.tab-btn')[idx].classList.add('active');
  if (!grids[idx]) loadTab(idx);
}

function filterGrid() {
  const kw = document.getElementById('search-input').value;
  if (grids[currentTab]) grids[currentTab].setGridOption('quickFilterText', kw);
}

function updateRowCount(tabIdx, count) {
  if (tabIdx === currentTab)
    document.getElementById('row-count').textContent = `총 ${count.toLocaleString('ko-KR')}건`;
}

function exportExcel() {
  if (grids[currentTab]) {
    grids[currentTab].exportDataAsCsv({
      fileName: `재고수불_${TAB_CONFIG[currentTab].label}_${document.getElementById('work-month').value}.csv`
    });
  }
}

window.onload = () => {
  const today = new Date();
  const ym = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
  document.getElementById('work-month').value = ym;
  loadTab(0);
};
