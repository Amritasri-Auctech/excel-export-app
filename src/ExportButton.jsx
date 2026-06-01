/* eslint-disable react/prop-types */
import { Download } from 'lucide-react'
import XLSXStyle from 'xlsx-js-style'
import { saveAs } from 'file-saver'

// ── Style definitions ──
const TITLE_STYLE = {
  font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } },
  fill: { fgColor: { rgb: '1F4E79' } },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  border: {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } },
  },
}

const HEADER_STYLE = {
  font: { bold: true, sz: 10, color: { rgb: 'FFFFFF' } },
  fill: { fgColor: { rgb: '2E75B6' } },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  border: {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } },
  },
}

const DATA_STYLE = {
  font: { sz: 10 },
  alignment: { vertical: 'center', wrapText: true },
  border: {
    top: { style: 'thin', color: { rgb: 'CCCCCC' } },
    bottom: { style: 'thin', color: { rgb: 'CCCCCC' } },
    left: { style: 'thin', color: { rgb: 'CCCCCC' } },
    right: { style: 'thin', color: { rgb: 'CCCCCC' } },
  },
}

const DATA_ALT_STYLE = {
  ...DATA_STYLE,
  fill: { fgColor: { rgb: 'EBF3FB' } },
}

const TOTAL_STYLE = {
  font: { bold: true, sz: 10 },
  fill: { fgColor: { rgb: 'FFF2CC' } },
  alignment: { horizontal: 'right', vertical: 'center' },
  border: {
    top: { style: 'medium', color: { rgb: '000000' } },
    bottom: { style: 'medium', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } },
  },
}

const TOTAL_NUM_STYLE = {
  ...TOTAL_STYLE,
  font: { bold: true, sz: 10, color: { rgb: 'C00000' } },
  numFmt: '#,##0',
}

// ── Helpers ──
function setCell(ws, r, c, value, style, numFmt) {
  const addr = XLSXStyle.utils.encode_cell({ r, c })
  const t = typeof value === 'number' ? 'n' : 's'
  ws[addr] = { v: value === '' || value == null ? '' : value, t, s: style }
  if (numFmt) ws[addr].z = numFmt
}

function addMergedTitle(ws, merges, title, row, colCount) {
  setCell(ws, row, 0, title, TITLE_STYLE)
  for (let c = 1; c < colCount; c++) setCell(ws, row, c, '', TITLE_STYLE)
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: colCount - 1 } })
  return row + 1
}

function addHeaders(ws, headers, row) {
  headers.forEach((h, c) => setCell(ws, row, c, h, HEADER_STYLE))
  return row + 1
}

function addDataRows(ws, rows, fields, startRow) {
  rows.forEach((rowData, i) => {
    const style = i % 2 === 0 ? DATA_STYLE : DATA_ALT_STYLE
    fields.forEach((field, c) => {
      const val = rowData[field] ?? ''
      setCell(ws, startRow + i, c, val, style, typeof val === 'number' ? '#,##0' : undefined)
    })
  })
  return startRow + rows.length
}

function addTotalRow(ws, merges, labelText, total, row, labelStartCol, labelEndCol, totalCol, colCount) {
  // empty cells before label
  for (let c = 0; c < labelStartCol; c++) setCell(ws, row, c, '', TOTAL_STYLE)
  // label merged
  setCell(ws, row, labelStartCol, labelText, TOTAL_STYLE)
  for (let c = labelStartCol + 1; c <= labelEndCol; c++) setCell(ws, row, c, '', TOTAL_STYLE)
  if (labelEndCol + 1 < labelStartCol) return row + 1
  merges.push({ s: { r: row, c: labelStartCol }, e: { r: row, c: labelEndCol } })
  // total value
  setCell(ws, row, totalCol, total, TOTAL_NUM_STYLE, '#,##0')
  // remaining cells
  for (let c = totalCol + 1; c < colCount; c++) setCell(ws, row, c, '', TOTAL_STYLE)
  return row + 1
}

function setColWidths(ws, widths) {
  ws['!cols'] = widths.map(w => ({ wch: w }))
}

function setRowHeight(ws, rowCount, height = 18) {
  ws['!rows'] = Array.from({ length: rowCount }, () => ({ hpt: height }))
}

// ══════════════════════════════════════════
// SHEET BUILDERS
// ══════════════════════════════════════════

function buildMembersSheet(data) {
  const ws = {}
  const merges = []
  let row = 0
  const cols = 6

  // DAWAMI
  row = addMergedTitle(ws, merges, 'Anjuman Wazifa-e-Sadat Wa Momineen — Dawami Members April-2026', row, cols)
  row = addHeaders(ws, ['S.No', 'Member No', 'Name', 'Address', 'Amount', 'Motivator'], row)
  row = addDataRows(ws, data.members.dawami, ['SNo', 'MemberNo', 'Name', 'Address', 'Amount', 'Motivator'], row)
  const dawamiTotal = data.members.dawami.reduce((s, r) => s + r.Amount, 0)
  row = addTotalRow(ws, merges, 'TOTAL (Dawami)', dawamiTotal, row, 2, 3, 4, cols)
  row++ // blank row

  // LIFE
  row = addMergedTitle(ws, merges, 'Anjuman Wazifa-e-Sadat Wa Momineen — Life Members April-2026', row, cols)
  row = addHeaders(ws, ['S.No', 'Member No', 'Name', 'Address', 'Amount', 'Motivator'], row)
  row = addDataRows(ws, data.members.life, ['SNo', 'MemberNo', 'Name', 'Address', 'Amount', 'Motivator'], row)
  const lifeTotal = data.members.life.reduce((s, r) => s + r.Amount, 0)
  row = addTotalRow(ws, merges, 'TOTAL (Life Member)', lifeTotal, row, 2, 3, 4, cols)
  row++

  // UMOOMI
  row = addMergedTitle(ws, merges, 'Anjuman Wazifa-e-Sadat Wa Momineen — Umoomi Members April-2026', row, cols)
  row = addHeaders(ws, ['S.No', 'Member No', 'Name', 'Address', 'Amount', 'Motivator'], row)
  row = addDataRows(ws, data.members.umoomi, ['SNo', 'MemberNo', 'Name', 'Address', 'Amount', 'Motivator'], row)
  const umoomiTotal = data.members.umoomi.reduce((s, r) => s + r.Amount, 0)
  row = addTotalRow(ws, merges, 'TOTAL (Chanda Umoomi)', umoomiTotal, row, 2, 3, 4, cols)

  ws['!ref'] = XLSXStyle.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row, c: cols - 1 } })
  ws['!merges'] = merges
  setColWidths(ws, [6, 12, 32, 40, 12, 32])
  setRowHeight(ws, row + 1)
  return ws
}

function buildAyanatSheet(data) {
  const ws = {}
  const merges = []
  let row = 0
  const cols = 5
  const total = data.ayanat.reduce((s, r) => s + r.Amount, 0)

  row = addMergedTitle(ws, merges, 'Anjuman Wazifa-e-Sadat Wa Momineen (Ayanat) April-2026', row, cols)
  row = addHeaders(ws, ['S.No', 'Name', 'Place', 'Amount', 'Motivator'], row)
  row = addDataRows(ws, data.ayanat, ['SNo', 'Name', 'Place', 'Amount', 'Motivator'], row)
  row = addTotalRow(ws, merges, 'TOTAL', total, row, 1, 2, 3, cols)
  row = addTotalRow(ws, merges, 'Total as per Income & Expenditure (Sheet)', total, row, 1, 2, 3, cols)

  ws['!ref'] = XLSXStyle.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row, c: cols - 1 } })
  ws['!merges'] = merges
  setColWidths(ws, [6, 32, 20, 12, 32])
  setRowHeight(ws, row + 1)
  return ws
}

function buildLoanSheet(data) {
  const ws = {}
  const merges = []
  let row = 0
  const cols = 9
  const total = data.returnOfLoan.reduce((s, r) => s + (typeof r.PaidThisMonth === 'number' ? r.PaidThisMonth : 0), 0)

  row = addMergedTitle(ws, merges, 'Anjuman Wazifa-e-Sadat Wa Momineen — Return of Wazifa April-2026', row, cols)
  row = addHeaders(ws, ['S.No', 'Wazifa No', 'Name', 'Place', 'O.P. Balance', 'Paid This Month', 'Total Refunded', 'Closing Balance', 'Motivator'], row)
  row = addDataRows(ws, data.returnOfLoan, ['SNo', 'WazifaNo', 'Name', 'Place', 'OPBalance', 'PaidThisMonth', 'TotalRefunded', 'ClosingBalance', 'Motivator'], row)
  row = addTotalRow(ws, merges, 'TOTAL', total, row, 3, 4, 5, cols)
  row = addTotalRow(ws, merges, 'Total as per Income & Expenditure (Sheet)', total, row, 3, 4, 5, cols)

  ws['!ref'] = XLSXStyle.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row, c: cols - 1 } })
  ws['!merges'] = merges
  setColWidths(ws, [6, 12, 28, 18, 14, 16, 16, 16, 30])
  setRowHeight(ws, row + 1)
  return ws
}

function buildFundSheet(data) {
  const ws = {}
  const merges = []
  let row = 0
  const cols = 7
  const total = data.fund.reduce((s, r) => s + r.Amount, 0)

  row = addMergedTitle(ws, merges, 'Anjuman Wazifa-e-Sadat Wa Momineen (Namzad Fund) April-2026', row, cols)
  row = addHeaders(ws, ['S.No', 'Fund No.', 'Name of Fund', 'Place', 'Donor', 'Amount', 'Motivator'], row)
  row = addDataRows(ws, data.fund, ['SNo', 'FundNo', 'FundName', 'Place', 'Donor', 'Amount', 'Motivator'], row)
  row = addTotalRow(ws, merges, 'TOTAL', total, row, 2, 4, 5, cols)
  row = addTotalRow(ws, merges, 'Total as per Income & Expenditure (Sheet)', total, row, 2, 4, 5, cols)

  ws['!ref'] = XLSXStyle.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row, c: cols - 1 } })
  ws['!merges'] = merges
  setColWidths(ws, [6, 10, 28, 18, 28, 12, 30])
  setRowHeight(ws, row + 1)
  return ws
}

function buildCalenderSheet(data) {
  const ws = {}
  const merges = []
  let row = 0
  const cols = 5
  const total = data.calender.reduce((s, r) => s + r.Amount, 0)

  row = addMergedTitle(ws, merges, 'Anjuman Wazifa-e-Sadat Wa Momineen (Calender) April-2026', row, cols)
  row = addHeaders(ws, ['S.No', 'Name', 'Place', 'Amount', 'Motivator'], row)
  row = addDataRows(ws, data.calender, ['SNo', 'Name', 'Place', 'Amount', 'Motivator'], row)
  row = addTotalRow(ws, merges, 'TOTAL', total, row, 1, 2, 3, cols)
  row = addTotalRow(ws, merges, 'Total as per Income & Expenditure (Sheet)', total, row, 1, 2, 3, cols)

  ws['!ref'] = XLSXStyle.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row, c: cols - 1 } })
  ws['!merges'] = merges
  setColWidths(ws, [6, 32, 20, 12, 32])
  setRowHeight(ws, row + 1)
  return ws
}

function buildKhumsSheet(data) {
  const ws = {}
  const merges = []
  let row = 0
  const cols = 5
  const total = data.khums.reduce((s, r) => s + r.Amount, 0)

  row = addMergedTitle(ws, merges, 'Anjuman Wazifa-e-Sadat Wa Momineen — Khums April-2026', row, cols)
  row = addHeaders(ws, ['S.No', 'Name of Donor', 'Place', 'Amount', 'Motivator'], row)
  row = addDataRows(ws, data.khums, ['SNo', 'DonorName', 'Place', 'Amount', 'Motivator'], row)
  row = addTotalRow(ws, merges, 'TOTAL', total, row, 1, 2, 3, cols)
  row = addTotalRow(ws, merges, 'Total as per Income & Expenditure (Sheet)', total, row, 1, 2, 3, cols)

  ws['!ref'] = XLSXStyle.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row, c: cols - 1 } })
  ws['!merges'] = merges
  setColWidths(ws, [6, 32, 20, 12, 32])
  setRowHeight(ws, row + 1)
  return ws
}

function buildIncomeExpenditureSheet(data) {
  const ws = {}
  const merges = []
  let row = 0
  const cols = 15
  const ie = data.incomeExpenditure

  row = addMergedTitle(ws, merges, 'Anjuman Wazifa-e-Sadat Wa Momineen — Income & Expenditure April-2026', row, cols)
  row = addHeaders(ws, [
    'S.No', 'Name of Local Secretary', 'Place',
    'Chanda', 'Dawami', 'Life', 'Ayanat', 'Fund', 'Calender',
    'Qarze Hasana', 'Khums', 'Total Income', 'Comm of L/S', 'Postage', 'Total Exp'
  ], row)

  ie.forEach((r, i) => {
    const style = i % 2 === 0 ? DATA_STYLE : DATA_ALT_STYLE
    const fields = ['SNo','LocalSecretary','Place','Chanda','Dawami','Life','Ayanat','Fund','Calender','QarzeHasana','Khums','TotalIncome','CommLS','Postage','TotalExp']
    fields.forEach((field, c) => {
      const val = r[field] ?? ''
      setCell(ws, row + i, c, val, style, typeof val === 'number' ? '#,##0' : undefined)
    })
  })
  row += ie.length

  // totals row
  const totals = ['Chanda','Dawami','Life','Ayanat','Fund','Calender','QarzeHasana','Khums','TotalIncome','CommLS','Postage','TotalExp']
    .map(k => ie.reduce((s, r) => s + (r[k] || 0), 0))

  setCell(ws, row, 0, '', TOTAL_STYLE)
  setCell(ws, row, 1, 'TOTAL', TOTAL_STYLE)
  setCell(ws, row, 2, '', TOTAL_STYLE)
  merges.push({ s: { r: row, c: 1 }, e: { r: row, c: 2 } })
  totals.forEach((val, i) => setCell(ws, row, 3 + i, val, TOTAL_NUM_STYLE, '#,##0'))

  ws['!ref'] = XLSXStyle.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row, c: cols - 1 } })
  ws['!merges'] = merges
  setColWidths(ws, [5, 28, 16, 10, 10, 10, 10, 10, 10, 12, 10, 12, 12, 10, 10])
  setRowHeight(ws, row + 1)
  return ws
}

// ── Main export function ──
function buildWorkbook(data) {
  const wb = XLSXStyle.utils.book_new()
  XLSXStyle.utils.book_append_sheet(wb, buildMembersSheet(data), 'members')
  XLSXStyle.utils.book_append_sheet(wb, buildAyanatSheet(data), 'ayanat')
  XLSXStyle.utils.book_append_sheet(wb, buildLoanSheet(data), 'return of loan')
  XLSXStyle.utils.book_append_sheet(wb, buildFundSheet(data), 'fund')
  XLSXStyle.utils.book_append_sheet(wb, buildCalenderSheet(data), 'calender')
  XLSXStyle.utils.book_append_sheet(wb, buildKhumsSheet(data), 'khums')
  XLSXStyle.utils.book_append_sheet(wb, buildIncomeExpenditureSheet(data), 'income expenditure')
  return wb
}

const ExportButton = ({ data, fileName = 'April-2026.xlsx' }) => {
  const handleExport = () => {
    const wb = buildWorkbook(data)
    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'array' })
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), fileName)
  }

  return (
    <button
      onClick={handleExport}
      className="bg-blue-600 text-white border border-blue-700 px-5 py-2.5 rounded hover:bg-blue-700 transition-colors flex items-center gap-2 font-medium shadow"
    >
      <Download className="w-4 h-4" />
      Export Excel
    </button>
  )
}

export default ExportButton
