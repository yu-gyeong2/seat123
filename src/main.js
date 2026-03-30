import * as XLSX from 'xlsx'
import './style.css'

const state = {
  seats: [],
  students: [],
  fixedAssignments: new Map(),
  preAssignments: new Map(),
  // Ctrl+자동배치로 사전배치를 "반영"한 이후엔 사전배치 좌석의 학생 이름을 보여줍니다.
  presetApplied: false,
  blockedSeats: new Set(),
  /** 'teacher' = 교탁이 위(교실 앞), 'student' = 학생이 교실 뒤에서 보는 배치 */
  viewPerspective: 'teacher',
}

const app = document.querySelector('#app')
app.innerHTML = `
  <main class="container">
    <header class="header">
      <h1>❤️ 학급 자리배치 프로그램</h1>
      <p>학생 명단을 입력하고 좌석을 자동으로 배치하세요.</p>
      <p class="header-credit">Made by 유경T</p>
    </header>

    <section class="panel controls">
      <div class="field-group wide seat-layout-field">
        <label for="seat-layout">자리 형태</label>
        <select id="seat-layout">
          <option value="individual">개별</option>
          <option value="pair">짝꿍 (옆자리 2명씩 붙음)</option>
          <option value="group">모둠 (모둠 수 × 모둠당 인원)</option>
        </select>
      </div>
      <div class="seat-setup-row">
        <div class="field-group">
          <label id="label-rows" for="rows">행(줄) 수</label>
          <input id="rows" type="number" min="1" max="10" value="5" />
        </div>
        <div class="field-group">
          <label id="label-cols" for="cols">열(칸) 수</label>
          <input id="cols" type="number" min="1" max="10" value="6" />
        </div>
        <div class="field-group seat-setup-btn-wrap">
          <button id="build-seat-map" type="button" class="seat-setup-btn">
            책상 배열
          </button>
        </div>
      </div>
      <div class="field-group wide">
        <label for="student-input">학생 이름 (줄바꿈 또는 쉼표로 구분)</label>
        <textarea id="student-input" rows="6" placeholder="강대현, 김건호, 김도연"></textarea>
      </div>
      <div class="field-group wide">
        <div class="row-actions">
          <input id="group-name" type="text" placeholder="그룹 이름 (예: 3반-1)" />
          <button id="save-students" type="button">명단 저장</button>
          <div class="row-actions-load">
            <select id="saved-groups">
              <option value="">저장된 그룹 선택</option>
            </select>
            <button id="load-students" type="button">명단 불러오기</button>
            <button id="delete-saved-students" type="button" class="delete-list-btn">명단 삭제</button>
          </div>
        </div>
      </div>
      <div class="field-group wide secret-toggle">
        <button id="secret-toggle" type="button" class="secret-btn" aria-label="옵션 보기">
          🥚 옵션
        </button>
      </div>

      <div id="advanced-controls" class="advanced-controls">
        <section class="panel legend">
          <p><strong>사용 방법</strong></p>
          <ul>
            <li>자리 만들기: 행·열 입력(개별·짝꿍) 또는 모둠 수·모둠당 인원(모둠)</li>
            <li>자리 없애기: 빈 좌석 클릭(해당 자리에는 학생 배치 x)</li>
            <li>사전 좌석 배치: 학생 선택 후 좌석 클릭 (⭐사전 배치후 명단 저장)</li>
            <li>사전 배치 반영은 <strong class="ctrl-key-hint">Ctrl</strong> 키를 누른 채 자리 배치 클릭</li>
            <li>특정 학생 공개 고정: 사전배치 후 <strong class="shift-key-hint">Shift</strong>+좌석클릭(초록색으로 변함)</li>
            <li>학생 분리: 분리할 학생 쌍에 입력(랜덤하게 떨어진 채로 배치됨)</li>
          </ul>
          <p id="status">좌석판을 먼저 만들어 주세요.</p>
        </section>
        <div class="field-group wide">
          <label for="separate-input">▶️ 분리할 학생 쌍 (한 줄에 1쌍)</label>
          <textarea id="separate-input" rows="4" placeholder="김건호-김도연&#10;김건호-강대현"></textarea>
        </div>
        <div class="field-group wide">
          <label for="preset-student-select">▶️ 사전 배치 학생 선택 후 좌석 클릭</label>
          <div class="preset-row">
            <select id="preset-student-select">
              <option value="">선택 안 함 (일반 모드)</option>
            </select>
            <button id="clear-preassignments" type="button">사전 배치 완료</button>
          </div>
        </div>
        <div class="field-group wide">
          <p class="secret-title">▶️ 사전 배치 현황</p>
          <div id="preassigned-list" class="preassigned-list"></div>
        </div>
      </div>
    </section>

    <section class="panel">
      <div id="seat-board" class="seat-board perspective-teacher">
        <div class="teacher">교탁</div>
        <div id="seat-grid" class="seat-grid"></div>
      </div>
      <div class="seat-actions">
        <div class="seat-actions-left">
          <label class="effect-toggle-label" for="effect-toggle">
            <input id="effect-toggle" type="checkbox" />
            효과 켜기
          </label>
          <button id="seat-reset-display" type="button" class="seat-reset-btn">초기화</button>
        </div>
        <div class="seat-actions-row">
          <button id="auto-assign" class="primary seat-primary" type="button">자리 배치 start</button>
          <button id="view-perspective-toggle" type="button" class="perspective-toggle-btn">학생뷰</button>
        </div>
      </div>
      <div class="seat-export">
        <button id="export-seat-excel" type="button" class="export-excel-btn">좌석 배치도 엑셀 저장</button>
      </div>
    </section>

    <div id="countdown-overlay" class="countdown-overlay" aria-hidden="true">
      <div id="countdown-number" class="countdown-number">5</div>
    </div>
  </main>
`

const rowsInput = document.querySelector('#rows')
const colsInput = document.querySelector('#cols')
const studentInput = document.querySelector('#student-input')
const groupNameInput = document.querySelector('#group-name')
const presetStudentSelect = document.querySelector('#preset-student-select')
const clearPreassignmentsBtn = document.querySelector('#clear-preassignments')
const saveStudentsBtn = document.querySelector('#save-students')
const loadStudentsBtn = document.querySelector('#load-students')
const deleteSavedStudentsBtn = document.querySelector('#delete-saved-students')
const savedGroupsSelect = document.querySelector('#saved-groups')
const seatBoardEl = document.querySelector('#seat-board')
const seatGrid = document.querySelector('#seat-grid')
const viewPerspectiveToggleBtn = document.querySelector('#view-perspective-toggle')
const exportSeatExcelBtn = document.querySelector('#export-seat-excel')
const statusText = document.querySelector('#status')
const buildSeatMapBtn = document.querySelector('#build-seat-map')
const seatLayoutSelect = document.querySelector('#seat-layout')
const labelRows = document.querySelector('#label-rows')
const labelCols = document.querySelector('#label-cols')
const autoAssignBtn = document.querySelector('#auto-assign')
const seatResetDisplayBtn = document.querySelector('#seat-reset-display')
const effectToggleInput = document.querySelector('#effect-toggle')
const countdownOverlay = document.querySelector('#countdown-overlay')
const countdownNumberEl = document.querySelector('#countdown-number')
const shuffleBtn = document.querySelector('#shuffle')
const advancedControlsEl = document.querySelector('#advanced-controls')
const secretToggleBtn = document.querySelector('#secret-toggle')

const separateInput = document.querySelector('#separate-input')
const preassignedListEl = document.querySelector('#preassigned-list')

function parseStudents(rawText) {
  return rawText
    .split(/[\n,]+/)
    .map((name) => name.trim())
    .filter(Boolean)
}

function setStudentsTextarea(students) {
  studentInput.value = students.join(', ')
  refreshPresetStudentSelect()
}

const STORAGE_PREFIX_V2 = 'seat123_students_v2:'
const STORAGE_LAST_GROUP_V2 = 'seat123_students_last_group_v2'
const STORAGE_KEY_V1 = 'seat123_students_v1'

function getGroupFromUI() {
  const groupFromSelect = savedGroupsSelect?.value
  const groupFromInput = groupNameInput?.value
  const group = (groupFromSelect || groupFromInput || '').trim()
  return group
}

function refreshSavedGroups() {
  if (!savedGroupsSelect) return
  const currentValue = savedGroupsSelect.value
  const groups = []
  for (let i = 0; i < localStorage.length; i += 1) {
    const key = localStorage.key(i)
    if (key && key.startsWith(STORAGE_PREFIX_V2)) {
      groups.push(key.slice(STORAGE_PREFIX_V2.length))
    }
  }

  groups.sort((a, b) => a.localeCompare(b, 'ko'))
  savedGroupsSelect.innerHTML = '<option value="">저장된 그룹 선택</option>'
  for (const g of groups) {
    const opt = document.createElement('option')
    opt.value = g
    opt.textContent = g
    savedGroupsSelect.appendChild(opt)
  }

  // 마지막에 저장한 그룹이 있으면 자동 선택
  const last = localStorage.getItem(STORAGE_LAST_GROUP_V2)
  if (last && groups.includes(last)) {
    savedGroupsSelect.value = last
  } else if (groups.includes(currentValue)) {
    savedGroupsSelect.value = currentValue
  }
}

function applyViewPerspective() {
  if (!seatBoardEl || !viewPerspectiveToggleBtn) return
  const isStudent = state.viewPerspective === 'student'
  seatBoardEl.classList.toggle('perspective-teacher', !isStudent)
  seatBoardEl.classList.toggle('perspective-student', isStudent)
  viewPerspectiveToggleBtn.textContent = isStudent ? '교사뷰' : '학생뷰'
  viewPerspectiveToggleBtn.setAttribute(
    'aria-label',
    isStudent ? '교사 뷰로 자리표 보기' : '학생 뷰로 자리표 보기'
  )
}

/** 저장된 객체에서 사전 배치 복원. 학생은 명단에 있고 좌석 id가 현재 판에 있어야 반영. */
function applyPreAssignmentsFromSavedObject(entries, studentsList) {
  state.preAssignments.clear()
  const studentSet = new Set(studentsList)
  const seatById = new Map(state.seats.map((s) => [s.id, s]))

  if (!entries || typeof entries !== 'object' || Array.isArray(entries)) {
    state.presetApplied = false
    return { applied: 0, dropped: 0 }
  }

  const pairs = Object.entries(entries)
    .map(([seatId, st]) => [seatId, String(st).trim()])
    .filter(([, name]) => name)

  const usedStudents = new Set()
  let applied = 0
  let dropped = 0

  for (const [seatId, name] of pairs) {
    if (!seatById.has(seatId) || !studentSet.has(name)) {
      dropped += 1
      continue
    }
    if (usedStudents.has(name)) {
      dropped += 1
      continue
    }
    usedStudents.add(name)
    state.preAssignments.set(seatId, name)
    state.fixedAssignments.delete(seatId)
    const seat = seatById.get(seatId)
    if (seat) seat.student = name
    applied += 1
  }
  state.presetApplied = false
  return { applied, dropped }
}

function saveStudentsToLocal() {
  const students = parseStudents(studentInput.value)
  if (students.length === 0) {
    updateStatus('저장할 명단이 비어 있습니다.')
    return
  }

  const groupName = (groupNameInput?.value || '').trim()
  if (!groupName) {
    updateStatus('그룹 이름을 입력해 주세요.')
    return
  }

  const preAssignmentsObj = Object.fromEntries(state.preAssignments)
  const payload = {
    students,
    savedAt: Date.now(),
    group: groupName,
    preAssignments: preAssignmentsObj,
    seatLayout: getSeatLayout(),
  }
  try {
    localStorage.setItem(`${STORAGE_PREFIX_V2}${groupName}`, JSON.stringify(payload))
    localStorage.setItem(STORAGE_LAST_GROUP_V2, groupName)
    refreshSavedGroups()
    const preN = state.preAssignments.size
    updateStatus(
      preN > 0
        ? `명단·사전 배치(${preN}건)가 저장되었습니다. (그룹: ${groupName})`
        : `명단이 저장되었습니다. (그룹: ${groupName})`
    )
  } catch {
    updateStatus('명단 저장에 실패했습니다. 브라우저 저장 공간을 확인해 주세요.')
  }
}

function deleteSavedStudentsFromLocal() {
  const groupName = getGroupFromUI()
  if (!groupName) {
    updateStatus('삭제할 그룹을 드롭다운에서 선택하거나 그룹 이름을 입력해 주세요.')
    return
  }

  const key = `${STORAGE_PREFIX_V2}${groupName}`
  if (!localStorage.getItem(key)) {
    updateStatus(`저장된 명단을 찾을 수 없습니다. (그룹: ${groupName})`)
    return
  }

  if (!window.confirm(`「${groupName}」명단을 브라우저에서 삭제할까요?`)) {
    return
  }

  localStorage.removeItem(key)
  if (localStorage.getItem(STORAGE_LAST_GROUP_V2) === groupName) {
    localStorage.removeItem(STORAGE_LAST_GROUP_V2)
  }
  refreshSavedGroups()
  updateStatus(`명단을 삭제했습니다. (그룹: ${groupName})`)
}

function loadStudentsFromLocal() {
  const groupName = getGroupFromUI()
  if (!groupName) {
    // v1(기존 단일 키)이 있으면 마지막 수단으로 불러오기
    const rawV1 = localStorage.getItem(STORAGE_KEY_V1)
    if (rawV1) {
      try {
        const parsedV1 = JSON.parse(rawV1)
        const students = Array.isArray(parsedV1?.students)
          ? parsedV1.students.map((x) => String(x).trim()).filter(Boolean)
          : []
        if (students.length === 0) {
          updateStatus('저장된 명단이 비어 있습니다.')
          return
        }
        setStudentsTextarea(students)
        state.students = students
        applyPreAssignmentsFromSavedObject(null, students)
        refreshPresetStudentSelect()
        updateStatus('저장된 명단을 불러왔습니다. (기존 저장 방식)')
        return
      } catch {
        // 아래 메시지로 처리
      }
    }
    updateStatus('불러올 그룹을 선택해 주세요.')
    return
  }

  const rawV2 = localStorage.getItem(`${STORAGE_PREFIX_V2}${groupName}`)
  if (!rawV2) {
    updateStatus(`그룹을 찾을 수 없습니다. (그룹: ${groupName})`)
    return
  }

  let parsed
  try {
    parsed = JSON.parse(rawV2)
  } catch {
    updateStatus('저장된 명단 형식이 올바르지 않습니다.')
    return
  }

  const students = Array.isArray(parsed?.students)
    ? parsed.students.map((x) => String(x).trim()).filter(Boolean)
    : []

  if (students.length === 0) {
    updateStatus('저장된 명단이 비어 있습니다.')
    return
  }

  setStudentsTextarea(students)
  state.students = students
  applySeatLayoutFromSaved(parsed)
  const preResult = applyPreAssignmentsFromSavedObject(parsed.preAssignments, students)
  refreshPresetStudentSelect()
  renderSeats()
  renderPreassignedList()
  let msg = `그룹의 명단을 불러왔습니다. (그룹: ${groupName})`
  if (preResult.applied > 0) msg += ` 사전 배치 ${preResult.applied}건 복원.`
  if (preResult.dropped > 0) msg += ` (생략 ${preResult.dropped}건)`
  updateStatus(msg)
}

function makeSeats(rows, cols) {
  const list = []
  for (let r = 1; r <= rows; r += 1) {
    for (let c = 1; c <= cols; c += 1) {
      // 행 기준(위→아래), 각 행 내에서 왼쪽→오른쪽 방향으로 1부터 증가
      // (열 우선) 왼쪽 열부터 위→아래로 번호가 증가하고, 다음 열로 넘어갑니다.
      const index = (r - 1) * cols + c
      list.push({ id: `${r}-${c}`, row: r, col: c, index, student: '' })
    }
  }
  return list
}

function parseSeparatedPairs(rawText) {
  return rawText
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => line.split(/[-,]/).map((name) => name.trim()).filter(Boolean))
    .filter((pair) => pair.length === 2 && pair[0] !== pair[1])
}

function shuffleArray(items) {
  const arr = [...items]
  for (let i = arr.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1))
    ;[arr[i], arr[j]] = [arr[j], arr[i]]
  }
  return arr
}

const SEAT_LAYOUTS = ['individual', 'pair', 'group']

function getSeatLayout() {
  const v = seatLayoutSelect?.value
  return SEAT_LAYOUTS.includes(v) ? v : 'individual'
}

function syncSeatDimensionLabels() {
  const g = getSeatLayout() === 'group'
  if (labelRows) labelRows.textContent = g ? '모둠 수' : '행(줄) 수'
  if (labelCols) labelCols.textContent = g ? '모둠당 인원' : '열(칸) 수'
  if (rowsInput) {
    rowsInput.title = g ? '모둠(팀) 개수' : '교실 앞에서 뒤로 가는 줄 수'
    rowsInput.max = g ? '24' : '10'
  }
  if (colsInput) {
    colsInput.title = g ? '한 모둠에 앉을 인원 수' : '한 줄의 좌석 칸 수'
    colsInput.max = g ? '30' : '10'
  }
  const maxR = rowsInput ? Number(rowsInput.max) : 10
  const maxC = colsInput ? Number(colsInput.max) : 10
  if (rowsInput) {
    const v = Number(rowsInput.value)
    if (v > maxR) rowsInput.value = String(maxR)
    if (v < 1 || Number.isNaN(v)) rowsInput.value = '1'
  }
  if (colsInput) {
    const v = Number(colsInput.value)
    if (v > maxC) colsInput.value = String(maxC)
    if (v < 1 || Number.isNaN(v)) colsInput.value = '1'
  }
}

function applySeatLayoutFromSaved(parsed) {
  const v = parsed?.seatLayout
  if (seatLayoutSelect && SEAT_LAYOUTS.includes(v)) {
    seatLayoutSelect.value = v
  }
  syncSeatDimensionLabels()
}

/** 모둠 안 책상 배치: 가로 2열(인원만큼 세로로 쌓음) */
function clusterGridDims(n) {
  const count = Math.max(1, Math.floor(Number(n)) || 1)
  if (count === 1) return { gc: 1, gr: 1 }
  const gc = 2
  const gr = Math.ceil(count / gc)
  return { gc, gr }
}

/** 모둠 박스 바깥 배치: 3열(모둠 1개만 있을 때만 1열) */
function outerBundanGridCols(teamCount) {
  const t = Math.max(1, Math.floor(Number(teamCount)) || 1)
  if (t <= 1) return 1
  return 3
}

function createSeatButton(seat) {
  const el = document.createElement('button')
  el.type = 'button'
  el.className = 'seat'
  el.dataset.seatId = seat.id

  const displayName = (seat.student || '').trim()

  if (state.blockedSeats.has(seat.id)) {
    el.classList.add('blocked')
  } else if (
    displayName &&
    state.fixedAssignments.has(seat.id) &&
    !state.preAssignments.has(seat.id)
  ) {
    el.classList.add('fixed')
  } else if (displayName) {
    el.classList.add('filled')
  } else {
    el.classList.add('empty')
  }

  if (state.blockedSeats.has(seat.id)) {
    el.innerHTML = `<span class="pos">${seat.index}</span><span class="blocked-icon" aria-hidden="true">X</span>`
  } else {
    el.innerHTML = `<span class="pos">${seat.index}</span><span class="name">${displayName}</span>`
  }
  return el
}

function renderSeats() {
  seatGrid.innerHTML = ''
  const cols = Number(colsInput.value)
  const rows = Number(rowsInput.value)
  const layout = getSeatLayout()

  seatGrid.className = `seat-grid layout-${layout}`

  if (layout === 'individual') {
    seatGrid.style.gridTemplateColumns = `repeat(${cols}, minmax(90px, 1fr))`
    for (const seat of state.seats) {
      seatGrid.appendChild(createSeatButton(seat))
    }
  } else if (layout === 'pair') {
    seatGrid.style.gridTemplateColumns = ''
    for (let r = 1; r <= rows; r += 1) {
      const rowEl = document.createElement('div')
      rowEl.className = 'seat-row-pair'
      for (let c = 1; c <= cols; c += 2) {
        const pairEl = document.createElement('div')
        pairEl.className = 'seat-pair'
        const s1 = state.seats.find((s) => s.row === r && s.col === c)
        const s2 = state.seats.find((s) => s.row === r && s.col === c + 1)
        if (s1) pairEl.appendChild(createSeatButton(s1))
        if (s2) pairEl.appendChild(createSeatButton(s2))
        rowEl.appendChild(pairEl)
      }
      seatGrid.appendChild(rowEl)
    }
  } else {
    // 모둠: 행 = 분단 번호, 열 = 분단 내 자리. 분단별 박스 + 안쪽은 책상 덩어리(클러스터) 배치.
    const outerCols = outerBundanGridCols(rows)
    seatGrid.style.gridTemplateColumns = `repeat(${outerCols}, minmax(220px, 1fr))`
    const { gc } = clusterGridDims(cols)
    for (let r = 1; r <= rows; r += 1) {
      const groupEl = document.createElement('div')
      groupEl.className = 'seat-group'

      const title = document.createElement('div')
      title.className = 'seat-group-title'
      title.textContent = `${r}모둠`

      const desks = document.createElement('div')
      desks.className = 'seat-group-desks'
      desks.style.setProperty('--cluster-cols', String(gc))

      for (let c = 1; c <= cols; c += 1) {
        const seat = state.seats.find((s) => s.row === r && s.col === c)
        if (seat) desks.appendChild(createSeatButton(seat))
      }
      groupEl.appendChild(title)
      groupEl.appendChild(desks)
      seatGrid.appendChild(groupEl)
    }
  }

  applyViewPerspective()
}

function renderPreassignedList() {
  if (!preassignedListEl) return

  const entries = Array.from(state.preAssignments.entries())
  entries.sort((a, b) => {
    const seatA = state.seats.find((s) => s.id === a[0])
    const seatB = state.seats.find((s) => s.id === b[0])
    const idxA = seatA?.index ?? 0
    const idxB = seatB?.index ?? 0
    return idxA - idxB
  })

  if (entries.length === 0) {
    preassignedListEl.innerHTML = '<p class="muted">사전 배치가 없습니다.</p>'
    return
  }

  preassignedListEl.innerHTML = entries
    .map(([seatId, student]) => {
      const seat = state.seats.find((s) => s.id === seatId)
      const seatIndex = seat?.index ?? seatId
      const seatPos = seat ? `(${seat.row}, ${seat.col})` : ''
      const planned = student
      return `<div class="preassigned-row"><div class="main">좌석 ${seatIndex} ${seatPos}</div><div class="sub">설정: ${planned}</div></div>`
    })
    .join('')
}

function updateStatus(message = '') {
  const assigned = state.seats.filter((seat) => seat.student).length
  const blocked = state.blockedSeats.size
  const total = state.seats.length
  const base = `총 ${total}석 / 배치 ${assigned}명 / 제외 ${blocked}석`
  statusText.textContent = message ? `${base} - ${message}` : base
}

function seatCellExportText(seat) {
  if (state.blockedSeats.has(seat.id)) return ''
  if (seat.student) return seat.student
  const pre = state.preAssignments.get(seat.id)
  if (pre) return pre
  return ''
}

function exportSeatChartToExcel() {
  if (!state.seats.length) {
    updateStatus('먼저 좌석판을 만들어 주세요.')
    return
  }
  const rows = Number(rowsInput.value)
  const cols = Number(colsInput.value)
  if (!rows || !cols) {
    updateStatus('행·열 정보를 확인할 수 없습니다.')
    return
  }

  const byId = new Map(state.seats.map((s) => [s.id, s]))
  const aoa = []
  const headerRow = Array(cols).fill('')
  headerRow[0] = '교탁'
  aoa.push(headerRow)

  for (let r = 1; r <= rows; r += 1) {
    const row = []
    for (let c = 1; c <= cols; c += 1) {
      const seat = byId.get(`${r}-${c}`)
      row.push(seat ? seatCellExportText(seat) : '')
    }
    aoa.push(row)
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa)
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: cols - 1 } }]

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, '좌석배치')

  const group = (groupNameInput?.value || savedGroupsSelect?.value || '').trim()
  const d = new Date()
  const dateStr = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`
  const safeGroup = group.replace(/[\\/:*?"<>|]/g, '_').slice(0, 40)
  const base = safeGroup ? `좌석배치_${safeGroup}_${dateStr}` : `좌석배치_${dateStr}`
  XLSX.writeFile(wb, `${base}.xlsx`)
  updateStatus('엑셀 파일을 저장했습니다.')
}

function buildSeatMap() {
  const rows = Number(rowsInput.value)
  const cols = Number(colsInput.value)

  if (!rows || !cols) {
    updateStatus('행과 열을 올바르게 입력해 주세요.')
    return
  }

  state.seats = makeSeats(rows, cols)
  state.students = parseStudents(studentInput.value)
  state.fixedAssignments.clear()
  state.preAssignments.clear()
  state.presetApplied = false
  state.blockedSeats.clear()

  refreshPresetStudentSelect()
  renderSeats()
  renderPreassignedList()
  updateStatus('좌석판이 생성되었습니다.')
}

function tryRestoreLastSavedGroup() {
  const groupName = (localStorage.getItem(STORAGE_LAST_GROUP_V2) || '').trim()
  if (!groupName) return
  const raw = localStorage.getItem(`${STORAGE_PREFIX_V2}${groupName}`)
  if (!raw) return
  let parsed
  try {
    parsed = JSON.parse(raw)
  } catch {
    return
  }
  const students = Array.isArray(parsed?.students)
    ? parsed.students.map((x) => String(x).trim()).filter(Boolean)
    : []
  if (students.length === 0) return

  setStudentsTextarea(students)
  state.students = students
  applySeatLayoutFromSaved(parsed)
  applyPreAssignmentsFromSavedObject(parsed.preAssignments, students)
  if (groupNameInput) groupNameInput.value = groupName
  refreshSavedGroups()
  if (savedGroupsSelect) savedGroupsSelect.value = groupName
  refreshPresetStudentSelect()
  renderSeats()
  renderPreassignedList()
  updateStatus(`마지막 저장 그룹을 불러왔습니다. (그룹: ${groupName})`)
}

function refreshPresetStudentSelect() {
  const selectedValue = presetStudentSelect.value
  const names = parseStudents(studentInput.value)
  presetStudentSelect.innerHTML = '<option value="">선택 안 함 (일반 모드)</option>'
  for (const name of names) {
    const opt = document.createElement('option')
    opt.value = name
    opt.textContent = name
    presetStudentSelect.appendChild(opt)
  }
  if (names.includes(selectedValue)) {
    presetStudentSelect.value = selectedValue
  } else {
    presetStudentSelect.value = ''
  }
}

function isNeighbor(seatA, seatB) {
  if (!seatA || !seatB) return false
  return Math.abs(seatA.row - seatB.row) <= 1 && Math.abs(seatA.col - seatB.col) <= 1
}

function satisfiesSeparationPairs(separatedPairs) {
  const studentToSeat = new Map()
  for (const seat of state.seats) {
    if (seat.student) {
      studentToSeat.set(seat.student, seat)
    }
  }

  for (const [first, second] of separatedPairs) {
    const aSeat = studentToSeat.get(first)
    const bSeat = studentToSeat.get(second)
    if (isNeighbor(aSeat, bSeat)) {
      return false
    }
  }
  return true
}

function applyPresetAssignments(students) {
  const studentSet = new Set(students)
  const usedSeat = new Set()
  const usedStudent = new Set()

  for (const [seatId, student] of state.preAssignments.entries()) {
    const seat = state.seats.find((item) => item.id === seatId)
    if (!seat) {
      return `사전 배치 좌표 오류: ${seatId}`
    }
    if (state.blockedSeats.has(seatId)) {
      return `사전 배치 좌석이 제외되어 있습니다: ${student} -> ${seatId}`
    }
    if (!studentSet.has(student)) {
      return `명단에 없는 학생이 사전 배치에 있습니다: ${student}`
    }
    if (usedSeat.has(seatId) || usedStudent.has(student)) {
      return `사전 배치 중복이 있습니다: ${student}`
    }
    usedSeat.add(seatId)
    usedStudent.add(student)
    state.fixedAssignments.set(seatId, student)
  }

  return ''
}

function autoAssign(applyPreset = false) {
  if (!state.seats.length) {
    updateStatus('먼저 좌석판을 만들어 주세요.')
    return
  }
  state.students = parseStudents(studentInput.value)
  const separatedPairs = parseSeparatedPairs(separateInput.value)

  const preSeatIds = Array.from(state.preAssignments.keys())
  for (const seatId of preSeatIds) {
    state.fixedAssignments.delete(seatId)
  }

  if (applyPreset) {
    const presetError = applyPresetAssignments(state.students)
    if (presetError) {
      updateStatus(presetError)
      renderSeats()
      renderPreassignedList()
      return
    }
    state.presetApplied = true
  } else {
    // Ctrl 없이 start: 사전 배치는 반영하지 않되, Shift 고정은 유지
    state.presetApplied = false
  }

  for (const seat of state.seats) {
    if (!state.fixedAssignments.has(seat.id)) {
      seat.student = ''
    } else {
      seat.student = state.fixedAssignments.get(seat.id) || ''
    }
  }

  const availableSeats = state.seats.filter(
    (seat) => !state.blockedSeats.has(seat.id) && !state.fixedAssignments.has(seat.id)
  )

  const fixedStudents = new Set(state.fixedAssignments.values())
  const remainingStudents = state.students.filter((name) => !fixedStudents.has(name))
  const assignableCount = Math.min(availableSeats.length, remainingStudents.length)
  let success = false
  const maxAttempts = 600

  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const randomized = shuffleArray(remainingStudents)
    availableSeats.forEach((seat, idx) => {
      seat.student = idx < assignableCount ? randomized[idx] : ''
    })
    if (satisfiesSeparationPairs(separatedPairs)) {
      success = true
      break
    }
  }

  if (!success && separatedPairs.length > 0) {
    updateStatus('분리 조건을 만족하는 배치를 찾지 못했습니다. 조건을 완화해 주세요.')
    renderSeats()
    renderPreassignedList()
    return
  }

  if (remainingStudents.length > availableSeats.length) {
    updateStatus('좌석이 부족하여 일부 학생은 배치되지 않았습니다.')
  } else if (applyPreset) {
    updateStatus('자동 배치가 완료되었습니다. (사전 배치 반영)')
  } else {
    updateStatus(
      '전체 무작위 배치가 완료되었습니다. (사전 배치 목록은 그대로 두었습니다. 반영은 Ctrl+start)'
    )
  }
  renderSeats()
  renderPreassignedList()
}

function resetSeatDisplay() {
  // 명단은 유지하고 좌석 상태만 초기화
  state.students = parseStudents(studentInput.value)
  state.fixedAssignments.clear()
  state.preAssignments.clear()
  state.presetApplied = false
  state.blockedSeats.clear()
  for (const seat of state.seats) {
    seat.student = ''
  }
  refreshPresetStudentSelect()
  renderSeats()
  renderPreassignedList()
  updateStatus('배치/제외/고정을 초기화했습니다. 명단은 유지됩니다.')
}

seatGrid.addEventListener('click', (event) => {
  const target = event.target.closest('.seat')
  if (!target) return

  const seatId = target.dataset.seatId
  const seat = state.seats.find((item) => item.id === seatId)
  if (!seat) return

  const selectedPresetStudent = presetStudentSelect.value
  if (selectedPresetStudent && !event.shiftKey) {
    if (state.blockedSeats.has(seatId)) {
      updateStatus('제외 좌석에는 사전 배치를 할 수 없습니다.')
      return
    }

    for (const [existingSeatId, existingStudent] of state.preAssignments.entries()) {
      if (existingStudent === selectedPresetStudent && existingSeatId !== seatId) {
        state.preAssignments.delete(existingSeatId)
        state.fixedAssignments.delete(existingSeatId)
        const oldSeat = state.seats.find((item) => item.id === existingSeatId)
        if (oldSeat) oldSeat.student = ''
      }
    }

    if (state.preAssignments.get(seatId) === selectedPresetStudent) {
      state.preAssignments.delete(seatId)
      seat.student = ''
      state.fixedAssignments.delete(seatId)
      state.presetApplied = false
      renderSeats()
      renderPreassignedList()
      updateStatus(`${selectedPresetStudent} 사전 배치를 해제했습니다.`)
      return
    }

    state.preAssignments.set(seatId, selectedPresetStudent)
    state.fixedAssignments.delete(seatId)
    seat.student = selectedPresetStudent
    state.presetApplied = false
    renderSeats()
    renderPreassignedList()
    window.alert(`${selectedPresetStudent} 학생의 사전 배치가 완료되었습니다.`)
    updateStatus(`${selectedPresetStudent} 학생을 (${seat.row}, ${seat.col})에 사전 배치했습니다.`)
    return
  }

  if (event.shiftKey) {
    // Shift+클릭은 "학생 고정/해제"로 동작
    if (!seat.student) {
      updateStatus('고정/해제는 학생이 배치된 좌석에서만 가능합니다.')
      return
    }
    if (state.blockedSeats.has(seatId)) {
      // 고정 시에는 제외 해제하는 편이 자연스러움
      state.blockedSeats.delete(seatId)
    }
    if (state.fixedAssignments.has(seatId)) {
      state.fixedAssignments.delete(seatId)
    } else {
      // 사전 배치 좌석을 Shift로 고정하면 "공개 고정"으로 전환
      // (사전 배치 목록에서는 제거되고, 일반 고정(초록)으로 표시)
      state.preAssignments.delete(seatId)
      state.fixedAssignments.set(seatId, seat.student)
    }
  } else {
    // 일반 클릭은 "제외(X) 토글"로 동작 (학생이 있어도 비워둠)
    if (state.blockedSeats.has(seatId)) {
      state.blockedSeats.delete(seatId)
    } else {
      state.blockedSeats.add(seatId)
    }
    state.fixedAssignments.delete(seatId)
    state.preAssignments.delete(seatId)
    seat.student = ''
  }

  renderSeats()
  renderPreassignedList()
  updateStatus()
})

function runCountdownThen(callback) {
  if (!countdownOverlay || !countdownNumberEl) {
    callback()
    return
  }
  let n = 5
  const finish = () => {
    countdownOverlay.classList.remove('show')
    countdownOverlay.setAttribute('aria-hidden', 'true')
    autoAssignBtn.disabled = false
    try {
      callback()
    } catch {
      /* 배치 중 오류가 나도 버튼은 복구 */
    }
  }
  const tick = () => {
    countdownNumberEl.textContent = String(n)
    if (n === 1) {
      setTimeout(finish, 1000)
      return
    }
    n -= 1
    setTimeout(tick, 1000)
  }
  autoAssignBtn.disabled = true
  countdownOverlay.classList.add('show')
  countdownOverlay.setAttribute('aria-hidden', 'false')
  tick()
}

buildSeatMapBtn.addEventListener('click', buildSeatMap)
autoAssignBtn.addEventListener('click', (event) => {
  const applyPreset = Boolean(event.ctrlKey)
  const run = () => autoAssign(applyPreset)
  if (effectToggleInput?.checked) {
    runCountdownThen(run)
  } else {
    run()
  }
})
if (shuffleBtn) {
  shuffleBtn.addEventListener('click', () => {
    autoAssign(false)
  })
}
seatResetDisplayBtn?.addEventListener('click', resetSeatDisplay)
viewPerspectiveToggleBtn?.addEventListener('click', () => {
  state.viewPerspective = state.viewPerspective === 'teacher' ? 'student' : 'teacher'
  applyViewPerspective()
})
exportSeatExcelBtn?.addEventListener('click', exportSeatChartToExcel)
studentInput.addEventListener('input', refreshPresetStudentSelect)
saveStudentsBtn.addEventListener('click', saveStudentsToLocal)
loadStudentsBtn.addEventListener('click', loadStudentsFromLocal)
deleteSavedStudentsBtn?.addEventListener('click', deleteSavedStudentsFromLocal)
rowsInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') buildSeatMap()
})
colsInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') buildSeatMap()
})
secretToggleBtn.addEventListener('click', () => {
  advancedControlsEl.classList.toggle('show')
})
seatLayoutSelect?.addEventListener('change', () => {
  syncSeatDimensionLabels()
  if (state.seats.length) renderSeats()
})

clearPreassignmentsBtn.addEventListener('click', () => {
  // 사전 배치 목록은 유지하고, 자리표에서는 이름만 숨김
  for (const [seatId] of state.preAssignments.entries()) {
    state.fixedAssignments.delete(seatId)
    const seat = state.seats.find((item) => item.id === seatId)
    if (seat) seat.student = ''
  }
  state.presetApplied = false
  if (presetStudentSelect) {
    presetStudentSelect.value = ''
  }
  renderSeats()
  renderPreassignedList()
  updateStatus('사전 배치가 완료되었습니다. 자리표에서는 이름이 숨겨집니다.')
})

refreshSavedGroups()
syncSeatDimensionLabels()
buildSeatMap()
tryRestoreLastSavedGroup()