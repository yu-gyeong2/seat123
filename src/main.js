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
  groupLeaders: new Set(),
  /** 'teacher' = 학생이 교실 뒤에서 보는 배치, 'student' = 교탁이 위(교실 앞) */
  viewPerspective: 'teacher',
}

const app = document.querySelector('#app')
app.innerHTML = `
  <main class="container">
    <header class="header">
      <h1>❤️ 학교 자리 배치 프로그램</h1>
      <p>학생 명단을 입력하고 좌석을 자동으로 배치하세요.</p>
      <p class="header-contact">
        계속 업데이트 중. 건의 사항 있으면 연락 주세요 :)
        <a href="mailto:yg.tech602@gmail.com">yg.tech602@gmail.com</a>
      </p>
      <p class="header-credit">Made by 유경T</p>
    </header>

    <section class="panel controls">
      <div class="field-group wide seat-layout-field">
        <label for="seat-layout">자리 형태</label>
        <select id="seat-layout">
          <option value="individual">개별</option>
          <option value="pair">짝꿍 (옆자리 2명씩 붙음)</option>
          <option value="group">모둠 (모둠 수 × 모둠당 인원)</option>
          <option value="group_diverse">모둠 (학생 특성 분류)</option>
        </select>
      </div>
      <div class="seat-setup-row">
        <div class="field-group">
          <label id="label-rows" for="rows">행(줄) 수</label>
          <input id="rows" type="number" min="1" max="10" value="6" />
        </div>
        <div class="field-group">
          <label id="label-cols" for="cols">열(칸) 수</label>
          <input id="cols" type="number" min="1" max="10" value="5" />
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
      <div id="trait-buckets-wrap" class="field-group wide trait-buckets-wrap" hidden>
        <p class="secret-title">▶️ 특성별 명단 입력 (줄바꿈/쉼표 구분)</p>
        <div class="trait-buckets">
          <div class="field-group">
            <label for="trait-bucket-1">특성 1</label>
            <textarea id="trait-bucket-1" rows="3" placeholder="김건호, 김도연"></textarea>
          </div>
          <div class="field-group">
            <label for="trait-bucket-2">특성 2</label>
            <textarea id="trait-bucket-2" rows="3" placeholder="강대현, 홍길동"></textarea>
          </div>
          <div class="field-group">
            <label for="trait-bucket-3">특성 3</label>
            <textarea id="trait-bucket-3" rows="3" placeholder="학생 이름"></textarea>
          </div>
          <div class="field-group">
            <label for="trait-bucket-4">특성 4</label>
            <textarea id="trait-bucket-4" rows="3" placeholder="학생 이름"></textarea>
          </div>
        </div>
      </div>
      <div class="field-group wide">
        <div class="row-actions">
          <input id="group-name" type="text" placeholder="그룹 이름 (예: 2-7)" />
          <button id="save-students" type="button">명단 저장</button>
          <div class="row-actions-load">
            <div class="saved-groups-widget">
              <input type="hidden" id="saved-group-key" value="" />
              <button
                type="button"
                id="saved-groups-trigger"
                class="saved-groups-trigger"
                aria-haspopup="listbox"
                aria-expanded="false"
              >
                <span id="saved-groups-display">저장된 그룹 선택</span>
                <span class="saved-groups-chevron" aria-hidden="true">▾</span>
              </button>
              <div id="saved-groups-popover" class="saved-groups-popover" hidden>
                <ul id="saved-groups-list" class="saved-groups-list" role="listbox"></ul>
                <p id="saved-groups-empty" class="saved-groups-empty muted" hidden>저장된 그룹이 없습니다.</p>
              </div>
            </div>
            <button id="load-students" type="button">명단 불러오기</button>
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
            <li>사전 좌석 배치: 학생 선택 후 좌석 클릭 → 배치할 학생들 모두 배치 후 「사전 배치 완료」 클릭</li>
            <li>사전 배치 반영은 <strong class="ctrl-key-hint">Ctrl</strong> 키를 누른 채 자리 배치 클릭</li>
            <li>특정 학생 공개 고정: 사전배치 후 <strong class="shift-key-hint">Shift</strong>+좌석클릭(초록색으로 변함)</li>
            <li>학생 분리: 분리할 학생 쌍에 입력(랜덤하게 떨어진 채로 배치됨)</li>
            <li>모둠장 지정👑: 모둠 자리 배치 후 모둠장으로 정해진 학생 클릭</li>
            <li>⭐사전 배치 및 분리 학생 저장을 원할 경우 「명단 저장」 한번 더 클릭</li>
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
      <div id="seat-board" class="seat-board perspective-student">
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

    <div
      id="delete-group-modal"
      class="delete-group-modal"
      hidden
      aria-modal="true"
      role="dialog"
      aria-labelledby="delete-group-modal-title"
    >
      <div class="delete-group-modal-backdrop" data-modal-dismiss="1" aria-hidden="true"></div>
      <div class="delete-group-modal-panel">
        <h2 id="delete-group-modal-title" class="delete-group-modal-title">명단 삭제</h2>
        <p id="delete-group-modal-text" class="delete-group-modal-text"></p>
        <div class="delete-group-modal-actions">
          <button type="button" id="delete-group-cancel" class="delete-group-btn-cancel">취소</button>
          <button type="button" id="delete-group-confirm" class="delete-group-btn-confirm">확인</button>
        </div>
      </div>
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
const savedGroupKeyInput = document.querySelector('#saved-group-key')
const savedGroupsWidget = document.querySelector('.saved-groups-widget')
const savedGroupsTrigger = document.querySelector('#saved-groups-trigger')
const savedGroupsDisplay = document.querySelector('#saved-groups-display')
const savedGroupsPopover = document.querySelector('#saved-groups-popover')
const savedGroupsListEl = document.querySelector('#saved-groups-list')
const savedGroupsEmptyEl = document.querySelector('#saved-groups-empty')
const deleteGroupModal = document.querySelector('#delete-group-modal')
const deleteGroupModalText = document.querySelector('#delete-group-modal-text')
const deleteGroupConfirmBtn = document.querySelector('#delete-group-confirm')
const deleteGroupCancelBtn = document.querySelector('#delete-group-cancel')

let pendingDeleteGroupName = ''
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
const traitBucketsWrapEl = document.querySelector('#trait-buckets-wrap')

const separateInput = document.querySelector('#separate-input')
const traitBucket1Input = document.querySelector('#trait-bucket-1')
const traitBucket2Input = document.querySelector('#trait-bucket-2')
const traitBucket3Input = document.querySelector('#trait-bucket-3')
const traitBucket4Input = document.querySelector('#trait-bucket-4')
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
  const groupFromKey = savedGroupKeyInput?.value
  const groupFromInput = groupNameInput?.value
  const group = (groupFromKey || groupFromInput || '').trim()
  return group
}

function setSavedGroupSelection(name) {
  const v = (name || '').trim()
  if (savedGroupKeyInput) savedGroupKeyInput.value = v
  if (savedGroupsDisplay) {
    savedGroupsDisplay.textContent = v || '저장된 그룹 선택'
  }
}

function closeSavedGroupsPopover() {
  if (!savedGroupsPopover || !savedGroupsTrigger) return
  savedGroupsPopover.hidden = true
  savedGroupsTrigger.setAttribute('aria-expanded', 'false')
}

function openSavedGroupsPopover() {
  if (!savedGroupsPopover || !savedGroupsTrigger) return
  savedGroupsPopover.hidden = false
  savedGroupsTrigger.setAttribute('aria-expanded', 'true')
}

function toggleSavedGroupsPopover() {
  if (!savedGroupsPopover) return
  if (savedGroupsPopover.hidden) openSavedGroupsPopover()
  else closeSavedGroupsPopover()
}

function closeDeleteGroupModal() {
  pendingDeleteGroupName = ''
  if (deleteGroupModal) deleteGroupModal.hidden = true
}

function openDeleteGroupModal(groupName) {
  const g = (groupName || '').trim()
  if (!g || !deleteGroupModal || !deleteGroupModalText) return
  pendingDeleteGroupName = g
  deleteGroupModalText.textContent = `「${g}」명단을 브라우저에서 삭제할까요?`
  deleteGroupModal.hidden = false
  closeSavedGroupsPopover()
  deleteGroupCancelBtn?.focus()
}

function performDeleteSavedGroup(groupName) {
  const g = (groupName || '').trim()
  if (!g) return
  const key = `${STORAGE_PREFIX_V2}${g}`
  if (!localStorage.getItem(key)) {
    updateStatus(`저장된 명단을 찾을 수 없습니다. (그룹: ${g})`)
    closeDeleteGroupModal()
    return
  }
  localStorage.removeItem(key)
  if (localStorage.getItem(STORAGE_LAST_GROUP_V2) === g) {
    localStorage.removeItem(STORAGE_LAST_GROUP_V2)
  }
  refreshSavedGroups()
  closeDeleteGroupModal()
  updateStatus(`명단을 삭제했습니다. (그룹: ${g})`)
}

function refreshSavedGroups() {
  const currentValue = (savedGroupKeyInput?.value || '').trim()
  const groups = []
  for (let i = 0; i < localStorage.length; i += 1) {
    const key = localStorage.key(i)
    if (key && key.startsWith(STORAGE_PREFIX_V2)) {
      groups.push(key.slice(STORAGE_PREFIX_V2.length))
    }
  }

  groups.sort((a, b) => a.localeCompare(b, 'ko'))

  if (savedGroupsListEl) {
    savedGroupsListEl.innerHTML = ''
    for (const g of groups) {
      const li = document.createElement('li')
      li.className = 'saved-groups-row'
      li.setAttribute('role', 'presentation')

      const pick = document.createElement('button')
      pick.type = 'button'
      pick.className = 'saved-groups-pick'
      pick.textContent = g
      pick.dataset.group = g
      pick.setAttribute('role', 'option')

      const del = document.createElement('button')
      del.type = 'button'
      del.className = 'saved-groups-delete'
      del.textContent = '삭제'
      del.dataset.group = g
      del.setAttribute('aria-label', `${g} 명단 삭제`)

      li.appendChild(pick)
      li.appendChild(del)
      savedGroupsListEl.appendChild(li)
    }
  }

  if (savedGroupsEmptyEl) {
    const empty = groups.length === 0
    savedGroupsEmptyEl.hidden = !empty
    if (savedGroupsListEl) savedGroupsListEl.hidden = empty
  }

  const last = localStorage.getItem(STORAGE_LAST_GROUP_V2)
  if (last && groups.includes(last)) {
    setSavedGroupSelection(last)
  } else if (groups.includes(currentValue)) {
    setSavedGroupSelection(currentValue)
  } else {
    setSavedGroupSelection('')
  }
}

function applyViewPerspective() {
  if (!seatBoardEl || !viewPerspectiveToggleBtn) return
  const isStudent = state.viewPerspective === 'student'
  // 사용자 요청: 교사뷰/학생뷰 표시를 서로 교체
  seatBoardEl.classList.toggle('perspective-teacher', isStudent)
  seatBoardEl.classList.toggle('perspective-student', !isStudent)
  viewPerspectiveToggleBtn.textContent = isStudent ? '교사뷰' : '학생뷰'
  viewPerspectiveToggleBtn.setAttribute(
    'aria-label',
    isStudent ? '교사 뷰로 자리표 보기' : '학생 뷰로 자리표 보기'
  )
}

/** 명단 불러오기 시 이전 자리표·사전 배치·고정 표시를 지움(저장본 사전 배치 적용 전에 호출) */
function clearSeatAssignmentsForNewRoster() {
  for (const seat of state.seats) {
    seat.student = ''
  }
  state.fixedAssignments.clear()
  state.preAssignments.clear()
  state.groupLeaders.clear()
  state.presetApplied = false
}

function getLeaderSeatIdInGroupRow(rowNumber) {
  for (const leaderSeatId of state.groupLeaders) {
    const leaderSeat = state.seats.find((s) => s.id === leaderSeatId)
    if (leaderSeat && leaderSeat.row === rowNumber) return leaderSeatId
  }
  return ''
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
    separatedRaw: separateInput?.value || '',
    traitBuckets: traitBucketsPayload(),
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

function loadStudentsFromLocal() {
  // 다른 그룹으로 전환할 때 이전 그룹의 분리 학생 텍스트가 잠깐이라도 남지 않게 먼저 비웁니다.
  if (separateInput) {
    separateInput.value = ''
  }
  setTraitBucketsFromSaved(null)

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
        if (separateInput) {
          const sep = typeof parsedV1.separatedRaw === 'string' ? parsedV1.separatedRaw : ''
          separateInput.value = sep
        }
        setTraitBucketsFromSaved(parsedV1?.traitBuckets)
        clearSeatAssignmentsForNewRoster()
        applyPreAssignmentsFromSavedObject(null, students)
        refreshPresetStudentSelect()
        renderSeats()
        renderPreassignedList()
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
  if (groupNameInput) {
    groupNameInput.value = groupName
  }
  if (separateInput) {
    const sep = typeof parsed.separatedRaw === 'string' ? parsed.separatedRaw : ''
    separateInput.value = sep
  }
  setTraitBucketsFromSaved(parsed?.traitBuckets)
  clearSeatAssignmentsForNewRoster()
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

function parseTraitBucketStudents(rawText) {
  return String(rawText || '')
    .split(/[\n,]+/)
    .map((name) => name.trim())
    .filter(Boolean)
}

function collectTraitBuckets(studentsList = []) {
  const entries = [
    ['1', parseTraitBucketStudents(traitBucket1Input?.value)],
    ['2', parseTraitBucketStudents(traitBucket2Input?.value)],
    ['3', parseTraitBucketStudents(traitBucket3Input?.value)],
    ['4', parseTraitBucketStudents(traitBucket4Input?.value)],
  ]
  const studentSet = new Set(studentsList)
  const traits = new Map()
  const mergedStudents = []
  let unknownStudentCount = 0
  let hasInput = false
  for (const [trait, names] of entries) {
    if (names.length > 0) hasInput = true
    for (const name of names) {
      if (studentSet.size > 0 && !studentSet.has(name)) {
        unknownStudentCount += 1
        continue
      }
      if (!traits.has(name)) mergedStudents.push(name)
      traits.set(name, trait)
    }
  }
  return { traits, mergedStudents, unknownStudentCount, hasInput }
}

function setTraitBucketsFromSaved(raw) {
  const parsed = raw && typeof raw === 'object' ? raw : {}
  if (traitBucket1Input) traitBucket1Input.value = String(parsed['1'] || '')
  if (traitBucket2Input) traitBucket2Input.value = String(parsed['2'] || '')
  if (traitBucket3Input) traitBucket3Input.value = String(parsed['3'] || '')
  if (traitBucket4Input) traitBucket4Input.value = String(parsed['4'] || '')
}

function traitBucketsPayload() {
  return {
    '1': traitBucket1Input?.value || '',
    '2': traitBucket2Input?.value || '',
    '3': traitBucket3Input?.value || '',
    '4': traitBucket4Input?.value || '',
  }
}

function shuffleArray(items) {
  const arr = [...items]
  for (let i = arr.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1))
    ;[arr[i], arr[j]] = [arr[j], arr[i]]
  }
  return arr
}

const SEAT_LAYOUTS = ['individual', 'pair', 'group', 'group_diverse']

function getSeatLayout() {
  const v = seatLayoutSelect?.value
  return SEAT_LAYOUTS.includes(v) ? v : 'individual'
}

function syncSeatDimensionLabels() {
  const layout = getSeatLayout()
  const g = layout === 'group' || layout === 'group_diverse'
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

function syncTraitBucketsVisibility() {
  if (!traitBucketsWrapEl) return
  const show = getSeatLayout() === 'group_diverse'
  traitBucketsWrapEl.hidden = !show
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
  const hasStudent = Boolean(displayName)

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
  if (state.groupLeaders.has(seat.id) && seat.student && !state.blockedSeats.has(seat.id)) {
    el.classList.add('group-leader')
  }

  if (state.blockedSeats.has(seat.id)) {
    el.innerHTML = `<span class="pos">${seat.index}</span><span class="blocked-icon" aria-hidden="true">X</span>`
  } else {
    const posMarkup = hasStudent ? '' : `<span class="pos">${seat.index}</span>`
    const crown = state.groupLeaders.has(seat.id) && seat.student ? '<span class="leader-crown" aria-hidden="true">👑</span>' : ''
    el.innerHTML = `${posMarkup}<span class="name">${displayName}</span>${crown}`
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

  const layout = getSeatLayout()
  const byId = new Map(state.seats.map((s) => [s.id, s]))
  const aoa = []
  const merges = []
  // 현재 코드에서는 teacher 상태가 학생뷰(반전), student 상태가 교사뷰(정방향)입니다.
  const isStudentView = state.viewPerspective === 'teacher'

  if (layout === 'group' || layout === 'group_diverse') {
    // 모둠 모드: 첫 열에 "n모둠" 표기 + 상단에 총 모둠 수 안내
    const totalCols = cols + 1
    const titleRow = Array(totalCols).fill('')
    titleRow[0] = '모둠 배치도'
    aoa.push(titleRow)
    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: totalCols - 1 } })

    const infoRow = Array(totalCols).fill('')
    infoRow[0] = `총 ${rows}모둠 / 모둠당 ${cols}명`
    aoa.push(infoRow)
    merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: totalCols - 1 } })

    const rowOrder = isStudentView
      ? Array.from({ length: rows }, (_, i) => rows - i)
      : Array.from({ length: rows }, (_, i) => i + 1)
    const colOrder = isStudentView
      ? Array.from({ length: cols }, (_, i) => cols - i)
      : Array.from({ length: cols }, (_, i) => i + 1)

    for (const r of rowOrder) {
      const row = [`${r}모둠`]
      for (const c of colOrder) {
        const seat = byId.get(`${r}-${c}`)
        row.push(seat ? seatCellExportText(seat) : '')
      }
      aoa.push(row)
    }
  } else if (layout === 'pair') {
    // 짝꿍 모드: 2자리 단위로 붙여두고, 짝 사이에는 빈 칸 1개를 넣어 시각적 간격 반영
    const pairGapCount = Math.floor((cols - 1) / 2)
    const totalCols = cols + pairGapCount
    const teacherRow = Array(totalCols).fill('')
    teacherRow[0] = '교탁'
    const rowOrder = isStudentView
      ? Array.from({ length: rows }, (_, i) => rows - i)
      : Array.from({ length: rows }, (_, i) => i + 1)
    const colOrder = isStudentView
      ? Array.from({ length: cols }, (_, i) => cols - i)
      : Array.from({ length: cols }, (_, i) => i + 1)

    if (!isStudentView) aoa.push(teacherRow)
    for (const r of rowOrder) {
      const row = []
      for (let i = 0; i < colOrder.length; i += 1) {
        const c = colOrder[i]
        const seat = byId.get(`${r}-${c}`)
        row.push(seat ? seatCellExportText(seat) : '')
        if (i % 2 === 1 && i < colOrder.length - 1) row.push('')
      }
      aoa.push(row)
    }
    if (isStudentView) aoa.push(teacherRow)
    const teacherRowIndex = isStudentView ? aoa.length - 1 : 0
    merges.push({ s: { r: teacherRowIndex, c: 0 }, e: { r: teacherRowIndex, c: totalCols - 1 } })
  } else {
    const teacherRow = Array(cols).fill('')
    teacherRow[0] = '교탁'
    const rowOrder = isStudentView
      ? Array.from({ length: rows }, (_, i) => rows - i)
      : Array.from({ length: rows }, (_, i) => i + 1)
    const colOrder = isStudentView
      ? Array.from({ length: cols }, (_, i) => cols - i)
      : Array.from({ length: cols }, (_, i) => i + 1)

    if (!isStudentView) aoa.push(teacherRow)
    for (const r of rowOrder) {
      const row = []
      for (const c of colOrder) {
        const seat = byId.get(`${r}-${c}`)
        row.push(seat ? seatCellExportText(seat) : '')
      }
      aoa.push(row)
    }
    if (isStudentView) aoa.push(teacherRow)
    const teacherRowIndex = isStudentView ? aoa.length - 1 : 0
    merges.push({ s: { r: teacherRowIndex, c: 0 }, e: { r: teacherRowIndex, c: cols - 1 } })
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa)
  ws['!merges'] = merges

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, '좌석배치')

  const group = getGroupFromUI()
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
  state.groupLeaders.clear()
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
  if (separateInput) {
    const sep = typeof parsed.separatedRaw === 'string' ? parsed.separatedRaw : ''
    separateInput.value = sep
  }
  setTraitBucketsFromSaved(parsed?.traitBuckets)
  applySeatLayoutFromSaved(parsed)
  applyPreAssignmentsFromSavedObject(parsed.preAssignments, students)
  if (groupNameInput) groupNameInput.value = groupName
  refreshSavedGroups()
  setSavedGroupSelection(groupName)
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
  state.groupLeaders.clear()
  const separatedPairs = parseSeparatedPairs(separateInput.value)
  const layout = getSeatLayout()
  const traitParsed = collectTraitBuckets(state.students)
  const studentTraits = traitParsed.traits

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

  const assignByGroupTraitDiversity = () => {
    const pool = shuffleArray(remainingStudents)
    const groupedSeats = new Map()
    for (const seat of availableSeats) {
      if (!groupedSeats.has(seat.row)) groupedSeats.set(seat.row, [])
      groupedSeats.get(seat.row).push(seat)
    }
    for (const rowSeats of groupedSeats.values()) {
      rowSeats.sort((a, b) => a.col - b.col)
    }
    const rows = Array.from(groupedSeats.keys()).sort((a, b) => a - b)
    const MAX_SAME_TRAIT_PER_GROUP = 2
    for (const row of rows) {
      const rowSeats = groupedSeats.get(row) || []
      /** 모둠(행) 안에서 특성별 인원 수 — 같은 특성은 최대 2명까지 */
      const traitCounts = new Map()
      for (const seat of state.seats) {
        if (seat.row !== row || !seat.student) continue
        const trait = studentTraits.get(seat.student)
        if (trait) traitCounts.set(trait, (traitCounts.get(trait) || 0) + 1)
      }
      for (const seat of rowSeats) {
        if (pool.length === 0) {
          seat.student = ''
          continue
        }
        const countOf = (trait) => (trait ? traitCounts.get(trait) || 0 : 0)
        let bestScore = Infinity
        const candidateIndices = []
        for (let i = 0; i < pool.length; i += 1) {
          const name = pool[i]
          const trait = studentTraits.get(name)
          let score
          if (!trait) {
            score = 0.5
          } else {
            const c = countOf(trait)
            if (c >= MAX_SAME_TRAIT_PER_GROUP) continue
            score = c === 0 ? 0 : 1
          }
          if (score < bestScore) {
            bestScore = score
            candidateIndices.length = 0
            candidateIndices.push(i)
          } else if (score === bestScore) {
            candidateIndices.push(i)
          }
        }
        let pickIndex =
          candidateIndices.length > 0
            ? candidateIndices[Math.floor(Math.random() * candidateIndices.length)]
            : 0
        const [picked] = pool.splice(pickIndex, 1)
        seat.student = picked || ''
        const pickedTrait = studentTraits.get(picked)
        if (pickedTrait) {
          traitCounts.set(pickedTrait, (traitCounts.get(pickedTrait) || 0) + 1)
        }
      }
    }
  }

  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    if (layout === 'group_diverse') {
      assignByGroupTraitDiversity()
    } else {
      const randomized = shuffleArray(remainingStudents)
      availableSeats.forEach((seat, idx) => {
        seat.student = idx < assignableCount ? randomized[idx] : ''
      })
    }
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
  } else if (layout === 'group_diverse') {
    let extra = ''
    if (traitParsed.unknownStudentCount > 0) {
      extra += ` 명단외 학생 ${traitParsed.unknownStudentCount}건 무시.`
    }
    updateStatus(`학생 특성 분류 모둠 배치가 완료되었습니다.${extra}`)
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
  state.groupLeaders.clear()
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
        state.groupLeaders.delete(existingSeatId)
        const oldSeat = state.seats.find((item) => item.id === existingSeatId)
        if (oldSeat) oldSeat.student = ''
      }
    }

    if (state.preAssignments.get(seatId) === selectedPresetStudent) {
      state.preAssignments.delete(seatId)
      seat.student = ''
      state.fixedAssignments.delete(seatId)
      state.groupLeaders.delete(seatId)
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

  if (
    !event.shiftKey &&
    (getSeatLayout() === 'group' || getSeatLayout() === 'group_diverse') &&
    seat.student &&
    !state.blockedSeats.has(seatId)
  ) {
    const currentLeader = getLeaderSeatIdInGroupRow(seat.row)
    if (currentLeader === seatId) {
      state.groupLeaders.delete(seatId)
      renderSeats()
      renderPreassignedList()
      updateStatus(`${seat.row}모둠장의 지정을 해제했습니다.`)
      return
    }
    if (currentLeader) state.groupLeaders.delete(currentLeader)
    state.groupLeaders.add(seatId)
    renderSeats()
    renderPreassignedList()
    updateStatus(`${seat.row}모둠장으로 ${seat.student} 학생을 지정했습니다.`)
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
    state.groupLeaders.delete(seatId)
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

savedGroupsTrigger?.addEventListener('click', (e) => {
  e.stopPropagation()
  toggleSavedGroupsPopover()
})

savedGroupsListEl?.addEventListener('click', (e) => {
  const delBtn = e.target.closest('.saved-groups-delete')
  const pickBtn = e.target.closest('.saved-groups-pick')
  if (delBtn?.dataset.group) {
    e.stopPropagation()
    openDeleteGroupModal(delBtn.dataset.group)
    return
  }
  if (pickBtn?.dataset.group) {
    setSavedGroupSelection(pickBtn.dataset.group)
    closeSavedGroupsPopover()
  }
})

document.addEventListener('click', (e) => {
  if (savedGroupsWidget?.contains(e.target)) return
  closeSavedGroupsPopover()
})

deleteGroupConfirmBtn?.addEventListener('click', () => {
  if (pendingDeleteGroupName) performDeleteSavedGroup(pendingDeleteGroupName)
})

deleteGroupCancelBtn?.addEventListener('click', () => {
  closeDeleteGroupModal()
})

deleteGroupModal?.addEventListener('click', (e) => {
  if (e.target.closest('[data-modal-dismiss]')) closeDeleteGroupModal()
})

document.addEventListener('keydown', (e) => {
  if (e.key !== 'Escape') return
  if (deleteGroupModal && !deleteGroupModal.hidden) {
    closeDeleteGroupModal()
    e.preventDefault()
    return
  }
  closeSavedGroupsPopover()
})
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
  syncTraitBucketsVisibility()
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
syncTraitBucketsVisibility()
buildSeatMap()