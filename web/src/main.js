/**
 * main.js — WebAIWord minimal frontend
 *
 * Responsibilities:
 *  1. On startup: load API config from localStorage and populate the modal fields.
 *  2. On "Save" in the API modal: persist config to localStorage.
 *  3. On startup: if a cached ai_view exists, show the "restore session" banner.
 *  4. On "Restore": load cached ai_view into the editor textarea.
 *  5. On editor input: throttled save of ai_view to localStorage.
 *  6. On "Clear cache": storageClearAll() and refresh UI.
 */

import {
  STORAGE_KEYS,
  storageGet,
  storageSet,
  storageRemove,
  storageClearAll,
} from './storageService.js'

// ─── Constants ────────────────────────────────────────────────────────────────
const DEFAULT_MODEL = 'gpt-4o'

// ─── DOM refs ───────────────────────────────────────────────────────────────
const btnApiKey      = document.getElementById('btn-apikey')
const btnClearCache  = document.getElementById('btn-clear-cache')
const restoreBanner  = document.getElementById('restore-banner')
const btnRestore     = document.getElementById('btn-restore')
const btnDismiss     = document.getElementById('btn-dismiss')
const restoreDocName = document.getElementById('restore-doc-name')
const editor         = document.getElementById('editor')
const statusBar      = document.getElementById('status-bar')

const modalApiKey    = document.getElementById('modal-apikey')
const btnSaveApiKey  = document.getElementById('btn-save-apikey')
const btnCloseModal  = document.getElementById('btn-close-modal')
const selectProvider = document.getElementById('select-provider')
const labelBaseUrl   = document.getElementById('label-baseurl')
const inputApiKey    = document.getElementById('input-apikey')
const inputBaseUrl   = document.getElementById('input-baseurl')
const inputModel     = document.getElementById('input-model')

// ─── Helpers ─────────────────────────────────────────────────────────────────
function showStatus(msg, isError = false) {
  statusBar.textContent = msg
  statusBar.style.color = isError ? '#dc2626' : '#64748b'
}

// ─── 1. Load API config on startup ───────────────────────────────────────────
function loadConfigIntoModal() {
  const provider = storageGet(STORAGE_KEYS.PROVIDER) || 'openai'
  const apiKey   = storageGet(STORAGE_KEYS.API_KEY)  || ''
  const baseUrl  = storageGet(STORAGE_KEYS.BASE_URL) || ''
  const model    = storageGet(STORAGE_KEYS.MODEL)    || ''

  selectProvider.value = provider
  inputApiKey.value    = apiKey
  inputBaseUrl.value   = baseUrl
  inputModel.value     = model
  // Show/hide Base URL field
  labelBaseUrl.classList.toggle('hidden', provider !== 'custom')
}

// ─── 2. Save API config ───────────────────────────────────────────────────────
function saveConfig() {
  const provider = selectProvider.value
  const apiKey   = inputApiKey.value.trim()
  const baseUrl  = inputBaseUrl.value.trim()
  const model    = inputModel.value.trim() || DEFAULT_MODEL

  if (!apiKey) {
    showStatus('⚠️ 请输入有效的 API Key', true)
    return
  }

  storageSet(STORAGE_KEYS.PROVIDER, provider)
  storageSet(STORAGE_KEYS.API_KEY,  apiKey)
  storageSet(STORAGE_KEYS.MODEL,    model)
  storageSet(STORAGE_KEYS.BASE_URL, baseUrl)

  modalApiKey.classList.add('hidden')
  showStatus(`✅ API 配置已保存（Provider: ${provider}）`)
}

// ─── 3. Restore session banner ────────────────────────────────────────────────
function checkRestoreSession() {
  const cached = storageGet(STORAGE_KEYS.AI_VIEW)
  if (!cached) return

  const docName = storageGet(STORAGE_KEYS.DOC_NAME) || '上次的文档'
  restoreDocName.textContent = docName
  restoreBanner.classList.remove('hidden')
}

// ─── 4. Restore cached ai_view ────────────────────────────────────────────────
function restoreSession() {
  const cached = storageGet(STORAGE_KEYS.AI_VIEW)
  if (!cached) return
  editor.value = cached
  restoreBanner.classList.add('hidden')
  showStatus('✅ 已恢复上次的编辑内容')
}

// ─── 5. Throttled save of editor content ─────────────────────────────────────
let saveTimer = null
const SAVE_DELAY_MS = 800

function scheduleAiViewSave() {
  clearTimeout(saveTimer)
  saveTimer = setTimeout(() => {
    const value = editor.value.trim()
    if (!value) return

    const ok = storageSet(STORAGE_KEYS.AI_VIEW, value)
    if (ok) {
      showStatus('💾 已自动保存编辑内容')
    } else {
      // Quota exceeded – try a truncated/plain version or warn
      showStatus('⚠️ 文档过大，无法缓存到 localStorage', true)
    }
  }, SAVE_DELAY_MS)
}

// ─── 6. Clear cache ───────────────────────────────────────────────────────────
function clearCache() {
  if (!confirm('确认清除所有本地缓存（包括 API Key 和文档内容）？')) return
  storageClearAll()
  editor.value = ''
  loadConfigIntoModal()     // reset form fields
  restoreBanner.classList.add('hidden')
  showStatus('🗑️ 本地缓存已清除')
}

// ─── Event listeners ─────────────────────────────────────────────────────────
btnApiKey.addEventListener('click', () => {
  loadConfigIntoModal()
  modalApiKey.classList.remove('hidden')
})
btnCloseModal.addEventListener('click', () => modalApiKey.classList.add('hidden'))
modalApiKey.addEventListener('click', (e) => {
  if (e.target === modalApiKey) modalApiKey.classList.add('hidden')
})
selectProvider.addEventListener('change', () => {
  labelBaseUrl.classList.toggle('hidden', selectProvider.value !== 'custom')
})
btnSaveApiKey.addEventListener('click', saveConfig)
btnClearCache.addEventListener('click', clearCache)
btnRestore.addEventListener('click', restoreSession)
btnDismiss.addEventListener('click', () => {
  restoreBanner.classList.add('hidden')
  storageRemove(STORAGE_KEYS.AI_VIEW)
  storageRemove(STORAGE_KEYS.DOC_NAME)
  showStatus('上次的缓存内容已忽略')
})
editor.addEventListener('input', scheduleAiViewSave)

// File import: read JSON file and put contents into editor
const fileInput = document.getElementById('file-input')
document.getElementById('btn-import').addEventListener('click', () => fileInput.click())
fileInput.addEventListener('change', async (e) => {
  const file = e.target.files?.[0]
  if (!file) return
  fileInput.value = ''

  try {
    const text = await file.text()
    // Validate that it is parseable JSON
    JSON.parse(text)
    editor.value = text
    storageSet(STORAGE_KEYS.AI_VIEW,  text)
    storageSet(STORAGE_KEYS.DOC_NAME, file.name)
    restoreBanner.classList.add('hidden')
    showStatus(`✅ 已导入：${file.name}`)
  } catch {
    showStatus('❌ 文件格式错误，请选择有效的 ai_view JSON 文件', true)
  }
})

// ─── Startup ──────────────────────────────────────────────────────────────────
loadConfigIntoModal()
checkRestoreSession()
showStatus('就绪。请导入 ai_view JSON 文件，或从上方恢复上次编辑。')
