import { useState, useRef, useCallback, useMemo, useEffect } from 'react'
import './App.css'
import { createEngine } from './engine/core.js'

const TOTAL_ROWS = 50
const TOTAL_COLS = 50

export default function App() {
  // Engine instance is created once and reused across renders
  // Note: The engine maintains its own internal state, so React state is only used for UI updates
  const [engine] = useState(() => createEngine(TOTAL_ROWS, TOTAL_COLS))
  const [version, setVersion] = useState(0)
  const [selectedCell, setSelectedCell] = useState(null)
  const [editingCell, setEditingCell] = useState(null)
  const [rowOrder, setRowOrder] = useState(Array.from({ length: TOTAL_ROWS }, (_, i) => i))
  const [editValue, setEditValue] = useState('')
  const [loadError, setLoadError] = useState(null)
  const [sortConfig, setSortConfig] = useState({ col: null, direction: 'none' })
  const [filterState, setFilterState] = useState({})
  const [filterInputs, setFilterInputs] = useState({})
  const [activeFilterCol, setActiveFilterCol] = useState(null)
  const [selectedRange, setSelectedRange] = useState(null)
  const [internalClipboard, setInternalClipboard] = useState(null)

  const LOCAL_STORAGE_KEY = 'ai-spreadsheet-state-v1'

  // Cell styles are stored separately from engine data
  // Format: { "row,col": { bold: bool, italic: bool, ... } }
  const [cellStyles, setCellStyles] = useState({})
  const cellInputRef = useRef(null)

  // ────── Filtering and Sorting ──────

  const visibleRows = useMemo(() => {
    let rows = Array.from({ length: engine.rows }, (_, i) => i)

    // Apply filters
    rows = rows.filter(rowIndex => {
      for (const [colStr, filterValue] of Object.entries(filterState)) {
        if (!filterValue || filterValue.trim() === '') continue
        const col = parseInt(colStr)
        const cellData = engine.getCell(rowIndex, col)
        const displayValue = String(cellData.computed !== null && cellData.computed !== '' ? cellData.computed : cellData.raw).toLowerCase()
        if (!displayValue.includes(filterValue.toLowerCase())) {
          return false
        }
      }
      return true
    })

    // Apply sort
    if (sortConfig.col !== null && sortConfig.direction !== 'none') {
      rows.sort((rowA, rowB) => {
        const cellA = engine.getCell(rowA, sortConfig.col)
        const cellB = engine.getCell(rowB, sortConfig.col)

        const valueA = cellA.computed !== null && cellA.computed !== '' ? cellA.computed : cellA.raw
        const valueB = cellB.computed !== null && cellB.computed !== '' ? cellB.computed : cellB.raw

        const numA = parseFloat(valueA)
        const numB = parseFloat(valueB)

        let comparison = 0
        if (!isNaN(numA) && !isNaN(numB)) {
          comparison = numA - numB
        } else {
          const strA = String(valueA).toLowerCase()
          const strB = String(valueB).toLowerCase()
          comparison = strA.localeCompare(strB)
        }

        return sortConfig.direction === 'asc' ? comparison : -comparison
      })
    }

    return rows
  }, [engine, filterState, sortConfig, version])

  const forceRerender = useCallback(() => setVersion(v => v + 1), [])

  useEffect(() => {
    try {
      const saved = localStorage.getItem(LOCAL_STORAGE_KEY)
      if (!saved) return
      const parsed = JSON.parse(saved)
      if (parsed && parsed.engineState) {
        engine.hydrate(parsed.engineState)
        setCellStyles(parsed.cellStyles || {})
        setRowOrder(Array.from({ length: engine.rows }, (_, i) => i))
        setVersion(v => v + 1)
      }
    } catch (err) {
      console.error('Failed to load spreadsheet state from localStorage', err)
      setLoadError('Storage data invalid (reset)')
      localStorage.removeItem(LOCAL_STORAGE_KEY)
    }
  }, [engine])

  useEffect(() => {
    const timer = setTimeout(() => {
      try {
        const state = {
          engineState: engine.serialize(),
          cellStyles,
          rowOrder
        }
        localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(state))
      } catch (err) {
        if (err instanceof DOMException && (err.name === 'QuotaExceededError' || err.name === 'NS_ERROR_DOM_QUOTA_REACHED')) {
          console.error('LocalStorage quota exceeded, cannot save spreadsheet state', err)
        } else {
          console.error('Failed to save spreadsheet state to localStorage', err)
        }
      }
    }, 500)

    return () => clearTimeout(timer)
  }, [engine, cellStyles, rowOrder, version])

  // ────── Cell style helpers ──────

  const getCellStyle = useCallback((row, col) => {
    const key = `${row},${col}`
    return cellStyles[key] || {
      bold: false, italic: false, underline: false,
      bg: 'white', color: '#202124', align: 'left', fontSize: 13
    }
  }, [cellStyles])

  const updateCellStyle = useCallback((row, col, updates) => {
    const key = `${row},${col}`
    setCellStyles(prev => ({
      ...prev,
      [key]: { ...getCellStyle(row, col), ...updates }
    }))
  }, [getCellStyle])

  // ────── Cell editing ──────

  const startEditing = useCallback((row, col) => {
    setSelectedCell({ r: row, c: col })
    setEditingCell({ r: row, c: col })
    const cellData = engine.getCell(row, col)
    setEditValue(cellData.raw)
    setTimeout(() => cellInputRef.current?.focus(), 0)
  }, [engine])

  const commitEdit = useCallback((row, col) => {
    // Only commit if the value actually changed to avoid unnecessary recalculations
    const currentCell = engine.getCell(row, col)
    if (currentCell.raw !== editValue) {
      engine.setCell(row, col, editValue)
      forceRerender()
    }
    setEditingCell(null)
  }, [engine, editValue, forceRerender])

  const handleCellClick = useCallback((event, row, col) => {
    if (editingCell && (editingCell.r !== row || editingCell.c !== col)) {
      commitEdit(editingCell.r, editingCell.c)
    }

    const newRange = event.shiftKey && selectedCell
      ? {
        r1: Math.min(selectedCell.r, row),
        c1: Math.min(selectedCell.c, col),
        r2: Math.max(selectedCell.r, row),
        c2: Math.max(selectedCell.c, col)
      }
      : { r1: row, c1: col, r2: row, c2: col }

    setSelectedRange(newRange)
    setSelectedCell({ r: row, c: col })

    if (!editingCell || editingCell.r !== row || editingCell.c !== col) {
      startEditing(row, col)
    }
  }, [editingCell, commitEdit, selectedCell, startEditing])

  // ────── Keyboard navigation ──────

  const handleKeyDown = useCallback((event, row, col) => {
    if (event.key === 'Enter') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col)
    } else if (event.key === 'Tab') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    } else if (event.key === 'Escape') {
      setEditValue(engine.getCell(row, col).raw)
      setEditingCell(null)
    } else if (event.key === 'ArrowDown') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col)
    } else if (event.key === 'ArrowUp') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.max(row - 1, 0), col)
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault()
      commitEdit(row, col)
      if (col > 0) {
        startEditing(row, col - 1)
      } else if (row > 0) {
        startEditing(row - 1, engine.cols - 1)
      }
    } else if (event.key === 'ArrowRight') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    }
  }, [engine, commitEdit, startEditing])

  // ────── Formula bar handlers ──────

  const handleFormulaBarKeyDown = useCallback((event) => {
    if (!editingCell) return
    handleKeyDown(event, editingCell.r, editingCell.c)
  }, [editingCell, handleKeyDown])

  const handleFormulaBarFocus = useCallback(() => {
    if (selectedCell && !editingCell) {
      setEditingCell(selectedCell)
      setEditValue(engine.getCell(selectedCell.r, selectedCell.c).raw)
    }
  }, [selectedCell, editingCell, engine])

  const handleFormulaBarChange = useCallback((value) => {
    if (!editingCell && selectedCell) setEditingCell(selectedCell)
    setEditValue(value)
  }, [editingCell, selectedCell])

  // ────── Undo / Redo ──────

  const handleUndo = useCallback(() => { if (engine.undo()) forceRerender() }, [engine, forceRerender])
  const handleRedo = useCallback(() => { if (engine.redo()) forceRerender() }, [engine, forceRerender])

  // ────── Formatting toggles ──────

  const toggleBold = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { bold: !style.bold })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleItalic = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { italic: !style.italic })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleUnderline = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { underline: !style.underline })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const changeFontSize = useCallback((size) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { fontSize: size })
  }, [selectedCell, updateCellStyle])

  const changeAlignment = useCallback((align) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { align })
  }, [selectedCell, updateCellStyle])

  const changeFontColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { color })
  }, [selectedCell, updateCellStyle])

  const changeBackgroundColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { bg: color })
  }, [selectedCell, updateCellStyle])

  // ────── Clear operations ──────

  const clearSelectedCell = useCallback(() => {
    if (!selectedCell) return
    engine.setCell(selectedCell.r, selectedCell.c, '')
    forceRerender()
    // Remove style entry for cleared cell
    // Note: This deletes the style object entirely - if you need to preserve default styles,
    // you may want to set them explicitly rather than deleting
    const key = `${selectedCell.r},${selectedCell.c}`
    setCellStyles(prev => { const next = { ...prev }; delete next[key]; return next })
    setEditValue('')
  }, [selectedCell, engine, forceRerender])

  const clearAllCells = useCallback(() => {
    for (let r = 0; r < engine.rows; r++) {
      for (let c = 0; c < engine.cols; c++) {
        engine.setCell(r, c, '')
      }
    }
    forceRerender()
    setCellStyles({})
    setSelectedCell(null)
    setEditingCell(null)
    setEditValue('')
  }, [engine, forceRerender])

  // ────── Row / Column operations ──────

  const insertRow = useCallback(() => {
    if (!selectedCell) return
    engine.insertRow(selectedCell.r)
    forceRerender()
    setSelectedCell({ r: selectedCell.r + 1, c: selectedCell.c })
  }, [selectedCell, engine, forceRerender])

  const deleteRow = useCallback(() => {
    if (!selectedCell) return
    engine.deleteRow(selectedCell.r)
    forceRerender()
    if (selectedCell.r >= engine.rows) {
      setSelectedCell({ r: engine.rows - 1, c: selectedCell.c })
    }
  }, [selectedCell, engine, forceRerender])

  const insertColumn = useCallback(() => {
    if (!selectedCell) return
    engine.insertColumn(selectedCell.c)
    forceRerender()
    setSelectedCell({ r: selectedCell.r, c: selectedCell.c + 1 })
  }, [selectedCell, engine, forceRerender])

  const deleteColumn = useCallback(() => {
    if (!selectedCell) return
    engine.deleteColumn(selectedCell.c)
    forceRerender()
    if (selectedCell.c >= engine.cols) {
      setSelectedCell({ r: selectedCell.r, c: engine.cols - 1 })
    }
  }, [selectedCell, engine, forceRerender])

  // ────── Sort & Filter handlers ──────

  const handleSortColumn = useCallback((col) => {
    setSortConfig(prev => {
      if (prev.col === col) {
        // Cycle direction
        if (prev.direction === 'asc') return { col, direction: 'desc' }
        if (prev.direction === 'desc') return { col, direction: 'none' }
        return { col, direction: 'asc' }
      } else {
        // New column, start with asc
        return { col, direction: 'asc' }
      }
    })
  }, [])

  const parseClipboardText = useCallback((text) => {
    if (!text || typeof text !== 'string') return []
    const rows = text.split(/\r?\n/)
      // Remove final empty row from trailing newline
      .filter((row, idx, arr) => !(idx === arr.length - 1 && row.trim() === ''))
      .map(row => row.split('\t'))
    return rows
  }, [])

  const pasteTextToGrid = useCallback((text) => {
    if (!text || !selectedCell) return

    const clipboardTable = parseClipboardText(text)
    if (clipboardTable.length === 0) return

    const pasteStart = selectedRange || selectedCell
    const startRow = pasteStart?.r1 !== undefined ? pasteStart.r1 : pasteStart.r
    const startCol = pasteStart?.c1 !== undefined ? pasteStart.c1 : pasteStart.c

    if (startRow === undefined || startCol === undefined) return

    setInternalClipboard(clipboardTable)

    const maxRow = startRow + clipboardTable.length - 1
    const maxCol = startCol + Math.max(...clipboardTable.map(row => row.length)) - 1

    console.log('Pasting data table:', clipboardTable)

    for (let r = 0; r < clipboardTable.length; r++) {
      for (let c = 0; c < clipboardTable[r].length; c++) {
        const targetRow = startRow + r
        const targetCol = startCol + c
        if (targetRow >= engine.rows || targetCol >= engine.cols) continue
        engine.setCell(targetRow, targetCol, clipboardTable[r][c])
      }
    }

    forceRerender()
    setSelectedCell({ r: Math.min(maxRow, engine.rows - 1), c: Math.min(maxCol, engine.cols - 1) })
    setSelectedRange({ r1: startRow, c1: startCol, r2: Math.min(maxRow, engine.rows - 1), c2: Math.min(maxCol, engine.cols - 1) })
  }, [engine, forceRerender, parseClipboardText, selectedCell])

  const handlePaste = useCallback((e) => {
    e.preventDefault()
    const text = e.clipboardData?.getData('text') || ''
    console.log('Paste event text:', JSON.stringify(text))
    pasteTextToGrid(text)
  }, [pasteTextToGrid])

  const pasteClipboard = useCallback(async () => {
    if (!selectedCell) return

    let text = ''
    try {
      if (navigator.clipboard && navigator.clipboard.readText) {
        text = await navigator.clipboard.readText()
      }
    } catch (err) {
      console.warn('Navigator clipboard readText failed, using internal clipboard fallback', err)
    }

    if (text && text.trim() !== '') {
      pasteTextToGrid(text)
      return
    }

    // fallback to internal clipboard
    if (internalClipboard && internalClipboard.length > 0) {
      const clipboardTable = internalClipboard
      const startRow = selectedCell.r
      const startCol = selectedCell.c

      for (let r = 0; r < clipboardTable.length; r++) {
        for (let c = 0; c < clipboardTable[r].length; c++) {
          const targetRow = startRow + r
          const targetCol = startCol + c
          if (targetRow >= engine.rows || targetCol >= engine.cols) continue
          engine.setCell(targetRow, targetCol, clipboardTable[r][c])
        }
      }

      forceRerender()
      const maxRow = startRow + clipboardTable.length - 1
      const maxCol = startCol + Math.max(...clipboardTable.map(row => row.length)) - 1
      setSelectedCell({ r: Math.min(maxRow, engine.rows - 1), c: Math.min(maxCol, engine.cols - 1) })
      setSelectedRange({ r1: startRow, c1: startCol, r2: Math.min(maxRow, engine.rows - 1), c2: Math.min(maxCol, engine.cols - 1) })
    }
  }, [engine, forceRerender, internalClipboard, pasteTextToGrid, selectedCell])

  const handleFilterChange = useCallback((col, value) => {
    setFilterInputs(prev => ({ ...prev, [col]: value }))
    setFilterState(prev => {
      if (!value || value.trim() === '') {
        const next = { ...prev }
        delete next[col]
        return next
      } else {
        return { ...prev, [col]: value }
      }
    })
  }, [])

  const handleClearFilter = useCallback((col) => {
    setFilterInputs(prev => ({ ...prev, [col]: '' }))
    setFilterState(prev => {
      const next = { ...prev }
      delete next[col]
      return next
    })
  }, [])

  const normalizeRange = useCallback((range) => {
    if (!range) return null
    const r1 = Math.min(range.r1, range.r2)
    const r2 = Math.max(range.r1, range.r2)
    const c1 = Math.min(range.c1, range.c2)
    const c2 = Math.max(range.c1, range.c2)
    return { r1, r2, c1, c2 }
  }, [])

  const getSelectedRange = useCallback(() => {
    if (!selectedRange && selectedCell) {
      return { r1: selectedCell.r, r2: selectedCell.r, c1: selectedCell.c, c2: selectedCell.c }
    }
    return normalizeRange(selectedRange)
  }, [selectedCell, selectedRange, normalizeRange])

  const copySelectedRange = useCallback(async () => {
    const range = getSelectedRange()
    if (!range) return

    const rows = []
    for (let r = range.r1; r <= range.r2; r++) {
      const cols = []
      for (let c = range.c1; c <= range.c2; c++) {
        const cell = engine.getCell(r, c)
        let value = cell.computed
        if (value === null || value === undefined || value === '') value = cell.raw || ''
        cols.push(value)
      }
      rows.push(cols)
    }

    const tsv = rows.map(r => r.map(item => String(item ?? '')).join('\t')).join('\n')
    setInternalClipboard(rows)

    console.log('Copy range TSV:\n' + tsv)

    try {
      await navigator.clipboard.writeText(tsv)
      console.log('Copied data to clipboard successfully')
    } catch (err) {
      console.warn('Clipboard write failed:', err)
    }
  }, [engine, getSelectedRange])

  useEffect(() => {
    const onKeyDown = (event) => {
      if (!event.ctrlKey && !event.metaKey) return

      if (event.key === 'c' || event.key === 'C') {
        event.preventDefault()
        copySelectedRange()
      } else if (event.key === 'v' || event.key === 'V') {
        event.preventDefault()
        pasteClipboard()
      }
    }

    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [copySelectedRange, pasteClipboard])

  const handleClearAllFilters = useCallback(() => {
    setFilterInputs({})
    setFilterState({})
    setSortConfig({ col: null, direction: 'none' })
  }, [])

  // ────── Derived state ──────

  const selectedCellStyle = useMemo(() => {
    return selectedCell ? getCellStyle(selectedCell.r, selectedCell.c) : null
  }, [selectedCell, getCellStyle])

  const getColumnLabel = useCallback((col) => {
    let label = ''
    let num = col + 1
    while (num > 0) {
      num--
      label = String.fromCharCode(65 + (num % 26)) + label
      num = Math.floor(num / 26)
    }
    return label
  }, [])

  const selectedCellLabel = selectedCell
    ? `${getColumnLabel(selectedCell.c)}${selectedCell.r + 1}`
    : 'No cell'

  // Formula bar shows the raw formula text, not the computed value
  // When editing, show the current editValue; otherwise show the cell's raw content
  // Note: This is different from the cell display, which shows computed values
  const formulaBarValue = editingCell
    ? editValue
    : (selectedCell ? engine.getCell(selectedCell.r, selectedCell.c).raw : '')

  // ────── Render ──────

  return (
    <div className="app-wrapper" onPaste={handlePaste}>
      <div className="app-header">
        <h2 className="app-title">📊 Spreadsheet App</h2>
      </div>

      <div className="main-content">

        {/* ── Toolbar ── */}
        <div className="toolbar">
          <div className="toolbar-group">
            <button className={`toolbar-btn bold-btn ${selectedCellStyle?.bold ? 'active' : ''}`} onClick={toggleBold} title="Bold">B</button>
            <button className={`toolbar-btn italic-btn ${selectedCellStyle?.italic ? 'active' : ''}`} onClick={toggleItalic} title="Italic">I</button>
            <button className={`toolbar-btn underline-btn ${selectedCellStyle?.underline ? 'active' : ''}`} onClick={toggleUnderline} title="Underline">U</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Size:</span>
            <select className="toolbar-select" value={selectedCellStyle?.fontSize || 13} onChange={(e) => changeFontSize(parseInt(e.target.value))}>
              {[8, 10, 11, 12, 13, 14, 16, 18, 20, 24].map(s => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
          </div>

          <div className="toolbar-group">
            <button className={`align-btn ${selectedCellStyle?.align === 'left' ? 'active' : ''}`} onClick={() => changeAlignment('left')} title="Align Left">⬤←</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'center' ? 'active' : ''}`} onClick={() => changeAlignment('center')} title="Align Center">⬤</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'right' ? 'active' : ''}`} onClick={() => changeAlignment('right')} title="Align Right">⬤→</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Text:</span>
            <input
              type="color"
              value={selectedCellStyle?.color || '#000000'}
              onChange={(e) => changeFontColor(e.target.value)}
              title="Font color"
              style={{ width: '32px', height: '32px', border: '1px solid #dadce0', cursor: 'pointer', borderRadius: '4px' }}
            />
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Fill:</span>
            <select className="toolbar-select" value={selectedCellStyle?.bg || 'white'} onChange={(e) => changeBackgroundColor(e.target.value)}>
              <option value="white">White</option>
              <option value="#ffff99">Yellow</option>
              <option value="#99ffcc">Green</option>
              <option value="#ffcccc">Red</option>
              <option value="#cce5ff">Blue</option>
              <option value="#e0ccff">Purple</option>
              <option value="#ffd9b3">Orange</option>
              <option value="#f0f0f0">Gray</option>
            </select>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={handleUndo} disabled={!engine.canUndo()} title="Undo">↶ Undo</button>
            <button className="toolbar-btn" onClick={handleRedo} disabled={!engine.canRedo()} title="Redo">↷ Redo</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={handleClearAllFilters} title="Clear all sorts and filters">🔄 Reset</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={insertRow} title="Insert Row">+ Row</button>
            <button className="toolbar-btn" onClick={deleteRow} title="Delete Row">- Row</button>
            <button className="toolbar-btn" onClick={insertColumn} title="Insert Column">+ Col</button>
            <button className="toolbar-btn" onClick={deleteColumn} title="Delete Column">- Col</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn danger" onClick={clearSelectedCell}>✕ Cell</button>
            <button className="toolbar-btn danger" onClick={clearAllCells}>✕ All</button>
          </div>
        </div>

        {/* ── Formula Bar ── */}
        <div className="formula-bar">
          <span className="formula-bar-label">{selectedCellLabel}</span>
          <input
            className="formula-bar-input"
            value={formulaBarValue}
            onChange={(e) => handleFormulaBarChange(e.target.value)}
            onKeyDown={handleFormulaBarKeyDown}
            onFocus={handleFormulaBarFocus}
            placeholder="Select a cell then type, or enter a formula like =SUM(A1:A5)"
          />
        </div>

        {/* ── Grid ── */}
        <div className="grid-scroll">
          <table className="grid-table">
            <thead>
              <tr>
                <th className="col-header-blank"></th>
                {Array.from({ length: engine.cols }, (_, colIndex) => (
                  <th key={colIndex} className="col-header">
                    <div className="col-header-content">
                      <div className="col-header-label" onClick={() => handleSortColumn(colIndex)} style={{ cursor: 'pointer', flex: 1 }}>
                        {getColumnLabel(colIndex)}
                        {sortConfig.col === colIndex && sortConfig.direction !== 'none' && (
                          <span className="sort-indicator" title={sortConfig.direction === 'asc' ? 'Ascending' : 'Descending'}>
                            {sortConfig.direction === 'asc' ? ' ↑' : ' ↓'}
                          </span>
                        )}
                      </div>
                      <div className="col-header-filter">
                        <button
                          className="filter-btn"
                          onClick={() => setActiveFilterCol(activeFilterCol === colIndex ? null : colIndex)}
                          title="Filter"
                          style={{
                            background: filterState[colIndex] ? '#1a73e8' : '#f8f9fa',
                            color: filterState[colIndex] ? '#fff' : '#5f6368',
                          }}
                        >
                          ⋮
                        </button>
                        {activeFilterCol === colIndex && (
                          <div className="filter-dropdown">
                            <input
                              type="text"
                              className="filter-input"
                              placeholder="Filter..."
                              value={filterInputs[colIndex] || ''}
                              onChange={(e) => handleFilterChange(colIndex, e.target.value)}
                              autoFocus
                            />
                            {filterState[colIndex] && (
                              <button className="filter-clear-btn" onClick={() => handleClearFilter(colIndex)}>
                                Clear
                              </button>
                            )}
                          </div>
                        )}
                      </div>
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {visibleRows.map((rowIndex) => (
                <tr key={rowIndex}>
                  <td className="row-header">{rowIndex + 1}</td>
                  {Array.from({ length: engine.cols }, (_, colIndex) => {
                    const inRange = selectedRange && rowIndex >= Math.min(selectedRange.r1, selectedRange.r2) && rowIndex <= Math.max(selectedRange.r1, selectedRange.r2) && colIndex >= Math.min(selectedRange.c1, selectedRange.c2) && colIndex <= Math.max(selectedRange.c1, selectedRange.c2)
                    const isSelected = selectedCell?.r === rowIndex && selectedCell?.c === colIndex
                    const isEditing = editingCell?.r === rowIndex && editingCell?.c === colIndex
                    const cellData = engine.getCell(rowIndex, colIndex)
                    const style = cellStyles[`${rowIndex},${colIndex}`] || {}
                    const displayValue = cellData.error
                      ? cellData.error
                      : (cellData.computed !== null && cellData.computed !== '' ? String(cellData.computed) : cellData.raw)

                    return (
                      <td
                        key={colIndex}
                        className={`cell ${isSelected || inRange ? 'selected' : ''}`}
                        style={{ background: style.bg || 'white' }}
                        onMouseDown={(e) => { e.preventDefault(); handleCellClick(e, rowIndex, colIndex) }}
                      >
                        {isEditing ? (
                          <input
                            autoFocus
                            onPaste={handlePaste}
                            className="cell-input"
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onBlur={() => commitEdit(rowIndex, colIndex)}
                            onKeyDown={(e) => handleKeyDown(e, rowIndex, colIndex)}
                            ref={isSelected ? cellInputRef : undefined}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: style.color || '#202124',
                              fontSize: (style.fontSize || 13) + 'px',
                              textAlign: style.align || 'left',
                              background: style.bg || 'white',
                            }}
                          />
                        ) : (
                          <div
                            className={`cell-display align-${style.align || 'left'} ${cellData.error ? 'error' : ''}`}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: cellData.error ? '#d93025' : (style.color || '#202124'),
                              fontSize: (style.fontSize || 13) + 'px',
                            }}
                          >
                            {displayValue}
                          </div>
                        )}
                      </td>
                    )
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <p className="footer-hint">
          Click a cell to edit · Enter/Tab/Arrow keys to navigate · Formulas: =A1+B1 · =SUM(A1:A5) · =AVG(A1:A5) · =MAX(A1:A5) · =MIN(A1:A5)
        </p>
      </div>
    </div>
  )
}
