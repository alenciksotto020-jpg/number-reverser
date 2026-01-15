import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';

export default function NumberReverser() {
  const [data, setData] = useState(null);
  const [fileName, setFileName] = useState('');
  const [stats, setStats] = useState({ input: 0, output: 0, reversed: 0, skipped: 0, splits: 0 });
  const [isDragging, setIsDragging] = useState(false);
  const [activeTab, setActiveTab] = useState('upload');
  const [copied, setCopied] = useState(false);

  // Patterns
  const P = {
    m1: /^[hv]\s*1$/i,
    m2: /^[ws]\s*2$/i,
    combined: /^\d+(\.\d+)?\s*\+\s*[hv]\s*1$/i,
    imposts: /imposts?:?/i,
    num: /^(\d+(\.\d+)?)/
  };

  // Helpers
  const test = (cell, p) => typeof cell === 'string' && p.test(cell.trim());
  const findIdx = (row, p) => row.findIndex(c => test(c, p));
  const findAllIdx = (row, p) => {
    const indices = [];
    for (let i = 0; i < row.length; i++) {
      if (test(row[i], p)) indices.push(i);
    }
    return indices;
  };
  const extractNum = (cell) => {
    const format = (n) => n % 1 === 0 ? Math.floor(n) : n;
    if (typeof cell === 'number') return format(cell);
    if (typeof cell === 'string') {
      const m = cell.trim().match(P.num);
      return m ? format(parseFloat(m[1])) : null;
    }
    return null;
  };
  const getMarkerLetter = (cell) => cell?.trim().match(/^([hv])/i)?.[1].toLowerCase();
  const getCombinedLetter = (cell) => cell?.trim().match(/\+\s*([hv])\s*1$/i)?.[1].toLowerCase();
  const altMarker = (l) => l === 'h' ? 'v' : 'h';
  const getM2 = (l) => l === 'h' ? 'W' : 'S';

  const getNums = (row, start, end = row.length, excludeCombined = false) => {
    const nums = [];
    for (let i = start + 1; i < end && i < row.length; i++) {
      if (excludeCombined && test(row[i], P.combined)) continue;
      const n = extractNum(row[i]);
      if (n !== null) nums.push(n);
    }
    return nums;
  };

  const isCompletelyEmptyRow = (row) => {
    return row.every(cell =>
      cell === '' ||
      cell === null ||
      cell === undefined ||
      (typeof cell === 'string' && cell.trim() === '')
    );
  };

  const buildRow = (len, marker, markerIdx, impostsIdx, impostsVal, nums) => {
    const row = Array(Math.max(len, markerIdx + 1 + nums.length)).fill('');
    if (impostsIdx !== -1) row[impostsIdx] = impostsVal;
    row[markerIdx] = marker;
    nums.forEach((n, i) => row[markerIdx + 1 + i] = n);
    return row;
  };

  const process = (input) => {
    const result = [];
    let reversed = 0, skipped = 0, splits = 0;

    for (let i = 0; i < input.length; i++) {
      const row = input[i];
      const m1Idx = findIdx(row, P.m1);

      if (m1Idx === -1) {
        if (findIdx(row, P.m2) === -1) result.push({ data: [...row], type: 'unchanged' });
        continue;
      }

      const letter = getMarkerLetter(row[m1Idx]);
      const combIdx = findIdx(row, P.combined);
      const impIdx = findIdx(row, P.imposts);
      const hasNext = i + 1 < input.length && findIdx(input[i + 1], P.m2) !== -1;

      if (combIdx !== -1) {
        // Handle multiple combined patterns
        const combIndices = findAllIdx(row, P.combined);
        let allValid = true;
        let currentStart = m1Idx;
        let currentLetter = letter;
        const segments = [];

        // Extract all segments
        for (let j = 0; j < combIndices.length; j++) {
          const idx = combIndices[j];
          const nums = [...getNums(row, currentStart, idx, true), extractNum(row[idx])];
          if (nums.length % 2 === 1) {
            segments.push({ nums, letter: currentLetter });
          } else {
            allValid = false;
            break;
          }
          currentStart = idx;
          currentLetter = getCombinedLetter(row[idx]);
        }

        // Check final segment
        const finalNums = getNums(row, combIndices[combIndices.length - 1], row.length, true);
        if (finalNums.length > 0 && finalNums.length % 2 === 1) {
          const finalLetter = getCombinedLetter(row[combIndices[combIndices.length - 1]]);
          segments.push({ nums: finalNums, letter: finalLetter });
        }

        if (allValid && segments.length > 0) {
          splits += segments.length - 1;

          // Process first segment
          const firstSeg = segments[0];
          const r1 = [...row];
          r1[combIndices[0]] = extractNum(row[combIndices[0]]);
          for (let j = combIndices[0] + 1; j < r1.length; j++) r1[j] = '';
          result.push({ data: r1, type: 'original' });

          if (hasNext) {
            const r2 = buildRow(input[++i].length, `${getM2(firstSeg.letter)}2`, m1Idx, impIdx, ':', [...firstSeg.nums].reverse());
            result.push({ data: r2, type: 'filled' });
            reversed++;
          }

          // Process remaining segments
          for (let j = 1; j < segments.length; j++) {
            const seg = segments[j];
            result.push({ data: buildRow(row.length, `${seg.letter} 1`, m1Idx, impIdx, 'Imposts:', seg.nums), type: 'original', isNew: true });
            result.push({ data: buildRow(row.length, `${getM2(seg.letter)}2`, m1Idx, impIdx, ':', [...seg.nums].reverse()), type: 'filled', isNew: true });
            reversed++;
          }
        } else {
          if (hasNext) i++;
          skipped++;
        }
        continue;
      }

      const nums = getNums(row, m1Idx);
      if (nums.length > 0 && nums.length % 2 === 1) {
        result.push({ data: [...row], type: 'original' });
        if (hasNext) {
          result.push({ data: buildRow(input[++i].length, `${getM2(letter)}2`, m1Idx, impIdx, ':', [...nums].reverse()), type: 'filled' });
          reversed++;
        }
      } else {
        if (hasNext) i++;
        skipped++;
      }
    }

    // Remove completely empty rows (where ALL cells are blank)
    const cleanedResult = result.filter(r => !isCompletelyEmptyRow(r.data));

    setData(cleanedResult);
    setStats({ input: input.length, output: cleanedResult.length, reversed, skipped, splits });
  };

  const parse = (d, isFile) => process(isFile ? d : d.trim().split('\n').map(l => l.split(/\t/).map(c => {
    const t = c.trim(), n = parseFloat(t);
    return !isNaN(n) && t === String(n) ? n : t;
  })));

  const handleFile = useCallback((file) => {
    if (!file) return;
    setFileName(file.name);
    const r = new FileReader();
    r.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: 'binary' });
      parse(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' }), true);
    };
    r.readAsBinaryString(file);
  }, []);

  const onFile = (e) => handleFile(e.target.files[0]);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.name.endsWith('.csv'))) {
      handleFile(file);
    }
  }, [handleFile]);

  const exportXL = () => {
    if (!data) return;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data.map(r => r.data)), 'Data');
    XLSX.writeFile(wb, fileName ? `reversed_${fileName}` : 'reversed.xlsx');
  };

  const copy = () => {
    if (!data) return;
    navigator.clipboard.writeText(data.map(r => r.data.join('\t')).join('\n'));
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const colLetter = (i) => {
    let s = '';
    while (i >= 0) {
      s = String.fromCharCode((i % 26) + 65) + s;
      i = Math.floor(i / 26) - 1;
    }
    return s;
  };

  const maxCols = data ? Math.max(...data.map(r => r.data.length)) : 0;

  const reset = () => {
    setData(null);
    setFileName('');
    setStats({ input: 0, output: 0, reversed: 0, skipped: 0, splits: 0 });
  };

  return (
    <div className="nr">
      <div className="nr-wrap">
        <header className="nr-header">
          <div className="nr-brand">
            <div className="nr-icon">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M7 16V4m0 0L3 8m4-4l4 4M17 8v12m0 0l4-4m-4 4l-4-4" />
              </svg>
            </div>
            <h1>Number Reverser</h1>
          </div>
          {data && (
            <button onClick={reset} className="nr-btn nr-btn-text">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
                <path d="M3 3v5h5" />
              </svg>
              Reset
            </button>
          )}
        </header>

        {!data ? (
          <main className="nr-main">
            <div className="nr-tabs">
              <button className={`nr-tab ${activeTab === 'upload' ? 'active' : ''}`} onClick={() => setActiveTab('upload')}>
                Upload
              </button>
              <button className={`nr-tab ${activeTab === 'paste' ? 'active' : ''}`} onClick={() => setActiveTab('paste')}>
                Paste
              </button>
            </div>

            {activeTab === 'upload' ? (
              <div
                className={`nr-drop ${isDragging ? 'dragging' : ''}`}
                onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={onDrop}
              >
                <input type="file" accept=".xlsx,.xls,.csv" onChange={onFile} id="file" />
                <label htmlFor="file">
                  <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                    <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z" />
                    <polyline points="14 2 14 8 20 8" />
                    <path d="M12 18v-6m-3 3l3-3 3 3" />
                  </svg>
                  <p><strong>Drop file here</strong> or click to browse</p>
                  <span>.xlsx, .xls, .csv</span>
                </label>
              </div>
            ) : (
              <textarea
                className="nr-paste"
                placeholder="Paste tab-separated data..."
                onChange={e => e.target.value.trim() && parse(e.target.value, false)}
              />
            )}

            <div className="nr-info">
              <h3>How it works</h3>
              <div className="nr-steps">
                <div className="nr-step">
                  <span>1</span>
                  <div>
                    <strong>Detect</strong>
                    <p>Finds h1/v1 markers</p>
                  </div>
                </div>
                <div className="nr-step">
                  <span>2</span>
                  <div>
                    <strong>Split</strong>
                    <p>Handles combined patterns</p>
                  </div>
                </div>
                <div className="nr-step">
                  <span>3</span>
                  <div>
                    <strong>Reverse</strong>
                    <p>Creates W2/S2 rows</p>
                  </div>
                </div>
              </div>
            </div>
          </main>
        ) : (
          <main className="nr-results">
            <div className="nr-stats">
              <div className="nr-stat nr-stat-input">
                <span className="nr-stat-val">{stats.input}</span>
                <span className="nr-stat-lbl">Input</span>
              </div>
              <div className="nr-stat nr-stat-splits">
                <span className="nr-stat-val">{stats.splits}</span>
                <span className="nr-stat-lbl">Splits</span>
              </div>
              <div className="nr-stat nr-stat-reversed">
                <span className="nr-stat-val">{stats.reversed}</span>
                <span className="nr-stat-lbl">Reversed</span>
              </div>
              <div className="nr-stat nr-stat-skipped">
                <span className="nr-stat-val">{stats.skipped}</span>
                <span className="nr-stat-lbl">Skipped</span>
              </div>
              <div className="nr-stat nr-stat-output">
                <span className="nr-stat-val">{stats.output}</span>
                <span className="nr-stat-lbl">Output</span>
              </div>
            </div>

            <div className="nr-table-wrap">
              <div className="nr-table-head">
                <h3>Results</h3>
                <span>{data.length} rows</span>
              </div>
              <div className="nr-table-scroll">
                <table className="nr-table">
                  <thead>
                    <tr>
                      <th>#</th>
                      <th>Status</th>
                      {Array.from({ length: maxCols }, (_, i) => <th key={i}>{colLetter(i)}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {data.map((r, ri) => (
                      <tr key={ri} className={`row-${r.type} ${r.isNew ? 'row-new' : ''}`}>
                        <td className="td-num">{ri + 1}</td>
                        <td>
                          <span className={`tag tag-${r.type} ${r.isNew ? 'tag-new' : ''}`}>
                            {r.isNew && '+ '}
                            {r.type === 'original' ? 'Original' : r.type === 'filled' ? 'Reversed' : 'Unchanged'}
                          </span>
                        </td>
                        {r.data.map((c, ci) => (
                          <td key={ci} className={typeof c === 'number' && r.type === 'filled' ? 'td-hl' : ''}>
                            {c ?? ''}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="nr-actions">
              <button onClick={exportXL} className="nr-btn nr-btn-primary">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <polyline points="7 10 12 15 17 10" />
                  <line x1="12" y1="15" x2="12" y2="3" />
                </svg>
                Download Excel
              </button>
              <button onClick={copy} className={`nr-btn nr-btn-outline ${copied ? 'copied' : ''}`}>
                {copied ? (
                  <>
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <polyline points="20 6 9 17 4 12" />
                    </svg>
                    Copied!
                  </>
                ) : (
                  <>
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="9" y="9" width="13" height="13" rx="2" ry="2" />
                      <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
                    </svg>
                    Copy
                  </>
                )}
              </button>
            </div>
          </main>
        )}

        <footer className="nr-footer">
          Splits combined patterns, reverses odd sequences, skips even counts
        </footer>
      </div>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

        * { margin: 0; padding: 0; box-sizing: border-box; }

        .nr {
          min-height: 100vh;
          background: linear-gradient(180deg, #f7f3ed 0%, #efe8df 100%);
          font-family: 'DM Sans', -apple-system, sans-serif;
          color: #3d3530;
        }

        .nr-wrap {
          max-width: 1000px;
          margin: 0 auto;
          padding: 48px 24px;
          min-height: 100vh;
          display: flex;
          flex-direction: column;
        }

        .nr-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 40px;
        }

        .nr-brand {
          display: flex;
          align-items: center;
          gap: 14px;
        }

        .nr-icon {
          width: 48px;
          height: 48px;
          background: linear-gradient(135deg, #e85d04 0%, #f48c06 50%, #faa307 100%);
          border-radius: 14px;
          display: flex;
          align-items: center;
          justify-content: center;
          color: white;
          box-shadow: 0 8px 24px -8px rgba(232, 93, 4, 0.5);
        }

        .nr-brand h1 {
          font-size: 1.6rem;
          font-weight: 700;
          color: #2a2522;
          letter-spacing: -0.02em;
        }

        .nr-btn {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 12px 22px;
          border: none;
          border-radius: 10px;
          font-size: 0.9rem;
          font-weight: 600;
          font-family: inherit;
          cursor: pointer;
          transition: all 0.2s ease;
        }

        .nr-btn-text {
          background: transparent;
          color: #8a7f76;
          padding: 8px 14px;
        }

        .nr-btn-text:hover {
          color: #3d3530;
          background: rgba(0,0,0,0.05);
        }

        .nr-btn-primary {
          background: linear-gradient(135deg, #e85d04 0%, #f48c06 100%);
          color: white;
          box-shadow: 0 6px 20px -6px rgba(232, 93, 4, 0.5);
        }

        .nr-btn-primary:hover {
          transform: translateY(-2px);
          box-shadow: 0 10px 28px -6px rgba(232, 93, 4, 0.6);
        }

        .nr-btn-outline {
          background: white;
          color: #5c524a;
          border: 2px solid #e0dbd6;
        }

        .nr-btn-outline:hover {
          border-color: #f48c06;
          color: #e85d04;
          background: #fffbf7;
        }

        .nr-btn-outline.copied {
          background: #ecfdf5;
          border-color: #34d399;
          color: #059669;
        }

        .nr-main {
          flex: 1;
          display: flex;
          flex-direction: column;
          gap: 24px;
        }

        .nr-tabs {
          display: flex;
          gap: 4px;
          background: rgba(0,0,0,0.04);
          padding: 5px;
          border-radius: 12px;
          width: fit-content;
        }

        .nr-tab {
          padding: 11px 22px;
          background: transparent;
          border: none;
          border-radius: 9px;
          font-size: 0.9rem;
          font-weight: 600;
          color: #8a7f76;
          cursor: pointer;
          font-family: inherit;
          transition: all 0.2s ease;
        }

        .nr-tab:hover {
          color: #5c524a;
        }

        .nr-tab.active {
          background: white;
          color: #e85d04;
          box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }

        .nr-drop {
          border: 2px dashed #d4ccc3;
          border-radius: 20px;
          padding: 56px 40px;
          text-align: center;
          background: white;
          transition: all 0.25s ease;
          position: relative;
          box-shadow: 0 4px 20px -10px rgba(0,0,0,0.08);
        }

        .nr-drop:hover, .nr-drop.dragging {
          border-color: #f48c06;
          background: linear-gradient(180deg, #fffdfb 0%, #fff8f2 100%);
          box-shadow: 0 8px 32px -10px rgba(244, 140, 6, 0.2);
        }

        .nr-drop input {
          position: absolute;
          inset: 0;
          opacity: 0;
          cursor: pointer;
        }

        .nr-drop label {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 14px;
          cursor: pointer;
          color: #8a7f76;
        }

        .nr-drop svg {
          color: #f48c06;
        }

        .nr-drop p {
          color: #5c524a;
          font-size: 1.05rem;
        }

        .nr-drop span {
          font-size: 0.85rem;
          color: #a9a099;
        }

        .nr-paste {
          width: 100%;
          height: 180px;
          padding: 20px;
          background: white;
          border: 2px solid #e0dbd6;
          border-radius: 16px;
          font-size: 0.9rem;
          font-family: 'DM Mono', monospace;
          color: #3d3530;
          resize: vertical;
          outline: none;
          transition: all 0.2s ease;
          box-shadow: 0 4px 20px -10px rgba(0,0,0,0.08);
        }

        .nr-paste:focus {
          border-color: #f48c06;
          box-shadow: 0 4px 20px -10px rgba(244, 140, 6, 0.25);
        }

        .nr-paste::placeholder {
          color: #b5aea6;
        }

        .nr-info {
          background: white;
          border: 1px solid #e8e2db;
          border-radius: 20px;
          padding: 28px;
          box-shadow: 0 4px 20px -10px rgba(0,0,0,0.08);
        }

        .nr-info h3 {
          font-size: 0.95rem;
          font-weight: 700;
          color: #5c524a;
          margin-bottom: 22px;
        }

        .nr-steps {
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 20px;
        }

        @media (max-width: 600px) {
          .nr-steps { grid-template-columns: 1fr; }
        }

        .nr-step {
          display: flex;
          gap: 14px;
          padding: 16px;
          background: linear-gradient(135deg, #faf7f4 0%, #f5f0eb 100%);
          border-radius: 14px;
        }

        .nr-step span {
          width: 32px;
          height: 32px;
          background: linear-gradient(135deg, #e85d04 0%, #f48c06 100%);
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 700;
          font-size: 0.85rem;
          color: white;
          flex-shrink: 0;
          box-shadow: 0 4px 12px -4px rgba(232, 93, 4, 0.4);
        }

        .nr-step strong {
          display: block;
          font-weight: 600;
          color: #3d3530;
          margin-bottom: 3px;
        }

        .nr-step p {
          font-size: 0.85rem;
          color: #8a7f76;
        }

        .nr-results {
          flex: 1;
          display: flex;
          flex-direction: column;
          gap: 24px;
        }

        .nr-stats {
          display: flex;
          gap: 14px;
          flex-wrap: wrap;
        }

        .nr-stat {
          background: white;
          border-radius: 16px;
          padding: 18px 24px;
          text-align: center;
          min-width: 110px;
          flex: 1;
          box-shadow: 0 4px 16px -8px rgba(0,0,0,0.1);
          border: 2px solid transparent;
          transition: all 0.2s ease;
        }

        .nr-stat:hover {
          transform: translateY(-2px);
        }

        .nr-stat-val {
          display: block;
          font-size: 2rem;
          font-weight: 700;
        }

        .nr-stat-lbl {
          font-size: 0.75rem;
          text-transform: uppercase;
          letter-spacing: 0.06em;
          font-weight: 600;
        }

        /* Input - Slate blue */
        .nr-stat-input {
          border-color: #cbd5e1;
          background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        }
        .nr-stat-input .nr-stat-val { color: #475569; }
        .nr-stat-input .nr-stat-lbl { color: #64748b; }

        /* Splits - Amber/Orange */
        .nr-stat-splits {
          border-color: #fed7aa;
          background: linear-gradient(135deg, #fffbeb 0%, #fef3c7 100%);
        }
        .nr-stat-splits .nr-stat-val { color: #d97706; }
        .nr-stat-splits .nr-stat-lbl { color: #b45309; }

        /* Reversed - Emerald green */
        .nr-stat-reversed {
          border-color: #a7f3d0;
          background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%);
        }
        .nr-stat-reversed .nr-stat-val { color: #059669; }
        .nr-stat-reversed .nr-stat-lbl { color: #047857; }

        /* Skipped - Rose/Red */
        .nr-stat-skipped {
          border-color: #fecaca;
          background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%);
        }
        .nr-stat-skipped .nr-stat-val { color: #dc2626; }
        .nr-stat-skipped .nr-stat-lbl { color: #b91c1c; }

        /* Output - Violet/Purple */
        .nr-stat-output {
          border-color: #ddd6fe;
          background: linear-gradient(135deg, #f5f3ff 0%, #ede9fe 100%);
        }
        .nr-stat-output .nr-stat-val { color: #7c3aed; }
        .nr-stat-output .nr-stat-lbl { color: #6d28d9; }

        .nr-table-wrap {
          background: white;
          border: 1px solid #e8e2db;
          border-radius: 20px;
          overflow: hidden;
          flex: 1;
          box-shadow: 0 4px 20px -10px rgba(0,0,0,0.1);
        }

        .nr-table-head {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 18px 24px;
          border-bottom: 1px solid #e8e2db;
          background: linear-gradient(180deg, #faf8f5 0%, #f5f2ed 100%);
        }

        .nr-table-head h3 {
          font-size: 1.05rem;
          font-weight: 700;
          color: #3d3530;
        }

        .nr-table-head span {
          font-size: 0.85rem;
          color: white;
          background: linear-gradient(135deg, #e85d04 0%, #f48c06 100%);
          padding: 5px 14px;
          border-radius: 20px;
          font-weight: 600;
        }

        .nr-table-scroll {
          overflow: auto;
          max-height: 420px;
        }

        .nr-table {
          width: 100%;
          border-collapse: collapse;
          font-size: 0.85rem;
        }

        .nr-table th {
          padding: 14px 12px;
          background: #faf8f5;
          font-weight: 700;
          color: #5c524a;
          text-align: center;
          position: sticky;
          top: 0;
          z-index: 10;
          border-bottom: 2px solid #e8e2db;
        }

        .nr-table td {
          padding: 12px;
          text-align: center;
          border-bottom: 1px solid #f0ebe6;
          font-family: 'DM Mono', monospace;
          color: #5c524a;
        }

        .td-num {
          color: #b5aea6;
          font-weight: 500;
        }

        .td-hl {
          color: #059669 !important;
          font-weight: 600;
          background: rgba(16, 185, 129, 0.06);
        }

        .row-original { background: #fdfcfb; }
        .row-filled { background: #fafdf9; }
        .row-new { background: #fffcf5; }

        .tag {
          display: inline-block;
          padding: 5px 12px;
          border-radius: 8px;
          font-size: 0.75rem;
          font-weight: 600;
          font-family: 'DM Sans', sans-serif;
        }

        .tag-original {
          background: #f1f5f9;
          color: #475569;
        }

        .tag-filled {
          background: #ecfdf5;
          color: #059669;
        }

        .tag-unchanged {
          background: #f5f0eb;
          color: #8a7f76;
        }

        .tag-new {
          background: #fef3c7;
          color: #d97706;
        }

        .nr-actions {
          display: flex;
          gap: 14px;
          justify-content: center;
        }

        .nr-footer {
          text-align: center;
          padding: 36px 0 0;
          color: #a9a099;
          font-size: 0.85rem;
          margin-top: auto;
        }

        ::-webkit-scrollbar {
          width: 10px;
          height: 10px;
        }

        ::-webkit-scrollbar-track {
          background: #f5f0eb;
        }

        ::-webkit-scrollbar-thumb {
          background: #d4ccc3;
          border-radius: 5px;
          border: 2px solid #f5f0eb;
        }

        ::-webkit-scrollbar-thumb:hover {
          background: #c9c0b5;
        }
      `}</style>
    </div>
  );
}
