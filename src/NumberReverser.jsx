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

  const getNums = (row, start, end = row.length) => {
    const nums = [];
    for (let i = start + 1; i < end && i < row.length; i++) {
      const n = extractNum(row[i]);
      if (n !== null) nums.push(n);
    }
    return nums;
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
        const nums1 = [...getNums(row, m1Idx, combIdx), extractNum(row[combIdx])];
        const nums2 = getNums(row, combIdx);

        if (nums1.length % 2 === 1 && nums2.length > 0 && nums2.length % 2 === 1) {
          splits++;
          const r1 = [...row];
          r1[combIdx] = extractNum(row[combIdx]);
          for (let j = combIdx + 1; j < r1.length; j++) r1[j] = '';
          result.push({ data: r1, type: 'original' });

          if (hasNext) {
            const r2 = buildRow(input[++i].length, `${getM2(letter)}2`, m1Idx, impIdx, ':', [...nums1].reverse());
            result.push({ data: r2, type: 'filled' });
            reversed++;
          }

          const combLetter = getCombinedLetter(row[combIdx]);
          result.push({ data: buildRow(row.length, `${combLetter} 1`, m1Idx, impIdx, 'Imposts:', nums2), type: 'original', isNew: true });
          result.push({ data: buildRow(row.length, `${getM2(combLetter)}2`, m1Idx, impIdx, ':', [...nums2].reverse()), type: 'filled', isNew: true });
          reversed++;
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

    setData(result);
    setStats({ input: input.length, output: result.length, reversed, skipped, splits });
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
              <div className="nr-stat">
                <span className="nr-stat-val">{stats.input}</span>
                <span className="nr-stat-lbl">Input</span>
              </div>
              <div className="nr-stat">
                <span className="nr-stat-val">{stats.splits}</span>
                <span className="nr-stat-lbl">Splits</span>
              </div>
              <div className="nr-stat">
                <span className="nr-stat-val">{stats.reversed}</span>
                <span className="nr-stat-lbl">Reversed</span>
              </div>
              <div className="nr-stat">
                <span className="nr-stat-val">{stats.skipped}</span>
                <span className="nr-stat-lbl">Skipped</span>
              </div>
              <div className="nr-stat">
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
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=DM+Mono&display=swap');

        * { margin: 0; padding: 0; box-sizing: border-box; }

        .nr {
          min-height: 100vh;
          background: #faf8f5;
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
          gap: 12px;
        }

        .nr-icon {
          width: 44px;
          height: 44px;
          background: #c9795a;
          border-radius: 12px;
          display: flex;
          align-items: center;
          justify-content: center;
          color: white;
        }

        .nr-brand h1 {
          font-size: 1.5rem;
          font-weight: 600;
          color: #2a2522;
        }

        .nr-btn {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 12px 20px;
          border: none;
          border-radius: 10px;
          font-size: 0.9rem;
          font-weight: 500;
          font-family: inherit;
          cursor: pointer;
          transition: all 0.2s ease;
        }

        .nr-btn-text {
          background: transparent;
          color: #8a7f76;
          padding: 8px 12px;
        }

        .nr-btn-text:hover {
          color: #3d3530;
          background: rgba(0,0,0,0.04);
        }

        .nr-btn-primary {
          background: #c9795a;
          color: white;
        }

        .nr-btn-primary:hover {
          background: #b56a4c;
        }

        .nr-btn-outline {
          background: white;
          color: #5c524a;
          border: 1px solid #e0dbd6;
        }

        .nr-btn-outline:hover {
          border-color: #c9795a;
          color: #c9795a;
        }

        .nr-btn-outline.copied {
          background: #e8f5e9;
          border-color: #81c784;
          color: #4caf50;
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
          background: #f0ebe6;
          padding: 4px;
          border-radius: 10px;
          width: fit-content;
        }

        .nr-tab {
          padding: 10px 20px;
          background: transparent;
          border: none;
          border-radius: 8px;
          font-size: 0.9rem;
          font-weight: 500;
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
          color: #3d3530;
          box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }

        .nr-drop {
          border: 2px dashed #d9d4ce;
          border-radius: 16px;
          padding: 48px;
          text-align: center;
          background: white;
          transition: all 0.2s ease;
          position: relative;
        }

        .nr-drop:hover, .nr-drop.dragging {
          border-color: #c9795a;
          background: #fdfcfb;
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
          gap: 12px;
          cursor: pointer;
          color: #8a7f76;
        }

        .nr-drop svg {
          color: #c9795a;
        }

        .nr-drop p {
          color: #5c524a;
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
          border: 1px solid #e0dbd6;
          border-radius: 12px;
          font-size: 0.9rem;
          font-family: 'DM Mono', monospace;
          color: #3d3530;
          resize: vertical;
          outline: none;
          transition: border-color 0.2s ease;
        }

        .nr-paste:focus {
          border-color: #c9795a;
        }

        .nr-paste::placeholder {
          color: #b5aea6;
        }

        .nr-info {
          background: white;
          border: 1px solid #e0dbd6;
          border-radius: 16px;
          padding: 24px;
        }

        .nr-info h3 {
          font-size: 0.9rem;
          font-weight: 600;
          color: #5c524a;
          margin-bottom: 20px;
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
          gap: 12px;
        }

        .nr-step span {
          width: 28px;
          height: 28px;
          background: #f5f0eb;
          border-radius: 8px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 600;
          font-size: 0.85rem;
          color: #c9795a;
          flex-shrink: 0;
        }

        .nr-step strong {
          display: block;
          font-weight: 600;
          color: #3d3530;
          margin-bottom: 2px;
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
          gap: 12px;
          flex-wrap: wrap;
        }

        .nr-stat {
          background: white;
          border: 1px solid #e0dbd6;
          border-radius: 12px;
          padding: 16px 24px;
          text-align: center;
          min-width: 100px;
          flex: 1;
        }

        .nr-stat-val {
          display: block;
          font-size: 1.75rem;
          font-weight: 600;
          color: #c9795a;
        }

        .nr-stat-lbl {
          font-size: 0.75rem;
          color: #8a7f76;
          text-transform: uppercase;
          letter-spacing: 0.05em;
        }

        .nr-table-wrap {
          background: white;
          border: 1px solid #e0dbd6;
          border-radius: 16px;
          overflow: hidden;
          flex: 1;
        }

        .nr-table-head {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 16px 20px;
          border-bottom: 1px solid #e0dbd6;
        }

        .nr-table-head h3 {
          font-size: 1rem;
          font-weight: 600;
        }

        .nr-table-head span {
          font-size: 0.85rem;
          color: #8a7f76;
          background: #f5f0eb;
          padding: 4px 12px;
          border-radius: 20px;
        }

        .nr-table-scroll {
          overflow: auto;
          max-height: 400px;
        }

        .nr-table {
          width: 100%;
          border-collapse: collapse;
          font-size: 0.85rem;
        }

        .nr-table th {
          padding: 12px;
          background: #faf8f5;
          font-weight: 600;
          color: #5c524a;
          text-align: center;
          position: sticky;
          top: 0;
          z-index: 10;
          border-bottom: 1px solid #e0dbd6;
        }

        .nr-table td {
          padding: 10px 12px;
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
          color: #c9795a !important;
          font-weight: 600;
        }

        .row-original { background: #fdfcfb; }
        .row-filled { background: #faf8f5; }
        .row-new { background: #fef9f6; }

        .tag {
          display: inline-block;
          padding: 4px 10px;
          border-radius: 6px;
          font-size: 0.75rem;
          font-weight: 500;
          font-family: 'DM Sans', sans-serif;
        }

        .tag-original {
          background: #f5f0eb;
          color: #8a7f76;
        }

        .tag-filled {
          background: #fef3ed;
          color: #c9795a;
        }

        .tag-unchanged {
          background: #f0ebe6;
          color: #a9a099;
        }

        .tag-new {
          background: #fff4e5;
          color: #e07c24;
        }

        .nr-actions {
          display: flex;
          gap: 12px;
          justify-content: center;
        }

        .nr-footer {
          text-align: center;
          padding: 32px 0 0;
          color: #a9a099;
          font-size: 0.85rem;
          margin-top: auto;
        }

        ::-webkit-scrollbar {
          width: 8px;
          height: 8px;
        }

        ::-webkit-scrollbar-track {
          background: #f5f0eb;
        }

        ::-webkit-scrollbar-thumb {
          background: #d9d4ce;
          border-radius: 4px;
        }

        ::-webkit-scrollbar-thumb:hover {
          background: #c9c3bc;
        }
      `}</style>
    </div>
  );
}
