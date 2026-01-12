import React, { useState, useCallback, useEffect } from 'react';
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
    if (typeof cell === 'number') return Math.floor(cell);
    if (typeof cell === 'string') {
      const m = cell.trim().match(P.num);
      return m ? Math.floor(parseFloat(m[1])) : null;
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
    <div className="nr-container">
      {/* Animated background */}
      <div className="nr-bg">
        <div className="nr-orb nr-orb-1" />
        <div className="nr-orb nr-orb-2" />
        <div className="nr-orb nr-orb-3" />
        <div className="nr-grid" />
      </div>

      <div className="nr-wrapper">
        {/* Header */}
        <header className="nr-header">
          <div className="nr-logo">
            <div className="nr-logo-icon">
              <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <path d="M7 16V4m0 0L3 8m4-4l4 4M17 8v12m0 0l4-4m-4 4l-4-4" />
              </svg>
            </div>
            <div className="nr-logo-text">
              <h1>Number Reverser</h1>
              <p>Transform Excel data with precision</p>
            </div>
          </div>
          {data && (
            <button onClick={reset} className="nr-btn nr-btn-ghost">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
                <path d="M3 3v5h5" />
              </svg>
              Start Over
            </button>
          )}
        </header>

        {!data ? (
          <div className="nr-input-section">
            {/* Tabs */}
            <div className="nr-tabs">
              <button
                className={`nr-tab ${activeTab === 'upload' ? 'active' : ''}`}
                onClick={() => setActiveTab('upload')}
              >
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <polyline points="17 8 12 3 7 8" />
                  <line x1="12" y1="3" x2="12" y2="15" />
                </svg>
                Upload File
              </button>
              <button
                className={`nr-tab ${activeTab === 'paste' ? 'active' : ''}`}
                onClick={() => setActiveTab('paste')}
              >
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" />
                  <rect x="8" y="2" width="8" height="4" rx="1" ry="1" />
                </svg>
                Paste Data
              </button>
            </div>

            {/* Upload Area */}
            {activeTab === 'upload' && (
              <div
                className={`nr-dropzone ${isDragging ? 'dragging' : ''}`}
                onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={onDrop}
              >
                <input type="file" accept=".xlsx,.xls,.csv" onChange={onFile} id="fileInput" />
                <label htmlFor="fileInput">
                  <div className="nr-dropzone-icon">
                    <svg width="56" height="56" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                      <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z" />
                      <polyline points="14 2 14 8 20 8" />
                      <path d="M12 18v-6m-3 3l3-3 3 3" />
                    </svg>
                  </div>
                  <p className="nr-dropzone-text">
                    <span>Click to upload</span> or drag and drop
                  </p>
                  <p className="nr-dropzone-hint">Excel (.xlsx, .xls) or CSV files</p>
                </label>
              </div>
            )}

            {/* Paste Area */}
            {activeTab === 'paste' && (
              <textarea
                className="nr-textarea"
                placeholder="Paste your tab-separated data here..."
                onChange={e => e.target.value.trim() && parse(e.target.value, false)}
              />
            )}

            {/* Info Card */}
            <div className="nr-card nr-info-card">
              <h3>
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10" />
                  <path d="M12 16v-4M12 8h.01" />
                </svg>
                How it works
              </h3>
              <div className="nr-steps">
                <div className="nr-step">
                  <div className="nr-step-num">1</div>
                  <div>
                    <strong>Detect Patterns</strong>
                    <p>Finds h1/v1 markers and combined patterns</p>
                  </div>
                </div>
                <div className="nr-step">
                  <div className="nr-step-num">2</div>
                  <div>
                    <strong>Split & Process</strong>
                    <p>Splits combined rows, processes odd sequences</p>
                  </div>
                </div>
                <div className="nr-step">
                  <div className="nr-step-num">3</div>
                  <div>
                    <strong>Reverse Numbers</strong>
                    <p>Creates W2/S2 rows with reversed values</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        ) : (
          <div className="nr-results">
            {/* Stats */}
            <div className="nr-stats">
              {[
                { label: 'Input', value: stats.input, color: 'blue' },
                { label: 'Splits', value: stats.splits, color: 'amber' },
                { label: 'Reversed', value: stats.reversed, color: 'green' },
                { label: 'Skipped', value: stats.skipped, color: 'red' },
                { label: 'Output', value: stats.output, color: 'purple' }
              ].map((stat) => (
                <div key={stat.label} className={`nr-stat nr-stat-${stat.color}`}>
                  <div className="nr-stat-value">{stat.value}</div>
                  <div className="nr-stat-label">{stat.label}</div>
                </div>
              ))}
            </div>

            {/* Table */}
            <div className="nr-card nr-table-card">
              <div className="nr-table-header">
                <h3>
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <rect x="3" y="3" width="18" height="18" rx="2" ry="2" />
                    <line x1="3" y1="9" x2="21" y2="9" />
                    <line x1="9" y1="21" x2="9" y2="9" />
                  </svg>
                  Processed Data
                </h3>
                <span className="nr-badge">{data.length} rows</span>
              </div>
              <div className="nr-table-wrap">
                <table className="nr-table">
                  <thead>
                    <tr>
                      <th className="nr-th-num">#</th>
                      <th className="nr-th-status">Status</th>
                      {Array.from({ length: maxCols }, (_, i) => (
                        <th key={i}>{colLetter(i)}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {data.map((r, ri) => (
                      <tr key={ri} className={`nr-row nr-row-${r.type} ${r.isNew ? 'nr-row-new' : ''}`}>
                        <td className="nr-td-num">{ri + 1}</td>
                        <td className="nr-td-status">
                          <span className={`nr-status nr-status-${r.type} ${r.isNew ? 'nr-status-new' : ''}`}>
                            {r.isNew && <span className="nr-sparkle">+</span>}
                            {r.type === 'original' ? 'Original' : r.type === 'filled' ? 'Reversed' : 'Unchanged'}
                          </span>
                        </td>
                        {r.data.map((c, ci) => (
                          <td key={ci} className={typeof c === 'number' && r.type === 'filled' ? 'nr-td-highlight' : ''}>
                            {c ?? ''}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Actions */}
            <div className="nr-actions">
              <button onClick={exportXL} className="nr-btn nr-btn-primary">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <polyline points="7 10 12 15 17 10" />
                  <line x1="12" y1="15" x2="12" y2="3" />
                </svg>
                Download Excel
              </button>
              <button onClick={copy} className={`nr-btn nr-btn-secondary ${copied ? 'copied' : ''}`}>
                {copied ? (
                  <>
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <polyline points="20 6 9 17 4 12" />
                    </svg>
                    Copied!
                  </>
                ) : (
                  <>
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="9" y="9" width="13" height="13" rx="2" ry="2" />
                      <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
                    </svg>
                    Copy to Clipboard
                  </>
                )}
              </button>
            </div>
          </div>
        )}

        <footer className="nr-footer">
          <p>Splits combined patterns &bull; Reverses odd sequences &bull; Skips even counts</p>
        </footer>
      </div>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

        * { margin: 0; padding: 0; box-sizing: border-box; }

        .nr-container {
          min-height: 100vh;
          background: linear-gradient(135deg, #0c0f1a 0%, #1a1f35 50%, #0f1629 100%);
          font-family: 'Inter', -apple-system, sans-serif;
          color: #e2e8f0;
          position: relative;
          overflow-x: hidden;
        }

        .nr-bg {
          position: fixed;
          inset: 0;
          pointer-events: none;
          overflow: hidden;
        }

        .nr-orb {
          position: absolute;
          border-radius: 50%;
          filter: blur(80px);
          opacity: 0.6;
          animation: float 20s ease-in-out infinite;
        }

        .nr-orb-1 {
          top: -15%;
          right: -10%;
          width: 600px;
          height: 600px;
          background: radial-gradient(circle, rgba(99, 102, 241, 0.3) 0%, transparent 70%);
        }

        .nr-orb-2 {
          bottom: -20%;
          left: -15%;
          width: 500px;
          height: 500px;
          background: radial-gradient(circle, rgba(16, 185, 129, 0.25) 0%, transparent 70%);
          animation-delay: -10s;
          animation-direction: reverse;
        }

        .nr-orb-3 {
          top: 40%;
          left: 30%;
          width: 400px;
          height: 400px;
          background: radial-gradient(circle, rgba(139, 92, 246, 0.2) 0%, transparent 70%);
          animation-delay: -5s;
        }

        .nr-grid {
          position: absolute;
          inset: 0;
          background-image:
            linear-gradient(rgba(255,255,255,0.02) 1px, transparent 1px),
            linear-gradient(90deg, rgba(255,255,255,0.02) 1px, transparent 1px);
          background-size: 60px 60px;
          mask-image: radial-gradient(ellipse at center, black 0%, transparent 70%);
        }

        @keyframes float {
          0%, 100% { transform: translate(0, 0) scale(1); }
          33% { transform: translate(30px, -30px) scale(1.05); }
          66% { transform: translate(-20px, 20px) scale(0.95); }
        }

        .nr-wrapper {
          position: relative;
          z-index: 1;
          max-width: 1200px;
          margin: 0 auto;
          padding: 40px 24px;
          min-height: 100vh;
          display: flex;
          flex-direction: column;
        }

        .nr-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 48px;
          flex-wrap: wrap;
          gap: 20px;
        }

        .nr-logo {
          display: flex;
          align-items: center;
          gap: 16px;
        }

        .nr-logo-icon {
          width: 56px;
          height: 56px;
          background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #a855f7 100%);
          border-radius: 16px;
          display: flex;
          align-items: center;
          justify-content: center;
          color: white;
          box-shadow:
            0 0 0 1px rgba(255,255,255,0.1),
            0 20px 40px -10px rgba(99, 102, 241, 0.5),
            0 0 60px -10px rgba(139, 92, 246, 0.3);
          animation: pulse 3s ease-in-out infinite;
        }

        @keyframes pulse {
          0%, 100% { box-shadow: 0 0 0 1px rgba(255,255,255,0.1), 0 20px 40px -10px rgba(99, 102, 241, 0.5), 0 0 60px -10px rgba(139, 92, 246, 0.3); }
          50% { box-shadow: 0 0 0 1px rgba(255,255,255,0.15), 0 25px 50px -10px rgba(99, 102, 241, 0.6), 0 0 80px -10px rgba(139, 92, 246, 0.4); }
        }

        .nr-logo-text h1 {
          font-size: 1.75rem;
          font-weight: 700;
          letter-spacing: -0.03em;
          background: linear-gradient(135deg, #f8fafc 0%, #cbd5e1 100%);
          -webkit-background-clip: text;
          -webkit-text-fill-color: transparent;
          background-clip: text;
        }

        .nr-logo-text p {
          font-size: 0.9rem;
          color: #64748b;
          margin-top: 2px;
        }

        .nr-btn {
          display: inline-flex;
          align-items: center;
          gap: 10px;
          padding: 14px 28px;
          border: none;
          border-radius: 12px;
          font-size: 0.95rem;
          font-weight: 600;
          font-family: inherit;
          cursor: pointer;
          transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .nr-btn-ghost {
          background: rgba(255,255,255,0.05);
          color: #94a3b8;
          border: 1px solid rgba(255,255,255,0.08);
          backdrop-filter: blur(10px);
        }

        .nr-btn-ghost:hover {
          background: rgba(255,255,255,0.1);
          color: #e2e8f0;
          transform: translateY(-2px);
        }

        .nr-btn-primary {
          background: linear-gradient(135deg, #10b981 0%, #059669 100%);
          color: white;
          box-shadow: 0 10px 40px -10px rgba(16, 185, 129, 0.5);
        }

        .nr-btn-primary:hover {
          transform: translateY(-3px);
          box-shadow: 0 20px 50px -10px rgba(16, 185, 129, 0.6);
        }

        .nr-btn-secondary {
          background: rgba(99, 102, 241, 0.1);
          color: #a5b4fc;
          border: 1px solid rgba(99, 102, 241, 0.3);
        }

        .nr-btn-secondary:hover {
          background: rgba(99, 102, 241, 0.2);
          transform: translateY(-2px);
        }

        .nr-btn-secondary.copied {
          background: rgba(16, 185, 129, 0.2);
          color: #34d399;
          border-color: rgba(16, 185, 129, 0.3);
        }

        .nr-input-section {
          display: flex;
          flex-direction: column;
          gap: 28px;
          flex: 1;
        }

        .nr-tabs {
          display: flex;
          gap: 8px;
          padding: 6px;
          background: rgba(255,255,255,0.03);
          border-radius: 14px;
          width: fit-content;
          border: 1px solid rgba(255,255,255,0.05);
        }

        .nr-tab {
          display: flex;
          align-items: center;
          gap: 10px;
          padding: 12px 24px;
          background: transparent;
          border: none;
          border-radius: 10px;
          color: #64748b;
          font-size: 0.95rem;
          font-weight: 500;
          font-family: inherit;
          cursor: pointer;
          transition: all 0.25s ease;
        }

        .nr-tab:hover {
          color: #94a3b8;
        }

        .nr-tab.active {
          background: linear-gradient(135deg, rgba(99, 102, 241, 0.2) 0%, rgba(139, 92, 246, 0.15) 100%);
          color: #c7d2fe;
          box-shadow: 0 4px 20px -5px rgba(99, 102, 241, 0.3);
        }

        .nr-dropzone {
          position: relative;
          border: 2px dashed rgba(99, 102, 241, 0.3);
          border-radius: 24px;
          padding: 60px 40px;
          text-align: center;
          background: linear-gradient(135deg, rgba(99, 102, 241, 0.05) 0%, rgba(139, 92, 246, 0.03) 100%);
          transition: all 0.3s ease;
          cursor: pointer;
        }

        .nr-dropzone:hover, .nr-dropzone.dragging {
          border-color: #6366f1;
          background: linear-gradient(135deg, rgba(99, 102, 241, 0.1) 0%, rgba(139, 92, 246, 0.08) 100%);
          transform: scale(1.01);
          box-shadow: 0 20px 60px -20px rgba(99, 102, 241, 0.3);
        }

        .nr-dropzone input {
          position: absolute;
          inset: 0;
          opacity: 0;
          cursor: pointer;
        }

        .nr-dropzone label {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 20px;
          cursor: pointer;
        }

        .nr-dropzone-icon {
          width: 100px;
          height: 100px;
          background: linear-gradient(135deg, rgba(99, 102, 241, 0.2) 0%, rgba(139, 92, 246, 0.15) 100%);
          border-radius: 24px;
          display: flex;
          align-items: center;
          justify-content: center;
          color: #a5b4fc;
          border: 1px solid rgba(99, 102, 241, 0.2);
        }

        .nr-dropzone-text {
          font-size: 1.1rem;
          color: #94a3b8;
        }

        .nr-dropzone-text span {
          color: #a5b4fc;
          font-weight: 600;
        }

        .nr-dropzone-hint {
          font-size: 0.85rem;
          color: #475569;
        }

        .nr-textarea {
          width: 100%;
          height: 220px;
          padding: 24px;
          background: rgba(15, 23, 42, 0.6);
          border: 2px solid rgba(99, 102, 241, 0.2);
          border-radius: 20px;
          color: #e2e8f0;
          font-size: 0.95rem;
          font-family: 'JetBrains Mono', monospace;
          resize: vertical;
          outline: none;
          transition: all 0.3s ease;
          backdrop-filter: blur(10px);
        }

        .nr-textarea:focus {
          border-color: #6366f1;
          box-shadow: 0 0 0 4px rgba(99, 102, 241, 0.1);
        }

        .nr-textarea::placeholder {
          color: #475569;
        }

        .nr-card {
          background: rgba(255,255,255,0.03);
          border: 1px solid rgba(255,255,255,0.06);
          border-radius: 24px;
          backdrop-filter: blur(20px);
        }

        .nr-info-card {
          padding: 32px;
        }

        .nr-info-card h3 {
          display: flex;
          align-items: center;
          gap: 12px;
          font-size: 1.1rem;
          font-weight: 600;
          color: #e2e8f0;
          margin-bottom: 28px;
        }

        .nr-steps {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
          gap: 24px;
        }

        .nr-step {
          display: flex;
          gap: 16px;
          align-items: flex-start;
        }

        .nr-step-num {
          width: 36px;
          height: 36px;
          background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          color: white;
          font-weight: 700;
          font-size: 0.9rem;
          flex-shrink: 0;
          box-shadow: 0 8px 20px -8px rgba(99, 102, 241, 0.5);
        }

        .nr-step strong {
          display: block;
          color: #e2e8f0;
          margin-bottom: 4px;
        }

        .nr-step p {
          color: #64748b;
          font-size: 0.85rem;
          line-height: 1.5;
        }

        .nr-results {
          display: flex;
          flex-direction: column;
          gap: 28px;
          flex: 1;
        }

        .nr-stats {
          display: grid;
          grid-template-columns: repeat(5, 1fr);
          gap: 16px;
        }

        @media (max-width: 768px) {
          .nr-stats { grid-template-columns: repeat(3, 1fr); }
        }

        @media (max-width: 480px) {
          .nr-stats { grid-template-columns: repeat(2, 1fr); }
        }

        .nr-stat {
          background: rgba(255,255,255,0.03);
          border: 1px solid rgba(255,255,255,0.06);
          border-radius: 16px;
          padding: 20px 16px;
          text-align: center;
          transition: all 0.3s ease;
        }

        .nr-stat:hover {
          transform: translateY(-4px);
          box-shadow: 0 20px 40px -20px rgba(0,0,0,0.3);
        }

        .nr-stat-value {
          font-size: 2rem;
          font-weight: 700;
          margin-bottom: 4px;
        }

        .nr-stat-label {
          font-size: 0.75rem;
          text-transform: uppercase;
          letter-spacing: 0.08em;
          color: #64748b;
        }

        .nr-stat-blue .nr-stat-value { color: #818cf8; }
        .nr-stat-amber .nr-stat-value { color: #fbbf24; }
        .nr-stat-green .nr-stat-value { color: #34d399; }
        .nr-stat-red .nr-stat-value { color: #f87171; }
        .nr-stat-purple .nr-stat-value { color: #c084fc; }

        .nr-stat-blue { border-color: rgba(99, 102, 241, 0.2); }
        .nr-stat-amber { border-color: rgba(245, 158, 11, 0.2); }
        .nr-stat-green { border-color: rgba(16, 185, 129, 0.2); }
        .nr-stat-red { border-color: rgba(239, 68, 68, 0.2); }
        .nr-stat-purple { border-color: rgba(168, 85, 247, 0.2); }

        .nr-table-card {
          overflow: hidden;
          flex: 1;
          display: flex;
          flex-direction: column;
        }

        .nr-table-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 20px 28px;
          border-bottom: 1px solid rgba(255,255,255,0.06);
        }

        .nr-table-header h3 {
          display: flex;
          align-items: center;
          gap: 12px;
          font-size: 1rem;
          font-weight: 600;
        }

        .nr-badge {
          font-size: 0.8rem;
          color: #94a3b8;
          background: rgba(255,255,255,0.05);
          padding: 6px 14px;
          border-radius: 20px;
          border: 1px solid rgba(255,255,255,0.06);
        }

        .nr-table-wrap {
          overflow: auto;
          max-height: 450px;
          flex: 1;
        }

        .nr-table {
          width: 100%;
          border-collapse: collapse;
          font-size: 0.85rem;
        }

        .nr-table th {
          padding: 14px 12px;
          background: rgba(99, 102, 241, 0.08);
          color: #a5b4fc;
          font-weight: 600;
          text-align: center;
          position: sticky;
          top: 0;
          z-index: 10;
          min-width: 60px;
          border-bottom: 1px solid rgba(99, 102, 241, 0.15);
        }

        .nr-th-num { min-width: 50px; }
        .nr-th-status { min-width: 100px; }

        .nr-table td {
          padding: 12px;
          text-align: center;
          border-bottom: 1px solid rgba(255,255,255,0.04);
          font-family: 'JetBrains Mono', monospace;
          color: #cbd5e1;
        }

        .nr-td-num {
          color: #475569;
          font-weight: 500;
        }

        .nr-td-highlight {
          color: #34d399 !important;
          font-weight: 600;
        }

        .nr-row-original { background: rgba(99, 102, 241, 0.04); }
        .nr-row-filled { background: rgba(16, 185, 129, 0.04); }
        .nr-row-new { background: rgba(245, 158, 11, 0.06); }

        .nr-status {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 5px 12px;
          border-radius: 6px;
          font-size: 0.75rem;
          font-weight: 500;
          font-family: 'Inter', sans-serif;
        }

        .nr-status-original {
          background: rgba(99, 102, 241, 0.15);
          color: #a5b4fc;
          border: 1px solid rgba(99, 102, 241, 0.25);
        }

        .nr-status-filled {
          background: rgba(16, 185, 129, 0.15);
          color: #34d399;
          border: 1px solid rgba(16, 185, 129, 0.25);
        }

        .nr-status-unchanged {
          background: rgba(100, 116, 139, 0.15);
          color: #94a3b8;
          border: 1px solid rgba(100, 116, 139, 0.25);
        }

        .nr-status-new {
          background: rgba(245, 158, 11, 0.15) !important;
          color: #fbbf24 !important;
          border: 1px solid rgba(245, 158, 11, 0.25) !important;
        }

        .nr-sparkle {
          font-weight: 700;
        }

        .nr-actions {
          display: flex;
          gap: 16px;
          justify-content: center;
          flex-wrap: wrap;
        }

        .nr-footer {
          text-align: center;
          padding: 40px 0 20px;
          color: #475569;
          font-size: 0.85rem;
          margin-top: auto;
        }

        /* Scrollbar */
        ::-webkit-scrollbar {
          width: 8px;
          height: 8px;
        }

        ::-webkit-scrollbar-track {
          background: rgba(255,255,255,0.02);
        }

        ::-webkit-scrollbar-thumb {
          background: rgba(255,255,255,0.1);
          border-radius: 4px;
        }

        ::-webkit-scrollbar-thumb:hover {
          background: rgba(255,255,255,0.15);
        }
      `}</style>
    </div>
  );
}
