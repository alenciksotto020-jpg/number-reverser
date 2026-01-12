import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';

export default function NumberReverser() {
  const [data, setData] = useState(null);
  const [fileName, setFileName] = useState('');
  const [stats, setStats] = useState({ input: 0, output: 0, reversed: 0, skipped: 0, splits: 0 });
  const [isDragging, setIsDragging] = useState(false);
  const [activeTab, setActiveTab] = useState('upload');

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
      const m = cell.match(P.num);
      return m ? Math.floor(parseFloat(m[1])) : null;
    }
    return null;
  };
  const getMarkerLetter = (cell) => cell?.trim().match(/^([hv])/i)?.[1].toLowerCase();
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
          for (let j = combIdx + 1; j < r1.length; j++) r1[j] = '';
          result.push({ data: r1, type: 'original' });

          if (hasNext) {
            const r2 = buildRow(input[++i].length, `${getM2(letter)}2`, m1Idx, impIdx, ':', [...nums1].reverse());
            result.push({ data: r2, type: 'filled' });
            reversed++;
          }

          const newLetter = altMarker(letter);
          result.push({ data: buildRow(row.length, `${newLetter} 1`, m1Idx, impIdx, 'Imposts:', nums2), type: 'original', isNew: true });
          result.push({ data: buildRow(row.length, `${getM2(newLetter)}2`, m1Idx, impIdx, ':', [...nums2].reverse()), type: 'filled', isNew: true });
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

  const copy = () => data && navigator.clipboard.writeText(data.map(r => r.data.join('\t')).join('\n'));

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
    <div style={styles.container}>
      {/* Animated background elements */}
      <div style={styles.bgOrb1} />
      <div style={styles.bgOrb2} />
      <div style={styles.bgOrb3} />

      <div style={styles.wrapper}>
        {/* Header */}
        <header style={styles.header}>
          <div style={styles.logoContainer}>
            <div style={styles.logoIcon}>
              <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M7 16V4m0 0L3 8m4-4l4 4M17 8v12m0 0l4-4m-4 4l-4-4" />
              </svg>
            </div>
            <div>
              <h1 style={styles.title}>Number Reverser</h1>
              <p style={styles.subtitle}>Transform your Excel data with precision</p>
            </div>
          </div>
          {data && (
            <button onClick={reset} style={styles.resetButton}>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
                <path d="M3 3v5h5" />
              </svg>
              Start Over
            </button>
          )}
        </header>

        {!data ? (
          /* Input Section */
          <div style={styles.inputSection}>
            {/* Tab Navigation */}
            <div style={styles.tabContainer}>
              <button
                style={{ ...styles.tab, ...(activeTab === 'upload' ? styles.activeTab : {}) }}
                onClick={() => setActiveTab('upload')}
              >
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <polyline points="17 8 12 3 7 8" />
                  <line x1="12" y1="3" x2="12" y2="15" />
                </svg>
                Upload File
              </button>
              <button
                style={{ ...styles.tab, ...(activeTab === 'paste' ? styles.activeTab : {}) }}
                onClick={() => setActiveTab('paste')}
              >
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" />
                  <rect x="8" y="2" width="8" height="4" rx="1" ry="1" />
                </svg>
                Paste Data
              </button>
            </div>

            {/* Upload Area */}
            {activeTab === 'upload' && (
              <div
                style={{
                  ...styles.dropZone,
                  ...(isDragging ? styles.dropZoneActive : {})
                }}
                onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={onDrop}
              >
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  onChange={onFile}
                  style={styles.fileInput}
                  id="fileInput"
                />
                <label htmlFor="fileInput" style={styles.dropZoneContent}>
                  <div style={styles.uploadIconContainer}>
                    <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                      <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z" />
                      <polyline points="14 2 14 8 20 8" />
                      <line x1="12" y1="18" x2="12" y2="12" />
                      <line x1="9" y1="15" x2="12" y2="12" />
                      <line x1="15" y1="15" x2="12" y2="12" />
                    </svg>
                  </div>
                  <p style={styles.dropText}>
                    <span style={styles.dropTextHighlight}>Click to upload</span> or drag and drop
                  </p>
                  <p style={styles.dropSubtext}>Excel files (.xlsx, .xls) or CSV</p>
                </label>
              </div>
            )}

            {/* Paste Area */}
            {activeTab === 'paste' && (
              <div style={styles.pasteContainer}>
                <textarea
                  placeholder="Paste your tab-separated data here...&#10;&#10;Example:&#10;Imposts:  v 1  53  1364  53.0+v1  53  184  66&#10;          S2"
                  onChange={e => e.target.value.trim() && parse(e.target.value, false)}
                  style={styles.textarea}
                />
              </div>
            )}

            {/* How it works */}
            <div style={styles.infoCard}>
              <h3 style={styles.infoTitle}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10" />
                  <path d="M12 16v-4" />
                  <path d="M12 8h.01" />
                </svg>
                How it works
              </h3>
              <div style={styles.infoGrid}>
                <div style={styles.infoItem}>
                  <div style={styles.infoNumber}>1</div>
                  <div>
                    <strong style={styles.infoItemTitle}>Detect Patterns</strong>
                    <p style={styles.infoItemText}>Finds h1/v1 markers and combined patterns like "53.0+v1"</p>
                  </div>
                </div>
                <div style={styles.infoItem}>
                  <div style={styles.infoNumber}>2</div>
                  <div>
                    <strong style={styles.infoItemTitle}>Split & Process</strong>
                    <p style={styles.infoItemText}>Splits combined rows and processes odd-count sequences</p>
                  </div>
                </div>
                <div style={styles.infoItem}>
                  <div style={styles.infoNumber}>3</div>
                  <div>
                    <strong style={styles.infoItemTitle}>Reverse Numbers</strong>
                    <p style={styles.infoItemText}>Creates W2/S2 rows with reversed number sequences</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        ) : (
          /* Results Section */
          <div style={styles.resultsSection}>
            {/* Stats Cards */}
            <div style={styles.statsGrid}>
              {[
                { label: 'Input Rows', value: stats.input, icon: '📥', color: '#6366f1' },
                { label: 'Splits', value: stats.splits, icon: '✂️', color: '#f59e0b' },
                { label: 'Reversed', value: stats.reversed, icon: '🔄', color: '#10b981' },
                { label: 'Skipped', value: stats.skipped, icon: '⏭️', color: '#ef4444' },
                { label: 'Output Rows', value: stats.output, icon: '📤', color: '#8b5cf6' }
              ].map((stat) => (
                <div key={stat.label} style={styles.statCard}>
                  <div style={{ ...styles.statIcon, background: `${stat.color}20` }}>
                    <span style={{ fontSize: '1.5rem' }}>{stat.icon}</span>
                  </div>
                  <div style={{ ...styles.statValue, color: stat.color }}>{stat.value}</div>
                  <div style={styles.statLabel}>{stat.label}</div>
                </div>
              ))}
            </div>

            {/* Data Table */}
            <div style={styles.tableContainer}>
              <div style={styles.tableHeader}>
                <h3 style={styles.tableTitle}>
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <rect x="3" y="3" width="18" height="18" rx="2" ry="2" />
                    <line x1="3" y1="9" x2="21" y2="9" />
                    <line x1="9" y1="21" x2="9" y2="9" />
                  </svg>
                  Processed Data
                </h3>
                <span style={styles.rowCount}>{data.length} rows</span>
              </div>
              <div style={styles.tableWrapper}>
                <table style={styles.table}>
                  <thead>
                    <tr>
                      <th style={{ ...styles.th, ...styles.thFirst }}>#</th>
                      <th style={styles.th}>Status</th>
                      {Array.from({ length: maxCols }, (_, i) => (
                        <th key={i} style={styles.th}>{colLetter(i)}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {data.map((r, ri) => {
                      const rowStyles = getRowStyles(r);
                      return (
                        <tr key={ri} style={rowStyles.row}>
                          <td style={{ ...styles.td, ...styles.tdRowNum }}>{ri + 1}</td>
                          <td style={{ ...styles.td, ...styles.tdStatus }}>
                            <span style={{ ...styles.statusBadge, ...rowStyles.badge }}>
                              {r.isNew && '✨ '}
                              {r.type === 'original' ? 'Original' : r.type === 'filled' ? 'Reversed' : 'Unchanged'}
                            </span>
                          </td>
                          {r.data.map((c, ci) => (
                            <td
                              key={ci}
                              style={{
                                ...styles.td,
                                ...(typeof c === 'number' && (r.type === 'filled' || r.isNew) ? styles.tdHighlight : {}),
                                color: typeof c === 'number' && (r.type === 'filled' || r.isNew) ? '#10b981' : '#e2e8f0'
                              }}
                            >
                              {c ?? ''}
                            </td>
                          ))}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Action Buttons */}
            <div style={styles.actionButtons}>
              <button onClick={exportXL} style={styles.primaryButton}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <polyline points="7 10 12 15 17 10" />
                  <line x1="12" y1="15" x2="12" y2="3" />
                </svg>
                Download Excel
              </button>
              <button onClick={copy} style={styles.secondaryButton}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <rect x="9" y="9" width="13" height="13" rx="2" ry="2" />
                  <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
                </svg>
                Copy to Clipboard
              </button>
            </div>
          </div>
        )}

        {/* Footer */}
        <footer style={styles.footer}>
          <p>Splits at combined patterns • Reverses odd sequences • Skips even counts</p>
        </footer>
      </div>
    </div>
  );
}

function getRowStyles(r) {
  if (r.isNew) {
    return {
      row: { background: 'rgba(245, 158, 11, 0.08)' },
      badge: { background: 'rgba(245, 158, 11, 0.2)', color: '#fbbf24', border: '1px solid rgba(245, 158, 11, 0.3)' }
    };
  }
  switch (r.type) {
    case 'original':
      return {
        row: { background: 'rgba(99, 102, 241, 0.06)' },
        badge: { background: 'rgba(99, 102, 241, 0.2)', color: '#818cf8', border: '1px solid rgba(99, 102, 241, 0.3)' }
      };
    case 'filled':
      return {
        row: { background: 'rgba(16, 185, 129, 0.06)' },
        badge: { background: 'rgba(16, 185, 129, 0.2)', color: '#34d399', border: '1px solid rgba(16, 185, 129, 0.3)' }
      };
    default:
      return {
        row: { background: 'transparent' },
        badge: { background: 'rgba(100, 116, 139, 0.2)', color: '#94a3b8', border: '1px solid rgba(100, 116, 139, 0.3)' }
      };
  }
}

const styles = {
  container: {
    minHeight: '100vh',
    background: '#0f172a',
    position: 'relative',
    overflow: 'hidden',
  },
  bgOrb1: {
    position: 'absolute',
    top: '-20%',
    right: '-10%',
    width: '600px',
    height: '600px',
    background: 'radial-gradient(circle, rgba(99, 102, 241, 0.15) 0%, transparent 70%)',
    borderRadius: '50%',
    filter: 'blur(40px)',
    animation: 'float 20s ease-in-out infinite',
  },
  bgOrb2: {
    position: 'absolute',
    bottom: '-20%',
    left: '-10%',
    width: '500px',
    height: '500px',
    background: 'radial-gradient(circle, rgba(16, 185, 129, 0.12) 0%, transparent 70%)',
    borderRadius: '50%',
    filter: 'blur(40px)',
    animation: 'float 25s ease-in-out infinite reverse',
  },
  bgOrb3: {
    position: 'absolute',
    top: '50%',
    left: '50%',
    transform: 'translate(-50%, -50%)',
    width: '800px',
    height: '800px',
    background: 'radial-gradient(circle, rgba(139, 92, 246, 0.08) 0%, transparent 70%)',
    borderRadius: '50%',
    filter: 'blur(60px)',
  },
  wrapper: {
    position: 'relative',
    zIndex: 1,
    maxWidth: '1400px',
    margin: '0 auto',
    padding: '32px 24px',
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '40px',
    flexWrap: 'wrap',
    gap: '16px',
  },
  logoContainer: {
    display: 'flex',
    alignItems: 'center',
    gap: '16px',
  },
  logoIcon: {
    width: '56px',
    height: '56px',
    background: 'linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%)',
    borderRadius: '16px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: 'white',
    boxShadow: '0 8px 32px rgba(99, 102, 241, 0.3)',
  },
  title: {
    fontSize: '2rem',
    fontWeight: '700',
    color: '#f8fafc',
    margin: 0,
    letterSpacing: '-0.02em',
  },
  subtitle: {
    fontSize: '0.95rem',
    color: '#64748b',
    margin: 0,
  },
  resetButton: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 20px',
    background: 'rgba(255, 255, 255, 0.05)',
    border: '1px solid rgba(255, 255, 255, 0.1)',
    borderRadius: '10px',
    color: '#94a3b8',
    fontSize: '0.9rem',
    cursor: 'pointer',
    transition: 'all 0.2s',
  },
  inputSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: '24px',
  },
  tabContainer: {
    display: 'flex',
    gap: '8px',
    padding: '6px',
    background: 'rgba(255, 255, 255, 0.03)',
    borderRadius: '14px',
    width: 'fit-content',
  },
  tab: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '12px 24px',
    background: 'transparent',
    border: 'none',
    borderRadius: '10px',
    color: '#64748b',
    fontSize: '0.95rem',
    fontWeight: '500',
    cursor: 'pointer',
    transition: 'all 0.2s',
  },
  activeTab: {
    background: 'rgba(99, 102, 241, 0.15)',
    color: '#a5b4fc',
  },
  dropZone: {
    position: 'relative',
    border: '2px dashed rgba(99, 102, 241, 0.3)',
    borderRadius: '20px',
    padding: '60px 40px',
    textAlign: 'center',
    background: 'rgba(99, 102, 241, 0.03)',
    transition: 'all 0.3s',
    cursor: 'pointer',
  },
  dropZoneActive: {
    borderColor: '#6366f1',
    background: 'rgba(99, 102, 241, 0.1)',
    transform: 'scale(1.01)',
  },
  fileInput: {
    position: 'absolute',
    inset: 0,
    opacity: 0,
    cursor: 'pointer',
  },
  dropZoneContent: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '16px',
    cursor: 'pointer',
  },
  uploadIconContainer: {
    width: '80px',
    height: '80px',
    background: 'linear-gradient(135deg, rgba(99, 102, 241, 0.2) 0%, rgba(139, 92, 246, 0.2) 100%)',
    borderRadius: '20px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: '#a5b4fc',
  },
  dropText: {
    fontSize: '1.1rem',
    color: '#94a3b8',
    margin: 0,
  },
  dropTextHighlight: {
    color: '#a5b4fc',
    fontWeight: '600',
  },
  dropSubtext: {
    fontSize: '0.9rem',
    color: '#475569',
    margin: 0,
  },
  pasteContainer: {
    borderRadius: '20px',
    overflow: 'hidden',
  },
  textarea: {
    width: '100%',
    height: '200px',
    padding: '20px',
    background: 'rgba(15, 23, 42, 0.8)',
    border: '2px solid rgba(99, 102, 241, 0.2)',
    borderRadius: '20px',
    color: '#e2e8f0',
    fontSize: '0.95rem',
    fontFamily: "'JetBrains Mono', 'Fira Code', monospace",
    resize: 'vertical',
    outline: 'none',
    transition: 'border-color 0.2s',
  },
  infoCard: {
    background: 'rgba(255, 255, 255, 0.02)',
    border: '1px solid rgba(255, 255, 255, 0.06)',
    borderRadius: '20px',
    padding: '28px',
  },
  infoTitle: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    fontSize: '1.1rem',
    fontWeight: '600',
    color: '#e2e8f0',
    margin: '0 0 20px 0',
  },
  infoGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
    gap: '20px',
  },
  infoItem: {
    display: 'flex',
    gap: '14px',
    alignItems: 'flex-start',
  },
  infoNumber: {
    width: '32px',
    height: '32px',
    background: 'linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%)',
    borderRadius: '10px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: 'white',
    fontWeight: '700',
    fontSize: '0.9rem',
    flexShrink: 0,
  },
  infoItemTitle: {
    color: '#e2e8f0',
    fontSize: '0.95rem',
    display: 'block',
    marginBottom: '4px',
  },
  infoItemText: {
    color: '#64748b',
    fontSize: '0.85rem',
    margin: 0,
    lineHeight: 1.5,
  },
  resultsSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: '24px',
  },
  statsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))',
    gap: '16px',
  },
  statCard: {
    background: 'rgba(255, 255, 255, 0.03)',
    border: '1px solid rgba(255, 255, 255, 0.06)',
    borderRadius: '16px',
    padding: '20px',
    textAlign: 'center',
    transition: 'transform 0.2s, box-shadow 0.2s',
  },
  statIcon: {
    width: '48px',
    height: '48px',
    borderRadius: '12px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    margin: '0 auto 12px',
  },
  statValue: {
    fontSize: '1.75rem',
    fontWeight: '700',
    marginBottom: '4px',
  },
  statLabel: {
    fontSize: '0.8rem',
    color: '#64748b',
    textTransform: 'uppercase',
    letterSpacing: '0.05em',
  },
  tableContainer: {
    background: 'rgba(255, 255, 255, 0.02)',
    border: '1px solid rgba(255, 255, 255, 0.06)',
    borderRadius: '20px',
    overflow: 'hidden',
  },
  tableHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '20px 24px',
    borderBottom: '1px solid rgba(255, 255, 255, 0.06)',
  },
  tableTitle: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    fontSize: '1rem',
    fontWeight: '600',
    color: '#e2e8f0',
    margin: 0,
  },
  rowCount: {
    fontSize: '0.85rem',
    color: '#64748b',
    background: 'rgba(255, 255, 255, 0.05)',
    padding: '6px 12px',
    borderRadius: '20px',
  },
  tableWrapper: {
    overflowX: 'auto',
    maxHeight: '500px',
    overflowY: 'auto',
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    fontSize: '0.85rem',
  },
  th: {
    padding: '14px 12px',
    background: 'rgba(99, 102, 241, 0.08)',
    color: '#a5b4fc',
    fontWeight: '600',
    textAlign: 'center',
    position: 'sticky',
    top: 0,
    zIndex: 10,
    minWidth: '60px',
    borderBottom: '1px solid rgba(99, 102, 241, 0.2)',
  },
  thFirst: {
    minWidth: '50px',
  },
  td: {
    padding: '12px',
    textAlign: 'center',
    borderBottom: '1px solid rgba(255, 255, 255, 0.04)',
    fontFamily: "'JetBrains Mono', 'Fira Code', monospace",
    fontSize: '0.85rem',
  },
  tdRowNum: {
    color: '#475569',
    fontWeight: '500',
  },
  tdStatus: {
    padding: '8px 12px',
  },
  tdHighlight: {
    fontWeight: '600',
  },
  statusBadge: {
    display: 'inline-block',
    padding: '4px 10px',
    borderRadius: '6px',
    fontSize: '0.75rem',
    fontWeight: '500',
    fontFamily: 'system-ui, sans-serif',
  },
  actionButtons: {
    display: 'flex',
    gap: '16px',
    justifyContent: 'center',
    flexWrap: 'wrap',
  },
  primaryButton: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    padding: '16px 32px',
    background: 'linear-gradient(135deg, #10b981 0%, #059669 100%)',
    border: 'none',
    borderRadius: '14px',
    color: 'white',
    fontSize: '1rem',
    fontWeight: '600',
    cursor: 'pointer',
    boxShadow: '0 8px 32px rgba(16, 185, 129, 0.3)',
    transition: 'transform 0.2s, box-shadow 0.2s',
  },
  secondaryButton: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    padding: '16px 32px',
    background: 'rgba(99, 102, 241, 0.1)',
    border: '2px solid rgba(99, 102, 241, 0.3)',
    borderRadius: '14px',
    color: '#a5b4fc',
    fontSize: '1rem',
    fontWeight: '600',
    cursor: 'pointer',
    transition: 'all 0.2s',
  },
  footer: {
    textAlign: 'center',
    padding: '32px 0 16px',
    color: '#475569',
    fontSize: '0.85rem',
  },
};

// Add keyframes for animations
const styleSheet = document.createElement('style');
styleSheet.textContent = `
  @keyframes float {
    0%, 100% { transform: translate(0, 0); }
    50% { transform: translate(30px, -30px); }
  }
  
  button:hover {
    transform: translateY(-2px);
  }
  
  .stat-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 12px 40px rgba(0, 0, 0, 0.3);
  }
`;
document.head.appendChild(styleSheet);
