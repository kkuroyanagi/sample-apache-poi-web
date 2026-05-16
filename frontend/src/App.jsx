import { useState } from 'react'
import './App.css'

// ── 共通ダウンロードフック ────────────────────────────────────────────────────

function useDownload(path, filename) {
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)

  const download = async () => {
    setLoading(true)
    setError(null)
    try {
      const res = await fetch(path)
      if (!res.ok) throw new Error(`サーバーエラー: ${res.status}`)
      const blob = await res.blob()
      triggerDownload(blob, filename)
    } catch (e) {
      setError(e.message)
    } finally {
      setLoading(false)
    }
  }

  return { loading, error, download }
}

// ── パフォーマンステスト用フック ──────────────────────────────────────────────

function usePerfDownload() {
  const [loadingMode, setLoadingMode] = useState(null) // 'xssf' | 'sxssf' | null
  const [results, setResults]         = useState([])
  const [error, setError]             = useState(null)

  const runTest = async (mode, rows) => {
    setLoadingMode(mode)
    setError(null)
    const t0 = performance.now()
    try {
      const res = await fetch(`/api/excel/download-${mode}?rows=${rows}`)
      if (!res.ok) {
        const body = await res.json().catch(() => ({}))
        throw new Error(body.error || `HTTP ${res.status}`)
      }
      const blob    = await res.blob()
      const totalMs = Math.round(performance.now() - t0)

      setResults(prev => [{
        mode:    mode.toUpperCase(),
        rows:    Number(res.headers.get('X-Row-Count')).toLocaleString(),
        genMs:   res.headers.get('X-Generation-Time-Ms') ?? '-',
        totalMs,
        heapMb:  res.headers.get('X-Heap-Delta-MB') ?? '-',
        fileMb:  (blob.size / 1024 / 1024).toFixed(2),
      }, ...prev].slice(0, 10))

      triggerDownload(blob, `perf_${mode}_${rows}.xlsx`)
    } catch (e) {
      setError(e.message)
    } finally {
      setLoadingMode(null)
    }
  }

  return { loadingMode, results, error, runTest }
}

// ── ユーティリティ ────────────────────────────────────────────────────────────

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url; a.download = filename; a.click()
  URL.revokeObjectURL(url)
}

// ── コンポーネント ────────────────────────────────────────────────────────────

function PerfCard() {
  const [rows, setRows] = useState(1000)
  const { loadingMode, results, error, runTest } = usePerfDownload()
  const PRESETS = [100, 500, 1_000, 5_000, 10_000, 50_000, 100_000]

  return (
    <main className="card">
      <h2>パフォーマンステスト — XSSF vs SXSSF</h2>
      <p className="note">
        列数 300（入力列 10 本＋データ列 290 本）の Excel を生成して比較します。<br />
        入力列（黄色・先頭 10 列）はマスタシートのプルダウンで選択可能。<br />
        XSSF は全行をヒープに保持（上限 5,000 行）。
        SXSSF はウィンドウ超過行をディスクに退避し省メモリで最大 100,000 行を処理します。
      </p>

      <div className="perf-controls">
        <div className="perf-row">
          <label>行数</label>
          <input
            type="number" value={rows} min={1} max={1000000}
            onChange={e => setRows(Math.max(1, Math.min(100000, Number(e.target.value))))}
          />
        </div>
        <div className="presets">
          {PRESETS.map(n => (
            <button key={n} className="preset" onClick={() => setRows(n)}>
              {n.toLocaleString()}
            </button>
          ))}
        </div>
      </div>

      <div className="perf-buttons">
        <button onClick={() => runTest('xssf', rows)} disabled={!!loadingMode}>
          {loadingMode === 'xssf' ? '生成中...' : 'XSSF でテスト'}
        </button>
        <button className="btn-sxssf" onClick={() => runTest('sxssf', rows)} disabled={!!loadingMode}>
          {loadingMode === 'sxssf' ? '生成中...' : 'SXSSF でテスト'}
        </button>
      </div>

      {error && <p className="error">{error}</p>}

      {results.length > 0 && (
        <div className="perf-results">
          <table>
            <thead>
              <tr>
                <th>方式</th><th>行数</th>
                <th>生成時間</th><th>転送込み合計</th>
                <th>ヒープ増減</th><th>ファイルサイズ</th>
              </tr>
            </thead>
            <tbody>
              {results.map((r, i) => (
                <tr key={i}>
                  <td className={r.mode === 'XSSF' ? 'tag-xssf' : 'tag-sxssf'}>{r.mode}</td>
                  <td>{r.rows}</td>
                  <td>{r.genMs} ms</td>
                  <td>{r.totalMs} ms</td>
                  <td>{r.heapMb} MB</td>
                  <td>{r.fileMb} MB</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </main>
  )
}

// ── メイン ────────────────────────────────────────────────────────────────────

export default function App() {
  const single    = useDownload('/api/excel/download',          'sales_report.xlsx')
  const multi     = useDownload('/api/excel/download-multi',    'sales_5sheets.xlsx')
  const editable  = useDownload('/api/excel/download-editable', 'sales_editable.xlsx')
  const protectedSheet = useDownload('/api/excel/download-protected', 'protected_category_sales.xlsx')

  return (
    <div className="container">
      <header>
        <h1>Excel ダウンロード</h1>
        <p>Apache POI で生成した売上レポートをダウンロードします。</p>
      </header>

      <main className="card">
        <h2>売上レポート（2025年4月） — 1シート</h2>
        <table>
          <tbody>
            <tr><th>期間</th><td>2025/04/01 〜 2025/04/30</td></tr>
            <tr><th>件数</th><td>10件</td></tr>
            <tr><th>内容</th><td>商品別売上（日付・商品名・カテゴリ・数量・単価・合計）</td></tr>
            <tr><th>形式</th><td>Excel (.xlsx)</td></tr>
          </tbody>
        </table>
        <button onClick={single.download} disabled={single.loading}>
          {single.loading ? 'ダウンロード中...' : 'Excel をダウンロード'}
        </button>
        {single.error && <p className="error">{single.error}</p>}
      </main>

      <main className="card">
        <h2>売上レポート（2025年4〜5月） — 5シート</h2>
        <table>
          <tbody>
            <tr><th>シート 1</th><td>4月 売上データ（明細）</td></tr>
            <tr><th>シート 2</th><td>5月 売上データ（明細）</td></tr>
            <tr><th>シート 3</th><td>カテゴリ別集計（4〜5月）</td></tr>
            <tr><th>シート 4</th><td>月次比較（増減・増減率）</td></tr>
            <tr><th>シート 5</th><td>サマリー（KPI・最多売上など）</td></tr>
          </tbody>
        </table>
        <button onClick={multi.download} disabled={multi.loading}>
          {multi.loading ? 'ダウンロード中...' : '5シート Excel をダウンロード'}
        </button>
        {multi.error && <p className="error">{multi.error}</p>}
      </main>

      <main className="card">
        <h2>売上入力シート — マスタ選択方式</h2>
        <table>
          <tbody>
            <tr><th>シート 1</th><td>売上入力（30行・黄色セルに入力）</td></tr>
            <tr><th>シート 2</th><td>商品マスタ（16商品・プルダウン参照元）</td></tr>
            <tr>
              <th>操作方法</th>
              <td>
                商品名セルのプルダウンから選択 →
                カテゴリ・単価が自動設定 →
                数量を入力 → 合計が自動計算
              </td>
            </tr>
          </tbody>
        </table>
        <button onClick={editable.download} disabled={editable.loading}>
          {editable.loading ? 'ダウンロード中...' : '入力シート Excel をダウンロード'}
        </button>
        {editable.error && <p className="error">{editable.error}</p>}
      </main>

      <main className="card">
        <h2>保護されたカテゴリ別売上シート — マスタ保護方式</h2>
        <table>
          <tbody>
            <tr><th>シート 1</th><td>カテゴリ別売上入力（20行・黄色セルに入力）</td></tr>
            <tr><th>シート 2</th><td>カテゴリマスタ（5カテゴリ・保護済み）</td></tr>
            <tr><th>シート 3</th><td>商品マスタ（20商品・保護済み）</td></tr>
            <tr>
              <th>操作方法</th>
              <td>
                カテゴリを選択 → 
                商品名をプルダウンから選択 →
                単価が自動設定 →
                数量を入力 → 合計が自動計算
              </td>
            </tr>
            <tr><th>特徴</th><td>マスタシートが保護されているため、誤編集を防止できます</td></tr>
          </tbody>
        </table>
        <button onClick={protectedSheet.download} disabled={protectedSheet.loading}>
          {protectedSheet.loading ? 'ダウンロード中...' : '保護シート Excel をダウンロード'}
        </button>
        {protectedSheet.error && <p className="error">{protectedSheet.error}</p>}
      </main>

      <PerfCard />
    </div>
  )
}
