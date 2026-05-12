import { useState } from 'react'
import './App.css'

export default function App() {
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)

  const handleDownload = async () => {
    setLoading(true)
    setError(null)
    try {
      const res = await fetch('/api/excel/download')
      if (!res.ok) throw new Error(`サーバーエラー: ${res.status}`)
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'sales_report.xlsx'
      a.click()
      URL.revokeObjectURL(url)
    } catch (e) {
      setError(e.message)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="container">
      <header>
        <h1>Excel ダウンロード</h1>
        <p>Apache POI で生成した売上レポートをダウンロードします。</p>
      </header>

      <main className="card">
        <h2>売上レポート（2025年4月）</h2>
        <table>
          <tbody>
            <tr><th>期間</th><td>2025/04/01 〜 2025/04/30</td></tr>
            <tr><th>件数</th><td>10件</td></tr>
            <tr><th>内容</th><td>商品別売上（日付・商品名・カテゴリ・数量・単価・合計）</td></tr>
            <tr><th>形式</th><td>Excel (.xlsx)</td></tr>
          </tbody>
        </table>

        <button onClick={handleDownload} disabled={loading}>
          {loading ? 'ダウンロード中...' : 'Excel をダウンロード'}
        </button>

        {error && <p className="error">{error}</p>}
      </main>
    </div>
  )
}
