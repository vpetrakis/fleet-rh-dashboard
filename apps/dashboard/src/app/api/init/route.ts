import { NextResponse } from 'next/server'

// The Python backend URL — works for both local dev and GitHub Codespaces
// (both run on the same machine; Next.js API routes talk server-to-server)
const BACKEND = process.env.BACKEND_URL ?? 'http://127.0.0.1:8000'

export async function GET() {
  try {
    const res = await fetch(`${BACKEND}/api/init`, {
      cache: 'no-store',   // Always fresh — never serve a cached response
      signal: AbortSignal.timeout(8_000),
    })
    const data = await res.json()
    return NextResponse.json(data)
  } catch {
    // Backend not running yet — return idle so the UI shows the upload prompt
    return NextResponse.json({ status: 'idle' })
  }
}