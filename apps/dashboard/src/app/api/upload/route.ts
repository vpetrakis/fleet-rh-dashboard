import { NextRequest, NextResponse } from 'next/server'

const BACKEND = process.env.BACKEND_URL ?? 'http://127.0.0.1:8000'

export async function POST(request: NextRequest) {
  try {
    const incoming = await request.formData()
    const file = incoming.get('file')

    if (!file || typeof file === 'string') {
      return NextResponse.json(
        { status: 'error', detail: 'No file received. Please attach a .doc or .docx report.' },
        { status: 400 },
      )
    }

    // Forward the file to FastAPI using a fresh FormData object
    const fwd = new FormData()
    fwd.append('file', file)

    const res = await fetch(`${BACKEND}/api/upload-report`, {
      method: 'POST',
      body: fwd,
      signal: AbortSignal.timeout(60_000),   // Large files can take time
    })

    const data = await res.json()

    if (!res.ok) {
      // FastAPI returned a 4xx / 5xx — bubble the message through
      return NextResponse.json(
        { status: 'error', detail: data.detail ?? `Backend error ${res.status}` },
        { status: res.status },
      )
    }

    return NextResponse.json(data)
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    return NextResponse.json(
      {
        status: 'error',
        detail:
          'Cannot connect to the Python backend. ' +
          'Make sure it is running: cd services/brain && python3 main.py — ' +
          msg,
      },
      { status: 503 },
    )
  }
}