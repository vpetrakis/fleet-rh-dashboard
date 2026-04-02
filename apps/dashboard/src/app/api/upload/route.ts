import { NextResponse } from 'next/server';
export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const response = await fetch('http://127.0.0.1:8000/api/upload-report', { method: 'POST', body: formData });
    return NextResponse.json(await response.json());
  } catch (error) {
    return NextResponse.json({ status: 'error', detail: 'Brain Offline.' }, { status: 500 });
  }
}
