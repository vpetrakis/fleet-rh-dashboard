import { NextResponse } from 'next/server';

export async function GET() {
  try {
    // Reach out to the local Python FastAPI server
    const res = await fetch('http://127.0.0.1:8000/api/init', { 
      cache: 'no-store' // Never cache this, always get the freshest data
    });
    
    if (!res.ok) {
      return NextResponse.json({ status: 'empty' });
    }
    
    const data = await res.json();
    return NextResponse.json(data);
    
  } catch (error) {
    return NextResponse.json({ 
      status: 'error', 
      detail: 'Python backend is currently offline.' 
    });
  }
}
