import { NextResponse } from 'next/server';

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    
    // Forwarding the payload directly to the local Python engine
    const response = await fetch('http://127.0.0.1:8000/api/upload-report', {
      method: 'POST',
      body: formData,
    });
    
    const data = await response.json();
    return NextResponse.json(data, { status: response.status });
    
  } catch (error) {
    console.error('Relay Error:', error);
    return NextResponse.json(
      { status: 'error', detail: 'Next.js could not reach Python on 127.0.0.1:8000' }, 
      { status: 500 }
    );
  }
}