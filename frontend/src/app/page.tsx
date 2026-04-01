"use client";

import React, { useState } from 'react';

export default function FleetDashboard() {
  const [status, setStatus] = useState<string>("");
  const [matrixData, setMatrixData] = useState<any[] | null>(null);

  const handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setStatus(`Processing: ${file.name}...`);
    setMatrixData(null); // Clear the old table before loading new data

    const formData = new FormData();
    formData.append("file", file);

    try {
      // Sending to our internal Relay
      const response = await fetch("/api/upload", {
        method: "POST",
        body: formData,
      });
      
      const data = await response.json();
      
      if (response.ok && data.status === "success") {
        setStatus(`✅ ${data.message}`);
        setMatrixData(data.matrix); // Pushing the Python data into the UI Table
      } else {
        setStatus(`❌ Error: ${data.detail || 'Upload failed'}`);
      }
    } catch (err) {
      console.error(err);
      setStatus("❌ System Error. Check frontend terminal logs.");
    }
  };

  return (
    <div style={{ backgroundColor: '#020617', color: 'white', minHeight: '100vh', padding: '50px', fontFamily: 'sans-serif', textAlign: 'center' }}>
      <h1 style={{ color: '#3b82f6', fontSize: '3rem', fontWeight: '900', letterSpacing: '-2px', margin: '0' }}>FLEET R.H. BOT</h1>
      <p style={{ color: '#94a3b8', marginBottom: '40px', textTransform: 'uppercase', letterSpacing: '2px' }}>Extraction Engine v2.0</p>

      {/* Drag & Drop Zone */}
      <div style={{ border: '3px dashed #1e293b', padding: '40px', borderRadius: '30px', backgroundColor: '#0f172a', position: 'relative', maxWidth: '600px', margin: '0 auto', marginBottom: '40px' }}>
        <input type="file" onChange={handleFile} style={{ position: 'absolute', inset: 0, opacity: 0, cursor: 'pointer', width: '100%', height: '100%' }} />
        <p style={{ fontSize: '1.2rem', fontWeight: 'bold', margin: '0' }}>Drop MV ALEXIS Report Here</p>
      </div>

      {status && <p style={{ color: '#60a5fa', marginBottom: '30px', fontSize: '1.1rem' }}>{status}</p>}

      {/* THE DATA MATRIX GRID */}
      {matrixData && (
        <div style={{ maxWidth: '800px', margin: '0 auto', backgroundColor: '#0f172a', borderRadius: '15px', overflow: 'hidden', border: '1px solid #1e293b' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
            <thead>
              <tr style={{ backgroundColor: '#1e293b', color: '#94a3b8', textTransform: 'uppercase', fontSize: '0.9rem' }}>
                <th style={{ padding: '15px 20px' }}>Equipment</th>
                <th style={{ padding: '15px 20px' }}>Prev Hours</th>
                <th style={{ padding: '15px 20px' }}>Current Hours</th>
                <th style={{ padding: '15px 20px', color: '#3b82f6' }}>Monthly Run</th>
              </tr>
            </thead>
            <tbody>
              {matrixData.map((row, idx) => (
                <tr key={idx} style={{ borderBottom: '1px solid #1e293b' }}>
                  <td style={{ padding: '15px 20px', fontWeight: 'bold' }}>{row.equipment}</td>
                  <td style={{ padding: '15px 20px', color: '#94a3b8' }}>{row.previous.toLocaleString()}</td>
                  <td style={{ padding: '15px 20px' }}>{row.current.toLocaleString()}</td>
                  <td style={{ padding: '15px 20px', color: '#3b82f6', fontWeight: 'bold' }}>+{row.monthly}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}