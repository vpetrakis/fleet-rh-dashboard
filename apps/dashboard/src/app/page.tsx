"use client";
import React, { useState } from 'react';
import { UploadCloud, CheckCircle2, AlertCircle, Settings, Ship } from 'lucide-react';

export default function FleetCommand() {
  const [status, setStatus] = useState<"idle" | "loading" | "success" | "error">("idle");
  const [message, setMessage] = useState("");
  const [data, setData] = useState<any>(null);
  
  const [activeTab, setActiveTab] = useState<"ME" | "DG_GEN" | "DG_CYL">("ME");
  const [activeDg, setActiveDg] = useState<number>(1);

  const handleUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setStatus("loading");
    setMessage("Detecting Vessel and extracting dynamic matrix...");
    const formData = new FormData();
    formData.append("file", file);

    try {
      const res = await fetch("/api/upload", { method: "POST", body: formData });
      const result = await res.json();
      
      if (result.status === "success") {
        setData(result);
        setStatus("success");
        setMessage(`${result.vessel} Profile Synchronized.`);
      } else {
        setStatus("error");
        setMessage(result.detail || "Extraction failed.");
      }
    } catch {
      setStatus("error");
      setMessage("Connection to Server lost.");
    }
  };

  const renderHealthCell = (hoursStr: string, limit: number, key: string) => {
    if (!hoursStr || hoursStr === "-") return <td key={key} className="p-5 text-center text-slate-600">-</td>;
    
    const hours = parseInt(hoursStr);
    const percent = Math.min((hours / limit) * 100, 100);
    
    let colorClass = "bg-emerald-500"; let textClass = "text-emerald-400";
    if (hours > limit) { colorClass = "bg-red-500"; textClass = "text-red-400 font-bold animate-pulse"; } 
    else if (percent > 85) { colorClass = "bg-amber-500"; textClass = "text-amber-400"; }

    return (
      <td key={key} className="p-5 text-center">
        <div className={`font-mono text-lg ${textClass}`}>{hours.toLocaleString()}</div>
        <div className="w-full bg-slate-800 h-1 mt-2 rounded-full overflow-hidden flex">
          <div className={`${colorClass} h-full`} style={{ width: `${percent}%` }}></div>
        </div>
      </td>
    );
  };

  return (
    <main className="min-h-screen flex flex-col items-center py-12 px-6 font-sans bg-slate-950">
      <div className="text-center mb-8">
        <Settings size={32} className="mx-auto text-blue-500 mb-3 animate-[spin_10s_linear_infinite]" />
        <h1 className="text-3xl font-extrabold text-white tracking-tight">FLEET ANALYTICS ENGINE</h1>
      </div>

      {!data && (
        <div className="relative w-full max-w-2xl group cursor-pointer border border-slate-700 bg-slate-900/60 p-10 rounded-2xl text-center hover:border-blue-500 transition-all">
          <input type="file" onChange={handleUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" accept=".doc,.docx" />
          <UploadCloud className="mx-auto h-12 w-12 text-blue-400 mb-3" />
          <h3 className="text-lg font-semibold text-slate-200">Engage Smart Extraction</h3>
        </div>
      )}

      <div className="h-6 mt-2 flex items-center justify-center space-x-2 text-sm">
        {status === 'success' && <CheckCircle2 size={16} className="text-emerald-400" />}
        {status === 'error' && <AlertCircle size={16} className="text-red-400" />}
        <span className={status === 'error' ? 'text-red-400' : 'text-blue-400'}>{message}</span>
      </div>

      {data && status === 'success' && (
        <div className="w-full max-w-7xl mt-4">
          
          <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 mb-6 flex items-center space-x-4">
            <Ship className="text-blue-500" size={24} />
            <div>
              <h2 className="text-white font-bold text-lg">{data.vessel}</h2>
              <p className="text-slate-400 text-xs uppercase tracking-wider">
                M/E Cylinders: {data.config.me_cylinders} | D/G Count: {data.config.dg_count} | Aux Cylinders: {data.config.aux_cylinders}
              </p>
            </div>
          </div>

          <div className="flex space-x-8 mb-6 border-b border-slate-800">
            <button onClick={() => setActiveTab("ME")} className={`pb-3 font-bold uppercase tracking-wider text-sm transition-colors ${activeTab === 'ME' ? 'text-blue-400 border-b-2 border-blue-400' : 'text-slate-500'}`}>Main Engine</button>
            <button onClick={() => setActiveTab("DG_GEN")} className={`pb-3 font-bold uppercase tracking-wider text-sm transition-colors ${activeTab === 'DG_GEN' ? 'text-blue-400 border-b-2 border-blue-400' : 'text-slate-500'}`}>D/G Overview</button>
            <button onClick={() => setActiveTab("DG_CYL")} className={`pb-3 font-bold uppercase tracking-wider text-sm transition-colors ${activeTab === 'DG_CYL' ? 'text-blue-400 border-b-2 border-blue-400' : 'text-slate-500'}`}>D/G Cylinders</button>
          </div>

          <div className="border border-slate-800 rounded-xl overflow-x-auto bg-slate-900/40 shadow-2xl">
            
            {/* DYNAMIC MAIN ENGINE TAB */}
            {activeTab === "ME" && (
              <table className="w-full text-left text-sm whitespace-nowrap">
                <thead className="bg-slate-900/90 text-slate-400 text-xs uppercase tracking-wider">
                  <tr>
                    <th className="p-5 font-semibold border-b border-slate-700">Equipment / Limit</th>
                    {Array.from({ length: data.config.me_cylinders }).map((_, i) => (
                      <th key={i} className="p-5 font-semibold border-b border-slate-700 text-center">CYL {i + 1}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/50">
                  {data.matrices.main_engine.map((row: any, i: number) => (
                    <tr key={i} className="hover:bg-slate-800/40">
                      <td className="p-5"><div className="font-bold text-blue-400">{row.component}</div><div className="text-xs text-slate-500">Limit: {row.limit.toLocaleString()}</div></td>
                      {Array.from({ length: data.config.me_cylinders }).map((_, j) => renderHealthCell(row[`cyl${j + 1}`], row.limit, `cyl-${j}`))}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            {/* DYNAMIC D/G OVERVIEW TAB */}
            {activeTab === "DG_GEN" && (
              <table className="w-full text-left text-sm whitespace-nowrap">
                <thead className="bg-slate-900/90 text-slate-400 text-xs uppercase tracking-wider">
                  <tr>
                    <th className="p-5 font-semibold border-b border-slate-700">Equipment / Limit</th>
                    {Array.from({ length: data.config.dg_count }).map((_, i) => (
                      <th key={i} className="p-5 font-semibold border-b border-slate-700 text-center">D/G {i + 1}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/50">
                  {data.matrices.dg_general.map((row: any, i: number) => (
                    <tr key={i} className="hover:bg-slate-800/40">
                      <td className="p-5"><div className="font-bold text-blue-400">{row.component}</div><div className="text-xs text-slate-500">Limit: {row.limit.toLocaleString()}</div></td>
                      {Array.from({ length: data.config.dg_count }).map((_, j) => renderHealthCell(row[`dg${j + 1}`], row.limit, `dg-${j}`))}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            {/* DYNAMIC D/G CYLINDERS TAB */}
            {activeTab === "DG_CYL" && (
              <div>
                <div className="bg-slate-800/50 p-4 border-b border-slate-700 flex justify-center space-x-4">
                  {Array.from({ length: data.config.dg_count }).map((_, i) => (
                    <button key={i} onClick={() => setActiveDg(i + 1)} className={`px-6 py-2 rounded-full text-xs font-bold uppercase tracking-wider ${activeDg === i + 1 ? 'bg-blue-500 text-white' : 'bg-slate-700 text-slate-400'}`}>Generator {i + 1}</button>
                  ))}
                </div>
                <table className="w-full text-left text-sm whitespace-nowrap">
                  <thead className="bg-slate-900/90 text-slate-400 text-xs uppercase tracking-wider">
                    <tr>
                      <th className="p-5 font-semibold border-b border-slate-700">Equipment / Limit</th>
                      {Array.from({ length: data.config.aux_cylinders }).map((_, i) => (
                         <th key={i} className="p-5 font-semibold border-b border-slate-700 text-center">CYL {i + 1}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50">
                    {data.matrices.aux_engine.map((row: any, i: number) => (
                      <tr key={i} className="hover:bg-slate-800/40">
                        <td className="p-5"><div className="font-bold text-blue-400">{row.component}</div><div className="text-xs text-slate-500">Limit: {row.limit.toLocaleString()}</div></td>
                        {Array.from({ length: data.config.aux_cylinders }).map((_, j) => renderHealthCell(row[`dg${activeDg}`][j], row.limit, `aux-${j}`))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      )}
    </main>
  );
}