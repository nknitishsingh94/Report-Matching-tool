import React, { useState, useRef } from 'react';
import { UploadCloud, FileSpreadsheet, CheckCircle, AlertCircle, Download, RefreshCcw, FileText, Zap } from 'lucide-react';

function App() {
  const [masterFile, setMasterFile] = useState(null);
  const [smallFiles, setSmallFiles] = useState([]);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);

  const masterInputRef = useRef(null);
  const smallInputRef = useRef(null);

  const handleMasterUpload = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      setMasterFile(e.target.files[0]);
    }
  };

  const handleSmallUpload = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      // Append new files to the existing array instead of replacing
      setSmallFiles(prev => [...prev, ...Array.from(e.target.files)]);
    }
  };

  const handleMatch = async () => {
    if (!masterFile) {
      setError("Please upload the Master Report first.");
      return;
    }
    if (smallFiles.length === 0) {
      setError("Please upload at least one Small Report.");
      return;
    }

    setLoading(true);
    setError(null);

    const formData = new FormData();
    formData.append('masterFile', masterFile);
    smallFiles.forEach(file => {
      formData.append('smallFiles', file);
    });

    try {
      const response = await fetch('http://localhost:3001/api/match', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "Failed to match reports");
      }

      const data = await response.json();
      setResult(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const downloadExcel = () => {
    if (!result || !result.fileBase64) return;
    
    const byteCharacters = atob(result.fileBase64);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = result.fileName || 'Reconciled_Report.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const reset = () => {
    setMasterFile(null);
    setSmallFiles([]);
    setResult(null);
    setError(null);
    if (masterInputRef.current) masterInputRef.current.value = "";
    if (smallInputRef.current) smallInputRef.current.value = "";
  };

  return (
    <div className="min-h-screen bg-[conic-gradient(at_bottom_right,_var(--tw-gradient-stops))] from-slate-100 via-indigo-50 to-purple-50 py-12 px-4 sm:px-6 lg:px-8 font-sans selection:bg-indigo-200 selection:text-indigo-900">
      
      {/* Decorative background blobs */}
      <div className="absolute top-0 left-0 w-full h-full overflow-hidden pointer-events-none z-0">
        <div className="absolute -top-[10%] -left-[10%] w-[40%] h-[40%] rounded-full bg-purple-300/20 blur-3xl mix-blend-multiply"></div>
        <div className="absolute top-[20%] -right-[10%] w-[30%] h-[50%] rounded-full bg-indigo-300/20 blur-3xl mix-blend-multiply"></div>
      </div>

      <div className="max-w-4xl mx-auto space-y-10 relative z-10">
        
        {/* Header */}
        <div className="text-center space-y-4">
          <div className="inline-flex items-center justify-center p-3 bg-white rounded-2xl shadow-sm mb-2 ring-1 ring-gray-900/5">
            <Zap className="w-8 h-8 text-indigo-600 fill-indigo-100" />
          </div>
          <h1 className="text-4xl sm:text-5xl font-extrabold text-transparent bg-clip-text bg-gradient-to-r from-gray-900 via-indigo-800 to-gray-900 tracking-tight">
            Data Reconciliation Hub
          </h1>
          <p className="max-w-2xl mx-auto text-lg text-gray-600 font-medium">
            Seamlessly match and synchronize Caller IDs across multiple reports in seconds.
          </p>
        </div>

        <div className="bg-white/70 backdrop-blur-xl rounded-[2rem] shadow-2xl overflow-hidden border border-white/50 p-6 sm:p-10 transition-all duration-500">
          
          {error && (
            <div className="mb-8 p-4 bg-red-50/80 backdrop-blur-sm rounded-xl border border-red-200 flex items-start gap-3 shadow-sm animate-in slide-in-from-top-4">
              <AlertCircle className="w-6 h-6 text-red-600 shrink-0 mt-0.5" />
              <div>
                <h4 className="text-sm font-semibold text-red-800">Error</h4>
                <p className="text-sm text-red-700 mt-1">{error}</p>
              </div>
            </div>
          )}

          {!result ? (
            <div className="space-y-10">
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                {/* Master Upload */}
                <div className="space-y-4 relative group">
                  <div className="flex items-center justify-between">
                    <h3 className="text-lg font-bold text-gray-800 flex items-center gap-2">
                      <span className="flex items-center justify-center w-7 h-7 rounded-lg bg-indigo-100 text-indigo-700 text-sm font-bold">1</span>
                      Master Report
                    </h3>
                  </div>
                  
                  <div 
                    className={`relative cursor-pointer overflow-hidden border-2 border-dashed rounded-2xl p-8 flex flex-col items-center justify-center text-center transition-all duration-300 ${
                      masterFile 
                        ? 'border-indigo-400 bg-indigo-50/50 shadow-inner' 
                        : 'border-gray-200 bg-gray-50/50 hover:bg-white hover:border-indigo-300 hover:shadow-lg hover:-translate-y-1'
                    }`}
                    onClick={() => masterInputRef.current?.click()}
                  >
                    <input type="file" className="hidden" accept=".xlsx, .xls, .csv" ref={masterInputRef} onChange={handleMasterUpload} />
                    
                    {masterFile ? (
                      <div className="animate-in zoom-in-95 duration-300 flex flex-col items-center">
                        <div className="w-16 h-16 bg-white rounded-full shadow-sm flex items-center justify-center mb-4">
                          <CheckCircle className="w-8 h-8 text-indigo-500" />
                        </div>
                        <p className="text-sm font-bold text-gray-900 truncate w-full px-4">{masterFile.name}</p>
                        <p className="text-xs font-medium text-indigo-600 mt-1 bg-indigo-100/50 px-3 py-1 rounded-full">
                          {(masterFile.size / 1024).toFixed(1)} KB
                        </p>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center">
                        <div className="w-16 h-16 bg-white rounded-full shadow-sm flex items-center justify-center mb-4 group-hover:scale-110 transition-transform duration-300">
                          <UploadCloud className="w-8 h-8 text-gray-400 group-hover:text-indigo-500 transition-colors" />
                        </div>
                        <p className="text-sm font-semibold text-gray-700">Drop your Master file here</p>
                        <p className="text-xs text-gray-500 mt-1">.xlsx, .xls, .csv</p>
                      </div>
                    )}
                  </div>
                </div>

                {/* Small Reports Upload */}
                <div className="space-y-4 relative group">
                  <div className="flex items-center justify-between">
                    <h3 className="text-lg font-bold text-gray-800 flex items-center gap-2">
                      <span className="flex items-center justify-center w-7 h-7 rounded-lg bg-purple-100 text-purple-700 text-sm font-bold">2</span>
                      Small Reports
                    </h3>
                  </div>
                  
                  <div 
                    className={`relative cursor-pointer overflow-hidden border-2 border-dashed rounded-2xl p-8 flex flex-col items-center justify-center text-center transition-all duration-300 ${
                      smallFiles.length > 0 
                        ? 'border-purple-400 bg-purple-50/50 shadow-inner' 
                        : 'border-gray-200 bg-gray-50/50 hover:bg-white hover:border-purple-300 hover:shadow-lg hover:-translate-y-1'
                    }`}
                    onClick={() => smallInputRef.current?.click()}
                  >
                    <input type="file" className="hidden" accept=".xlsx, .xls, .csv" multiple ref={smallInputRef} onChange={handleSmallUpload} />
                    
                    {smallFiles.length > 0 ? (
                      <div className="animate-in zoom-in-95 duration-300 flex flex-col items-center w-full">
                        <div className="w-16 h-16 bg-white rounded-full shadow-sm flex items-center justify-center mb-4">
                          <FileText className="w-8 h-8 text-purple-500" />
                        </div>
                        <p className="text-sm font-bold text-gray-900">{smallFiles.length} files selected</p>
                        <p className="text-xs font-medium text-purple-600 mt-1 bg-purple-100/50 px-3 py-1 rounded-full truncate max-w-full">
                          {smallFiles.map(f => f.name).join(', ')}
                        </p>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center">
                        <div className="w-16 h-16 bg-white rounded-full shadow-sm flex items-center justify-center mb-4 group-hover:scale-110 transition-transform duration-300">
                          <FileSpreadsheet className="w-8 h-8 text-gray-400 group-hover:text-purple-500 transition-colors" />
                        </div>
                        <p className="text-sm font-semibold text-gray-700">Drop multiple reports here</p>
                        <p className="text-xs text-gray-500 mt-1">.xlsx, .xls, .csv</p>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* Action Button */}
              <div className="pt-4 border-t border-gray-100">
                <button
                  onClick={handleMatch}
                  disabled={loading || !masterFile || smallFiles.length === 0}
                  className={`relative overflow-hidden w-full py-5 rounded-2xl font-bold text-lg text-white shadow-xl transition-all duration-300 flex items-center justify-center gap-3 ${
                    loading || !masterFile || smallFiles.length === 0
                      ? 'bg-gray-300 cursor-not-allowed shadow-none' 
                      : 'bg-gradient-to-r from-indigo-600 via-purple-600 to-indigo-600 bg-[length:200%_100%] hover:bg-[100%_0] hover:shadow-indigo-500/25 hover:-translate-y-0.5'
                  }`}
                >
                  {loading ? (
                    <>
                      <RefreshCcw className="w-6 h-6 animate-spin" />
                      Reconciling Data...
                    </>
                  ) : (
                    <>
                      Start Reconciliation
                    </>
                  )}
                </button>
              </div>
            </div>
          ) : (
            /* Results View */
            <div className="space-y-10 animate-in fade-in slide-in-from-bottom-8 duration-500">
              <div className="text-center space-y-4">
                <div className="inline-flex items-center justify-center w-20 h-20 rounded-full bg-gradient-to-tr from-green-400 to-emerald-500 shadow-lg shadow-green-500/30 mb-2">
                  <CheckCircle className="w-10 h-10 text-white" />
                </div>
                <h2 className="text-3xl font-extrabold text-gray-900 tracking-tight">Reconciliation Complete</h2>
                <p className="text-gray-500 font-medium text-lg">Your master report has been successfully updated with the latest data.</p>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-6">
                <div className="bg-white rounded-2xl p-6 text-center shadow-sm border border-gray-100 hover:shadow-md transition-shadow">
                  <p className="text-sm text-gray-500 font-bold uppercase tracking-wider mb-2">Total Processed</p>
                  <p className="text-4xl font-black text-gray-900">{result.stats.totalMaster}</p>
                </div>
                
                <div className="bg-gradient-to-br from-emerald-50 to-green-50 rounded-2xl p-6 text-center border border-green-100 shadow-sm shadow-green-100/50 hover:shadow-md transition-shadow relative overflow-hidden">
                  <div className="absolute top-0 right-0 p-4 opacity-10">
                    <CheckCircle className="w-16 h-16 text-green-600" />
                  </div>
                  <p className="text-sm text-green-700 font-bold uppercase tracking-wider mb-2 relative z-10">Matched</p>
                  <p className="text-4xl font-black text-green-700 relative z-10">{result.stats.matched}</p>
                </div>
                
                <div className="bg-gradient-to-br from-rose-50 to-red-50 rounded-2xl p-6 text-center border border-red-100 shadow-sm shadow-red-100/50 hover:shadow-md transition-shadow relative overflow-hidden">
                  <div className="absolute top-0 right-0 p-4 opacity-10">
                    <AlertCircle className="w-16 h-16 text-red-600" />
                  </div>
                  <p className="text-sm text-red-700 font-bold uppercase tracking-wider mb-2 relative z-10">Unmatched</p>
                  <p className="text-4xl font-black text-red-700 relative z-10">{result.stats.notFound}</p>
                </div>
              </div>

              <div className="flex flex-col sm:flex-row gap-4 pt-6 border-t border-gray-100">
                <button
                  onClick={reset}
                  className="flex-1 py-4 px-6 bg-white border-2 border-gray-200 text-gray-700 rounded-2xl font-bold hover:bg-gray-50 hover:border-gray-300 transition-all focus:ring-4 focus:ring-gray-100"
                >
                  Start Over
                </button>
                <button
                  onClick={downloadExcel}
                  className="flex-[2] py-4 px-6 bg-gradient-to-r from-gray-900 to-gray-800 text-white rounded-2xl font-bold shadow-xl shadow-gray-900/20 hover:shadow-gray-900/30 hover:-translate-y-0.5 transition-all flex items-center justify-center gap-3"
                >
                  <Download className="w-6 h-6" />
                  Download Master Report
                </button>
              </div>
            </div>
          )}

        </div>
      </div>
    </div>
  );
}

export default App;
