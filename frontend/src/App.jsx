import React, { useState, useRef } from 'react';
import { UploadCloud, FileSpreadsheet, CheckCircle, AlertCircle, Download, RefreshCcw, FileText, Zap, Eye, X, PhoneCall, Activity, Database } from 'lucide-react';

function App() {
  const [masterFile, setMasterFile] = useState(null);
  const [smallFiles, setSmallFiles] = useState([]);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const [previewFile, setPreviewFile] = useState(null);
  const [filterStatus, setFilterStatus] = useState('');
  const [filterDate, setFilterDate] = useState('');
  const [filterExt, setFilterExt] = useState('');
  const [filterTz, setFilterTz] = useState('');
  const [filterTime, setFilterTime] = useState('');
  
  // Validation State
  const [mode, setMode] = useState('reconcile'); // 'reconcile' | 'validate'
  const [validationFile, setValidationFile] = useState(null);
  const [validationResult, setValidationResult] = useState(null);
  
  const validationInputRef = useRef(null);

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
      const newFiles = Array.from(e.target.files).map(file => ({
        file,
        label: file.name.split('.')[0]
      }));
      setSmallFiles(prev => [...prev, ...newFiles]);
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
    smallFiles.forEach((item) => {
      formData.append('smallFiles', item.file);
      formData.append('labels', item.label);
    });

    try {
      const apiUrl = import.meta.env.VITE_API_URL || '';
      const response = await fetch(`${apiUrl}/api/match`, {
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
      setLoading(false);
    }
  };

  const handleValidationUpload = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      setValidationFile(e.target.files[0]);
    }
  };

  const handleValidation = async () => {
    if (!validationFile) {
      setError("Please upload a sheet to validate.");
      return;
    }

    setLoading(true);
    setError(null);

    const formData = new FormData();
    formData.append('sheet', validationFile);

    try {
      const apiUrl = import.meta.env.VITE_API_URL || '';
      const response = await fetch(`${apiUrl}/api/validate-numbers`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "Failed to validate numbers");
      }

      const data = await response.json();
      setValidationResult(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const downloadExcel = (fileObj) => {
    if (!fileObj || !fileObj.fileBase64) return;

    const byteCharacters = atob(fileObj.fileBase64);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const isZip = fileObj.fileName && fileObj.fileName.endsWith('.zip');
    const mimeType = isZip 
      ? 'application/zip' 
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    
    const blob = new Blob([byteArray], { type: mimeType });

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileObj.fileName || 'Reconciled_Report.xlsx';
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
    setPreviewFile(null);
    setValidationFile(null);
    setValidationResult(null);
    if (masterInputRef.current) masterInputRef.current.value = "";
    if (smallInputRef.current) smallInputRef.current.value = "";
    if (validationInputRef.current) validationInputRef.current.value = "";
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
            {mode === 'reconcile' ? <Zap className="w-8 h-8 text-indigo-600 fill-indigo-100" /> : <PhoneCall className="w-8 h-8 text-indigo-600 fill-indigo-100" />}
          </div>
          <h1 className="text-4xl sm:text-5xl font-extrabold text-transparent bg-clip-text bg-gradient-to-r from-gray-900 via-indigo-800 to-gray-900 tracking-tight">
            {mode === 'reconcile' ? 'Data Reconciliation Hub' : 'Number Validation Hub'}
          </h1>
          <p className="max-w-2xl mx-auto text-lg text-gray-600 font-medium">
            {mode === 'reconcile' 
              ? 'Seamlessly match and synchronize Caller IDs across multiple reports in seconds.'
              : 'Upload a sheet of phone numbers to detect VoIP, mobile, and landline types instantly.'}
          </p>
        </div>

        {/* Mode Toggle */}
        <div className="flex justify-center">
          <div className="inline-flex bg-white/60 backdrop-blur-md p-1.5 rounded-2xl shadow-sm border border-white/50">
            <button
              onClick={() => { setMode('reconcile'); reset(); }}
              className={`flex items-center gap-2 px-6 py-2.5 rounded-xl font-bold text-sm transition-all duration-300 ${mode === 'reconcile' ? 'bg-white text-indigo-700 shadow-md ring-1 ring-black/5' : 'text-gray-500 hover:text-gray-700 hover:bg-white/40'}`}
            >
              <Database className="w-4 h-4" />
              Reconciliation
            </button>
            <button
              onClick={() => { setMode('validate'); reset(); }}
              className={`flex items-center gap-2 px-6 py-2.5 rounded-xl font-bold text-sm transition-all duration-300 ${mode === 'validate' ? 'bg-white text-indigo-700 shadow-md ring-1 ring-black/5' : 'text-gray-500 hover:text-gray-700 hover:bg-white/40'}`}
            >
              <PhoneCall className="w-4 h-4" />
              Number Validation
            </button>
          </div>
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

          {mode === 'reconcile' ? (
            !result ? (
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
                    className={`relative cursor-pointer overflow-hidden border-2 border-dashed rounded-2xl p-8 flex flex-col items-center justify-center text-center transition-all duration-300 ${masterFile
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
                    className={`relative overflow-hidden border-2 border-dashed rounded-2xl p-6 flex flex-col items-center justify-center text-center transition-all duration-300 ${smallFiles.length > 0
                        ? 'border-purple-400 bg-purple-50/50 shadow-inner'
                        : 'border-gray-200 bg-gray-50/50 hover:bg-white hover:border-purple-300 hover:shadow-lg hover:-translate-y-1 cursor-pointer'
                      }`}
                    onClick={(e) => {
                      if (smallFiles.length === 0) smallInputRef.current?.click();
                    }}
                  >
                    <input type="file" className="hidden" accept=".xlsx, .xls, .csv" multiple ref={smallInputRef} onChange={handleSmallUpload} />

                    {smallFiles.length > 0 ? (
                      <div className="animate-in zoom-in-95 duration-300 flex flex-col w-full text-left space-y-3" onClick={(e) => e.stopPropagation()}>
                        <div className="flex items-center gap-2 mb-2 justify-center">
                          <FileText className="w-6 h-6 text-purple-500" />
                          <p className="text-sm font-bold text-gray-900">{smallFiles.length} files selected</p>
                        </div>
                        <div className="max-h-48 overflow-y-auto space-y-2 pr-2">
                          {smallFiles.map((item, index) => (
                            <div key={index} className="flex items-center gap-2 p-2 bg-white rounded-lg border border-purple-100 shadow-sm">
                              <div className="flex-1 min-w-0">
                                <p className="text-xs text-gray-500 truncate" title={item.file.name}>{item.file.name}</p>
                              </div>
                              <input
                                type="text"
                                value={item.label}
                                onChange={(e) => {
                                  const newFiles = [...smallFiles];
                                  newFiles[index].label = e.target.value;
                                  setSmallFiles(newFiles);
                                }}
                                className="text-sm border border-gray-200 rounded px-2 py-1 focus:outline-none focus:border-purple-400 w-28 sm:w-32"
                                placeholder="Label"
                              />
                              <button
                                type="button"
                                onClick={() => setSmallFiles(smallFiles.filter((_, i) => i !== index))}
                                className="text-red-400 hover:text-red-600 px-1 font-bold text-lg leading-none"
                              >
                                &times;
                              </button>
                            </div>
                          ))}
                        </div>
                        <button
                          type="button"
                          onClick={() => smallInputRef.current?.click()}
                          className="w-full mt-2 py-2 text-sm text-purple-600 bg-white hover:bg-purple-100 rounded-lg font-medium transition-colors border border-purple-200 border-dashed shadow-sm"
                        >
                          + Add More Reports
                        </button>
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
                  className={`relative overflow-hidden w-full py-5 rounded-2xl font-bold text-lg text-white shadow-xl transition-all duration-300 flex items-center justify-center gap-3 ${loading || !masterFile || smallFiles.length === 0
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

              <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                <div className="bg-white rounded-2xl p-4 text-center shadow-sm border border-gray-100 hover:shadow-md transition-shadow">
                  <p className="text-xs text-gray-500 font-bold uppercase tracking-wider mb-2">Total (Master)</p>
                  <p className="text-2xl font-black text-gray-900">{result.stats.totalMaster}</p>
                </div>

                <div className="bg-gradient-to-br from-emerald-50 to-green-50 rounded-2xl p-4 text-center border border-green-100 shadow-sm shadow-green-100/50 hover:shadow-md transition-shadow relative overflow-hidden">
                  <p className="text-xs text-green-700 font-bold uppercase tracking-wider mb-2 relative z-10">Matched</p>
                  <p className="text-2xl font-black text-green-700 relative z-10">{result.stats.matched}</p>
                </div>

                <div className="bg-gradient-to-br from-rose-50 to-red-50 rounded-2xl p-4 text-center border border-red-100 shadow-sm shadow-red-100/50 hover:shadow-md transition-shadow relative overflow-hidden">
                  <p className="text-xs text-red-700 font-bold uppercase tracking-wider mb-2 relative z-10">Missing</p>
                  <p className="text-2xl font-black text-red-700 relative z-10">{result.stats.notFound}</p>
                </div>
                
                <div className="bg-gradient-to-br from-amber-50 to-yellow-50 rounded-2xl p-4 text-center border border-yellow-100 shadow-sm shadow-yellow-100/50 hover:shadow-md transition-shadow relative overflow-hidden">
                  <p className="text-xs text-yellow-700 font-bold uppercase tracking-wider mb-2 relative z-10">Extra</p>
                  <p className="text-2xl font-black text-yellow-700 relative z-10">{result.stats.extra || 0}</p>
                </div>
              </div>

              {/* Download Master File */}
              <div className="space-y-4 pt-6 border-t border-gray-100 mt-6">
                <h3 className="text-xl font-bold text-gray-900 text-center">Download Master Report</h3>
                <div className="flex justify-center">
                  {result.masterFile && (
                    <div className="w-full sm:w-auto flex flex-col sm:flex-row gap-3">
                      <button
                        onClick={() => downloadExcel(result.masterFile)}
                        className="flex-1 py-4 px-8 bg-gradient-to-r from-gray-900 to-gray-800 text-white rounded-2xl font-bold shadow-xl shadow-gray-900/20 hover:shadow-gray-900/30 hover:-translate-y-0.5 transition-all flex items-center justify-center gap-3"
                      >
                        <Download className="w-5 h-5 shrink-0" />
                        <span className="truncate">Download {result.masterFile.fileName}</span>
                      </button>
                      {result.masterFile.previewData && (
                        <button
                          onClick={() => setPreviewFile(result.masterFile)}
                          className="flex-1 py-4 px-8 bg-white border border-gray-200 text-gray-800 rounded-2xl font-bold shadow-sm hover:bg-gray-50 transition-all flex items-center justify-center gap-3"
                        >
                          <Eye className="w-5 h-5 shrink-0 text-indigo-600" />
                          <span>Preview Data</span>
                        </button>
                      )}
                    </div>
                  )}
                  {/* Fallback */}
                  {!result.masterFile && result.fileBase64 && (
                    <button
                      onClick={() => downloadExcel(result)}
                      className="w-full sm:w-auto py-4 px-8 bg-gradient-to-r from-gray-900 to-gray-800 text-white rounded-2xl font-bold shadow-xl shadow-gray-900/20 hover:shadow-gray-900/30 hover:-translate-y-0.5 transition-all flex items-center justify-center gap-3"
                    >
                      <Download className="w-5 h-5 shrink-0" />
                      <span className="truncate">Download {result.fileName}</span>
                    </button>
                  )}
                </div>
              </div>

              {/* Small Files Individual Downloads & Stats */}
              {result.smallFiles && result.smallFiles.length > 0 && (
                <div className="space-y-4 pt-6 border-t border-gray-100 mt-6">
                  <h3 className="text-xl font-bold text-gray-900 text-center">Small Reports Details & Downloads</h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {result.smallFiles.map((sf, idx) => (
                      <div key={idx} className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden flex flex-col hover:shadow-md transition-shadow">
                        <div className="p-4 border-b border-gray-50 flex items-center gap-3">
                          <FileText className="w-5 h-5 text-indigo-500 shrink-0" />
                          <h4 className="font-bold text-gray-800 truncate" title={sf.fileName}>{sf.fileName}</h4>
                        </div>
                        <div className="grid grid-cols-2 gap-2 p-4 bg-gray-50/50">
                          <div className="text-center p-2 bg-green-50 rounded-lg">
                            <p className="text-xs text-green-700 font-bold uppercase mb-1">Matched</p>
                            <p className="text-xl font-black text-green-700">{sf.matched}</p>
                          </div>
                          <div className="text-center p-2 bg-red-50 rounded-lg">
                            <p className="text-xs text-red-700 font-bold uppercase mb-1">Unmatched</p>
                            <p className="text-xl font-black text-red-700">{sf.notFound}</p>
                          </div>
                        </div>
                        <div className="p-4 mt-auto flex flex-col gap-2">
                          <button
                            onClick={() => downloadExcel(sf)}
                            className="w-full py-3 px-4 bg-gradient-to-r from-indigo-600 to-purple-600 text-white rounded-xl font-bold shadow-md shadow-indigo-500/20 hover:shadow-indigo-500/30 hover:-translate-y-0.5 transition-all flex items-center justify-center gap-2"
                          >
                            <Download className="w-4 h-4 shrink-0" />
                            Download File
                          </button>
                          {sf.previewData && (
                            <button
                              onClick={() => setPreviewFile(sf)}
                              className="w-full py-3 px-4 bg-white border border-gray-200 text-gray-700 rounded-xl font-bold hover:bg-gray-50 transition-all flex items-center justify-center gap-2"
                            >
                              <Eye className="w-4 h-4 shrink-0 text-indigo-600" />
                              Preview Data
                            </button>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              <div className="pt-6 border-t border-gray-100 mt-6">
                <button
                  onClick={reset}
                  className="w-full py-4 px-6 bg-white border-2 border-gray-200 text-gray-700 rounded-2xl font-bold hover:bg-gray-50 hover:border-gray-300 transition-all focus:ring-4 focus:ring-gray-100"
                >
                  Start Over
                </button>
              </div>
            </div>
            )
          ) : (
            /* VALIDATION UI */
            !validationResult ? (
              <div className="space-y-10">
                {/* Validation Upload */}
                <div className="space-y-4 relative group max-w-lg mx-auto">
                  <div className="flex items-center justify-center">
                    <h3 className="text-lg font-bold text-gray-800">Upload Numbers Sheet</h3>
                  </div>

                  <div
                    className={`relative cursor-pointer overflow-hidden border-2 border-dashed rounded-2xl p-8 flex flex-col items-center justify-center text-center transition-all duration-300 ${validationFile
                        ? 'border-indigo-400 bg-indigo-50/50 shadow-inner'
                        : 'border-gray-200 bg-gray-50/50 hover:bg-white hover:border-indigo-300 hover:shadow-lg hover:-translate-y-1'
                      }`}
                    onClick={() => validationInputRef.current?.click()}
                  >
                    <input type="file" className="hidden" accept=".xlsx, .xls, .csv" ref={validationInputRef} onChange={handleValidationUpload} />

                    {validationFile ? (
                      <div className="animate-in zoom-in-95 duration-300 flex flex-col items-center">
                        <div className="w-16 h-16 bg-white rounded-full shadow-sm flex items-center justify-center mb-4">
                          <CheckCircle className="w-8 h-8 text-indigo-500" />
                        </div>
                        <p className="text-sm font-bold text-gray-900 truncate w-full px-4">{validationFile.name}</p>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center">
                        <div className="w-16 h-16 bg-white rounded-full shadow-sm flex items-center justify-center mb-4 group-hover:scale-110 transition-transform duration-300">
                          <UploadCloud className="w-8 h-8 text-gray-400 group-hover:text-indigo-500 transition-colors" />
                        </div>
                        <p className="text-sm font-semibold text-gray-700">Drop your file here</p>
                        <p className="text-xs text-gray-500 mt-1">.xlsx, .csv</p>
                      </div>
                    )}
                  </div>
                </div>

                <div className="pt-4 border-t border-gray-100 max-w-lg mx-auto">
                  <button
                    onClick={handleValidation}
                    disabled={loading || !validationFile}
                    className={`relative overflow-hidden w-full py-5 rounded-2xl font-bold text-lg text-white shadow-xl transition-all duration-300 flex items-center justify-center gap-3 ${loading || !validationFile
                        ? 'bg-gray-300 cursor-not-allowed shadow-none'
                        : 'bg-gradient-to-r from-indigo-600 via-purple-600 to-indigo-600 bg-[length:200%_100%] hover:bg-[100%_0] hover:shadow-indigo-500/25 hover:-translate-y-0.5'
                      }`}
                  >
                    {loading ? (
                      <>
                        <RefreshCcw className="w-6 h-6 animate-spin" />
                        Validating...
                      </>
                    ) : (
                      <>
                        Validate Numbers
                      </>
                    )}
                  </button>
                </div>
              </div>
            ) : (
              <div className="space-y-10 animate-in fade-in slide-in-from-bottom-8 duration-500">
                <div className="text-center space-y-4">
                  <div className="inline-flex items-center justify-center w-20 h-20 rounded-full bg-gradient-to-tr from-indigo-400 to-purple-500 shadow-lg shadow-indigo-500/30 mb-2">
                    <Activity className="w-10 h-10 text-white" />
                  </div>
                  <h2 className="text-3xl font-extrabold text-gray-900 tracking-tight">Validation Complete</h2>
                </div>

                <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                  <div className="bg-white rounded-2xl p-4 text-center shadow-sm border border-gray-100">
                    <p className="text-xs text-gray-500 font-bold uppercase mb-1">Total</p>
                    <p className="text-2xl font-black text-gray-900">{validationResult.stats.total}</p>
                  </div>
                  <div className="bg-green-50 rounded-2xl p-4 text-center border border-green-100">
                    <p className="text-xs text-green-700 font-bold uppercase mb-1">Mobile</p>
                    <p className="text-2xl font-black text-green-700">{validationResult.stats.mobile}</p>
                  </div>
                  <div className="bg-red-50 rounded-2xl p-4 text-center border border-red-100">
                    <p className="text-xs text-red-700 font-bold uppercase mb-1">VoIP</p>
                    <p className="text-2xl font-black text-red-700">{validationResult.stats.voip}</p>
                  </div>
                  <div className="bg-blue-50 rounded-2xl p-4 text-center border border-blue-100">
                    <p className="text-xs text-blue-700 font-bold uppercase mb-1">Landline</p>
                    <p className="text-2xl font-black text-blue-700">{validationResult.stats.landline}</p>
                  </div>
                </div>

                <div className="flex justify-center gap-4 pt-4">
                  <button
                    onClick={() => downloadExcel(validationResult.file)}
                    className="py-4 px-8 bg-gradient-to-r from-indigo-600 to-purple-600 text-white rounded-2xl font-bold shadow-xl shadow-indigo-500/20 flex items-center gap-3 hover:-translate-y-0.5 transition-all"
                  >
                    <Download className="w-5 h-5 shrink-0" />
                    Download Validated Sheet
                  </button>
                  <button
                    onClick={() => setPreviewFile(validationResult.file)}
                    className="py-4 px-8 bg-white border border-gray-200 text-gray-800 rounded-2xl font-bold shadow-sm hover:bg-gray-50 flex items-center gap-3 transition-all"
                  >
                    <Eye className="w-5 h-5 text-indigo-600" />
                    Preview
                  </button>
                </div>

                {/* Active Log */}
                {validationResult.activeLog && validationResult.activeLog.length > 0 && (
                  <div className="mt-8">
                    <h3 className="text-lg font-bold text-gray-800 mb-4">Recent Validations (Active Log)</h3>
                    <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
                      <div className="max-h-64 overflow-y-auto">
                        <table className="w-full text-sm text-left">
                          <thead className="bg-gray-50 text-gray-500 font-semibold sticky top-0">
                            <tr>
                              <th className="p-3 border-b">Time</th>
                              <th className="p-3 border-b">Phone</th>
                              <th className="p-3 border-b">Type</th>
                              <th className="p-3 border-b">Carrier</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-100">
                            {validationResult.activeLog.map((log, i) => (
                              <tr key={i} className="hover:bg-gray-50">
                                <td className="p-3 text-gray-500">{new Date(log.timestamp).toLocaleTimeString()}</td>
                                <td className="p-3 font-medium text-gray-800">{log.phone}</td>
                                <td className="p-3">
                                  <span className={`px-2 py-1 rounded text-xs font-bold ${log.type === 'voip' ? 'bg-red-100 text-red-700' : log.type === 'mobile' ? 'bg-green-100 text-green-700' : 'bg-blue-100 text-blue-700'}`}>
                                    {log.type}
                                  </span>
                                </td>
                                <td className="p-3 text-gray-600">{log.carrier}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                )}

                <div className="pt-6 border-t border-gray-100 mt-6">
                  <button
                    onClick={reset}
                    className="w-full py-4 px-6 bg-white border-2 border-gray-200 text-gray-700 rounded-2xl font-bold hover:bg-gray-50 hover:border-gray-300 transition-all focus:ring-4 focus:ring-gray-100"
                  >
                    Start Over
                  </button>
                </div>

              </div>
            )
          )}

        </div>
      </div>

      {/* Preview Modal */}
      {previewFile && previewFile.previewData && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-gray-900/50 backdrop-blur-sm p-4 animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-6xl max-h-[90vh] flex flex-col overflow-hidden animate-in slide-in-from-bottom-8">
            <div className="flex flex-col p-6 border-b border-gray-100 bg-gray-50/50 gap-4">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <FileSpreadsheet className="w-6 h-6 text-indigo-600" />
                  <div>
                    <h3 className="text-xl font-bold text-gray-900">{previewFile.fileName}</h3>
                    <p className="text-sm text-gray-500 font-medium">Data Preview</p>
                  </div>
                </div>
                <button 
                  onClick={() => setPreviewFile(null)}
                  className="p-2 hover:bg-gray-200 rounded-full transition-colors text-gray-500 hover:text-gray-900"
                >
                  <X className="w-6 h-6" />
                </button>
              </div>
              <div className="flex gap-4 flex-wrap">
                  <select className="border border-gray-300 rounded px-3 py-1.5 text-sm" value={filterStatus} onChange={e => setFilterStatus(e.target.value)}>
                      <option value="">All Statuses</option>
                      <option value="MATCHED">Matched</option>
                      <option value="MISSING">Missing</option>
                      <option value="EXTRA">Extra</option>
                  </select>
                  <input type="date" className="border border-gray-300 rounded px-3 py-1.5 text-sm" value={filterDate} onChange={e => setFilterDate(e.target.value)} placeholder="Filter Date" />
                  <input type="time" className="border border-gray-300 rounded px-3 py-1.5 text-sm" value={filterTime} onChange={e => setFilterTime(e.target.value)} placeholder="Filter Time" />
                  <input type="text" className="border border-gray-300 rounded px-3 py-1.5 text-sm" value={filterExt} onChange={e => setFilterExt(e.target.value)} placeholder="Filter Extension" />
                  
                  <select className="border border-gray-300 rounded px-3 py-1.5 text-sm" value={filterTz} onChange={e => setFilterTz(e.target.value)}>
                      <option value="">All Time Zones</option>
                      <option value="USA/Canada - Eastern">US Eastern (EDT/EST)</option>
                      <option value="USA/Canada - Central">US Central (CDT/CST)</option>
                      <option value="USA/Canada - Mountain">US Mountain (MDT/MST)</option>
                      <option value="USA/Canada - Pacific">US Pacific (PDT/PST)</option>
                      <option value="United Kingdom">United Kingdom (GMT/BST)</option>
                      <option value="Central Europe">Central Europe (CET/CEST)</option>
                      <option value="India">India (IST)</option>
                      <option value="Australia">Australia</option>
                      <option value="Africa">Africa</option>
                  </select>
              </div>
            </div>
            <div className="flex-1 overflow-auto p-0">
              {previewFile.previewData.length > 0 ? (
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="bg-gray-100 sticky top-0 z-10 shadow-sm">
                    <tr>
                      {Object.keys(previewFile.previewData[0]).map((key, i) => (
                        <th key={i} className="p-3 font-semibold text-gray-700 border-b border-gray-200 whitespace-nowrap">
                          {key}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {(previewFile.previewData || [])
                    .filter(row => {
                        if (filterStatus && row['Match Status'] && !row['Match Status'].includes(filterStatus)) return false;
                        if (filterDate && row['Date'] && !String(row['Date']).includes(filterDate)) return false;
                        if (filterExt && row['Extension'] && !String(row['Extension']).includes(filterExt)) return false;
                        if (filterTz && row['Time Zone'] && !String(row['Time Zone']).toLowerCase().includes(filterTz.toLowerCase())) return false;
                        if (filterTime && row['Time'] && !String(row['Time']).includes(filterTime)) return false;
                        return true;
                    })
                    .map((row, i) => (
                      <tr 
                        key={i} 
                        className={`hover:bg-gray-50 transition-colors ${row['Match Status']?.includes('MATCHED') ? 'bg-green-50/30' : row['Match Status']?.includes('MISSING') ? 'bg-red-50/30' : row['Match Status']?.includes('EXTRA') ? 'bg-yellow-50/30' : ''}`}
                      >
                        {Object.keys(previewFile.previewData[0]).map((key, j) => (
                          <td 
                            key={j} 
                            className={`p-3 whitespace-nowrap ${key === 'Match Status' ? (row[key]?.includes('MATCHED') ? 'text-green-700 font-bold' : row[key]?.includes('MISSING') ? 'text-red-700 font-bold' : row[key]?.includes('EXTRA') ? 'text-yellow-700 font-bold' : 'text-gray-600') : 'text-gray-600'}`}
                          >
                            {row[key]}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <div className="p-12 text-center text-gray-500 font-medium">
                  No data available for preview.
                </div>
              )}
            </div>
            <div className="p-4 border-t border-gray-100 bg-gray-50 flex justify-end gap-3">
              <button
                onClick={() => setPreviewFile(null)}
                className="py-2 px-6 bg-white border border-gray-300 text-gray-700 rounded-xl font-bold hover:bg-gray-50 transition-all"
              >
                Close
              </button>
              <button
                onClick={() => downloadExcel(previewFile)}
                className="py-2 px-6 bg-indigo-600 text-white rounded-xl font-bold shadow-md shadow-indigo-500/20 hover:bg-indigo-700 hover:-translate-y-0.5 transition-all flex items-center gap-2"
              >
                <Download className="w-4 h-4" />
                Download Full File
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
