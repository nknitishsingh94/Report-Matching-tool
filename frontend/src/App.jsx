import React, { useState, useRef } from 'react';
import { UploadCloud, FileSpreadsheet, CheckCircle, AlertCircle, Download, RefreshCcw } from 'lucide-react';

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
      setSmallFiles(Array.from(e.target.files));
    }
  };

  const handleMatch = async () => {
    if (!masterFile) {
      setError("Please upload the Master Report.");
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
    
    // Convert base64 to blob
    const byteCharacters = atob(result.fileBase64);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    // Create download link
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
    <div className="min-h-screen bg-gray-50 py-12 px-4 sm:px-6 lg:px-8 font-sans">
      <div className="max-w-4xl mx-auto space-y-8">
        
        {/* Header */}
        <div className="text-center">
          <h1 className="text-3xl sm:text-4xl font-extrabold text-gray-900 tracking-tight flex flex-col sm:flex-row items-center justify-center gap-2 sm:gap-3">
            <FileSpreadsheet className="w-10 h-10 text-indigo-600" />
            Report Reconciliation
          </h1>
          <p className="mt-3 text-base sm:text-lg text-gray-500">
            Easily match and update Caller IDs across multiple Excel reports.
          </p>
        </div>

        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-gray-100 p-8">
          
          {error && (
            <div className="mb-6 p-4 bg-red-50 rounded-lg border border-red-200 flex items-start gap-3">
              <AlertCircle className="w-5 h-5 text-red-600 mt-0.5" />
              <p className="text-red-700">{error}</p>
            </div>
          )}

          {!result ? (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
              
              {/* Master Upload */}
              <div className="space-y-4">
                <h3 className="text-lg font-semibold text-gray-800 flex items-center gap-2">
                  <span className="bg-indigo-100 text-indigo-800 w-6 h-6 rounded-full flex items-center justify-center text-sm">1</span>
                  Upload Master Report
                </h3>
                <div 
                  className={`border-2 border-dashed rounded-xl p-6 flex flex-col items-center justify-center text-center transition-colors ${masterFile ? 'border-green-400 bg-green-50' : 'border-gray-300 hover:border-indigo-400 bg-gray-50 hover:bg-indigo-50'}`}
                  onClick={() => masterInputRef.current?.click()}
                >
                  <input type="file" className="hidden" accept=".xlsx, .xls" ref={masterInputRef} onChange={handleMasterUpload} />
                  {masterFile ? (
                    <>
                      <CheckCircle className="w-10 h-10 text-green-500 mb-2" />
                      <p className="text-sm font-medium text-green-800 truncate w-full">{masterFile.name}</p>
                      <p className="text-xs text-green-600 mt-1">{(masterFile.size / 1024).toFixed(1)} KB</p>
                    </>
                  ) : (
                    <>
                      <UploadCloud className="w-10 h-10 text-gray-400 mb-2" />
                      <p className="text-sm font-medium text-gray-700">Click to upload Master File</p>
                      <p className="text-xs text-gray-500 mt-1">Excel formats only (.xlsx, .xls)</p>
                    </>
                  )}
                </div>
              </div>

              {/* Small Reports Upload */}
              <div className="space-y-4">
                <h3 className="text-lg font-semibold text-gray-800 flex items-center gap-2">
                  <span className="bg-indigo-100 text-indigo-800 w-6 h-6 rounded-full flex items-center justify-center text-sm">2</span>
                  Upload Small Reports
                </h3>
                <div 
                  className={`border-2 border-dashed rounded-xl p-6 flex flex-col items-center justify-center text-center transition-colors ${smallFiles.length > 0 ? 'border-blue-400 bg-blue-50' : 'border-gray-300 hover:border-indigo-400 bg-gray-50 hover:bg-indigo-50'}`}
                  onClick={() => smallInputRef.current?.click()}
                >
                  <input type="file" className="hidden" accept=".xlsx, .xls" multiple ref={smallInputRef} onChange={handleSmallUpload} />
                  {smallFiles.length > 0 ? (
                    <>
                      <FileSpreadsheet className="w-10 h-10 text-blue-500 mb-2" />
                      <p className="text-sm font-medium text-blue-800">{smallFiles.length} files selected</p>
                      <p className="text-xs text-blue-600 mt-1 text-left max-w-full truncate">
                        {smallFiles.map(f => f.name).join(', ')}
                      </p>
                    </>
                  ) : (
                    <>
                      <UploadCloud className="w-10 h-10 text-gray-400 mb-2" />
                      <p className="text-sm font-medium text-gray-700">Click to upload multiple files</p>
                      <p className="text-xs text-gray-500 mt-1">Excel formats only (.xlsx, .xls)</p>
                    </>
                  )}
                </div>
              </div>

              {/* Action Button */}
              <div className="md:col-span-2 pt-4">
                <button
                  onClick={handleMatch}
                  disabled={loading || !masterFile || smallFiles.length === 0}
                  className={`w-full py-4 rounded-xl font-bold text-lg text-white shadow-lg transition-all flex items-center justify-center gap-2 ${
                    loading || !masterFile || smallFiles.length === 0
                      ? 'bg-gray-400 cursor-not-allowed' 
                      : 'bg-indigo-600 hover:bg-indigo-700 hover:shadow-indigo-200'
                  }`}
                >
                  {loading ? (
                    <>
                      <RefreshCcw className="w-6 h-6 animate-spin" />
                      Processing Reports...
                    </>
                  ) : (
                    <>Match Reports</>
                  )}
                </button>
              </div>
            </div>
          ) : (
            /* Results View */
            <div className="space-y-8 animate-in fade-in zoom-in duration-300">
              <div className="text-center">
                <div className="inline-flex items-center justify-center w-16 h-16 rounded-full bg-green-100 mb-4">
                  <CheckCircle className="w-8 h-8 text-green-600" />
                </div>
                <h2 className="text-2xl font-bold text-gray-900">Reconciliation Complete</h2>
                <p className="text-gray-500 mt-2">Your master report has been successfully updated.</p>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <div className="bg-gray-50 rounded-xl p-4 text-center border border-gray-100">
                  <p className="text-sm text-gray-500 font-medium">Total Records</p>
                  <p className="text-3xl font-bold text-gray-900 mt-1">{result.stats.totalMaster}</p>
                </div>
                <div className="bg-green-50 rounded-xl p-4 text-center border border-green-100">
                  <p className="text-sm text-green-600 font-medium">Matched</p>
                  <p className="text-3xl font-bold text-green-700 mt-1">{result.stats.matched}</p>
                </div>
                <div className="bg-red-50 rounded-xl p-4 text-center border border-red-100">
                  <p className="text-sm text-red-600 font-medium">Not Found</p>
                  <p className="text-3xl font-bold text-red-700 mt-1">{result.stats.notFound}</p>
                </div>
              </div>

              <div className="flex flex-col sm:flex-row gap-4 pt-4">
                <button
                  onClick={reset}
                  className="flex-1 py-3 px-4 bg-white border-2 border-gray-200 text-gray-700 rounded-xl font-semibold hover:bg-gray-50 transition-colors"
                >
                  Start Over
                </button>
                <button
                  onClick={downloadExcel}
                  className="flex-1 py-3 px-4 bg-green-600 text-white rounded-xl font-bold shadow-lg shadow-green-200 hover:bg-green-700 transition-colors flex items-center justify-center gap-2"
                >
                  <Download className="w-5 h-5" />
                  Download Final Excel
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
