import React, { useState } from 'react';
import { FileUpload } from './components/FileUpload';
import { Dashboard } from './components/Dashboard';
import { analyzeExcelData } from './lib/analyzer';
import { AnalysisResult } from './types';

export default function App() {
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleFileUpload = async (file: File, clsFile?: File) => {
    setIsAnalyzing(true);
    setError(null);
    try {
      // Simulate slight delay for UX
      await new Promise(r => setTimeout(r, 600)); 
      const analysisResult = await analyzeExcelData(file, clsFile);
      setResult(analysisResult);
    } catch (err) {
      if (err instanceof Error) {
        setError(err.message);
      } else {
        setError("Đã xảy ra lỗi không xác định.");
      }
    } finally {
      setIsAnalyzing(false);
    }
  };

  if (result) {
    return <Dashboard result={result} onReset={() => setResult(null)} />;
  }

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-4">
      {/* Background decoration */}
      <div className="absolute inset-0 z-0 overflow-hidden pointer-events-none">
        <div className="absolute -top-[20%] -left-[10%] w-[50%] h-[50%] rounded-full bg-blue-100/40 blur-3xl" />
        <div className="absolute top-[60%] -right-[10%] w-[50%] h-[50%] rounded-full bg-indigo-100/40 blur-3xl" />
      </div>

      <div className="relative z-10 w-full">
        {error && (
          <div className="max-w-2xl mx-auto mb-4 bg-red-50 text-red-600 border border-red-200 px-4 py-3 rounded-xl flex items-start">
            <svg className="w-5 h-5 mr-3 mt-0.5 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
            <div className="font-medium text-sm">{error}</div>
          </div>
        )}
        <FileUpload onFileUpload={handleFileUpload} isLoading={isAnalyzing} />
      </div>
      
      <div className="mt-12 text-center text-slate-400 text-sm z-10">
        Mọi phân tích được thực hiện cục bộ trên trình duyệt của bạn.<br />
        Dữ liệu y tế riêng tư không được gửi tới bất kỳ máy chủ nào.
      </div>
    </div>
  );
}
