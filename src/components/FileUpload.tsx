import React, { useRef, useState } from "react";
import { UploadCloud, FileSpreadsheet, AlertCircle, CheckSquare, Square, RefreshCcw } from "lucide-react";
import { cn } from "../lib/utils";

interface FileUploadProps {
  onFileUpload: (file: File, clsFile?: File) => void;
  isLoading: boolean;
}

export function FileUpload({ onFileUpload, isLoading }: FileUploadProps) {
  const [isDragActive, setIsDragActive] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [checkCLS, setCheckCLS] = useState(false);
  
  const [mainFile, setMainFile] = useState<File | null>(null);
  const [clsFile, setClsFile] = useState<File | null>(null);

  const mainInputRef = useRef<HTMLInputElement>(null);
  const clsInputRef = useRef<HTMLInputElement>(null);

  const handleDragEnter = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);
  };

  const isValidType = (file: File) => {
    return file.name.endsWith(".xlsx") || file.name.endsWith(".xls") || file.name.endsWith(".csv");
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (!isValidType(file)) {
        setError("Vui lòng chọn file định dạng Excel (.xlsx, .xls) hoặc CSV.");
        return;
      }
      setMainFile(file);
      setError(null);
    }
  };

  const handleMainChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (!isValidType(file)) {
        setError("Vui lòng chọn file định dạng Excel (.xlsx, .xls) hoặc CSV.");
        return;
      }
      setMainFile(file);
      setError(null);
    }
    if (mainInputRef.current) mainInputRef.current.value = '';
  };

  const handleClsChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (!isValidType(file)) {
        setError("Vui lòng chọn file định dạng Excel (.xlsx, .xls) hoặc CSV cho file CLS.");
        return;
      }
      setClsFile(file);
      setError(null);
    }
    if (clsInputRef.current) clsInputRef.current.value = '';
  };

  const handleSubmit = () => {
    if (!mainFile) {
      setError("Vui lòng tải lên file Sổ thủ thuật/phẫu thuật.");
      return;
    }
    if (checkCLS && !clsFile) {
      setError("Vui lòng tải lên file dịch vụ Cận Lâm Sàng.");
      return;
    }
    onFileUpload(mainFile, checkCLS && clsFile ? clsFile : undefined);
  };

  return (
    <div className="w-full max-w-2xl mx-auto mt-12 p-6 bg-white rounded-2xl shadow-xl shadow-slate-200/50">
      <div className="text-center mb-6">
        <h2 className="text-2xl font-bold text-slate-800 tracking-tight">
          Phân Tích File Lịch Thủ Thuật
        </h2>
        <p className="mt-2 text-slate-500">
          Upload file Excel chứa thông tin thủ thuật y tế để tự động phát hiện trùng lặp.
        </p>
      </div>

      <div className="space-y-6">
        <div>
          <h3 className="text-sm font-semibold text-slate-700 mb-2">1. File Sổ thủ thuật/phẫu thuật</h3>
          <div
            className={cn(
              "relative group border-2 border-dashed rounded-xl p-8 cursor-pointer transition-all duration-200 flex flex-col items-center justify-center",
              isDragActive ? "border-blue-500 bg-blue-50/50" : "border-slate-300 hover:border-blue-400 hover:bg-slate-50"
            )}
            onDragEnter={handleDragEnter}
            onDragOver={handleDragEnter}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => mainInputRef.current?.click()}
          >
            <input
              ref={mainInputRef}
              type="file"
              className="hidden"
              accept=".xlsx, .xls, .csv"
              onChange={handleMainChange}
              disabled={isLoading}
            />
            
            {mainFile ? (
              <div className="flex items-center text-blue-600 font-medium">
                <FileSpreadsheet className="w-8 h-8 mr-3" />
                {mainFile.name}
              </div>
            ) : (
              <div className="flex flex-col items-center">
                <UploadCloud className="w-8 h-8 text-blue-500 mb-3" />
                <p className="text-sm font-semibold text-slate-700 mb-1">
                  Kéo thả file vào đây hoặc <span className="text-blue-600">chọn file</span>
                </p>
              </div>
            )}
          </div>
          
          <div className="bg-amber-50 text-amber-700 text-xs p-3 rounded-lg flex items-start mt-3 border border-amber-100">
            <AlertCircle className="w-4 h-4 mr-2 mt-0.5 flex-shrink-0" />
            <span>
              <strong>Đường dẫn tải file:</strong> Báo cáo và tra cứu &rarr; Phẫu thuật thủ thuật &rarr; 01. Báo cáo nhân sự tham gia PTTT chi tiết (YHCT BLC) &rarr; Chọn đối tượng BHYT &rarr; Xuất file excel
            </span>
          </div>
        </div>

        <div className="flex items-center">
          <button 
            type="button" 
            onClick={() => setCheckCLS(!checkCLS)}
            className="flex items-center text-slate-700 hover:text-blue-600 transition-colors"
          >
            {checkCLS ? <CheckSquare className="w-5 h-5 text-blue-600 mr-2" /> : <Square className="w-5 h-5 text-slate-400 mr-2" />}
            <span className="font-medium text-sm">Kiểm tra kèm dịch vụ Cận Lâm Sàng (CLS)</span>
          </button>
        </div>

        {checkCLS && (
          <div className="animate-in slide-in-from-top-2 fade-in duration-300">
            <h3 className="text-sm font-semibold text-slate-700 mb-2">2. File Dịch vụ Cận Lâm Sàng</h3>
            <div
              className={cn(
                "relative group border-2 border-dashed rounded-xl p-6 cursor-pointer transition-all duration-200 flex flex-col items-center justify-center",
                "border-slate-300 hover:border-blue-400 hover:bg-slate-50"
              )}
              onClick={() => clsInputRef.current?.click()}
            >
              <input
                ref={clsInputRef}
                type="file"
                className="hidden"
                accept=".xlsx, .xls, .csv"
                onChange={handleClsChange}
                disabled={isLoading}
              />
              {clsFile ? (
                <div className="flex items-center text-blue-600 font-medium">
                  <FileSpreadsheet className="w-6 h-6 mr-3" />
                  {clsFile.name}
                </div>
              ) : (
                <div className="flex flex-col items-center">
                  <UploadCloud className="w-6 h-6 text-slate-400 group-hover:text-blue-500 mb-2 transition-colors" />
                  <p className="text-sm text-slate-600">
                    Bấm để <span className="text-blue-600 font-medium">chọn file CLS</span>
                  </p>
                </div>
              )}
            </div>
          </div>
        )}

        {error && (
          <div className="p-3 bg-red-50 text-red-600 rounded-lg text-sm flex items-center">
            <AlertCircle className="w-4 h-4 mr-2" />
            {error}
          </div>
        )}

        <button
          onClick={handleSubmit}
          disabled={isLoading || !mainFile}
          className={cn(
            "w-full py-3 rounded-xl font-medium transition-all flex items-center justify-center",
            isLoading || !mainFile 
              ? "bg-slate-100 text-slate-400 cursor-not-allowed" 
              : "bg-blue-600 hover:bg-blue-700 text-white shadow-lg shadow-blue-200"
          )}
        >
          {isLoading ? (
            <>
              <RefreshCcw className="w-5 h-5 mr-3 animate-spin" />
              Đang phân tích...
            </>
          ) : (
            "Phân tích dữ liệu"
          )}
        </button>
      </div>
    </div>
  );
}
