import React, { useMemo, useState } from "react";
import { AnalysisResult, Conflict, ProcessedProcedure } from "../types";
import { formatDateTime } from "../lib/utils";
import { AlertTriangle, Clock, RefreshCcw, User, UserX, FileWarning, Search, ChevronRight, Download } from "lucide-react";
import { differenceInMinutes, isValid } from "date-fns";
import * as XLSX from "xlsx";

interface DashboardProps {
  result: AnalysisResult;
  onReset: () => void;
}

export function Dashboard({ result, onReset }: DashboardProps) {
  const [activeTab, setActiveTab] = useState<'overview' | 'conflicts' | 'data'>('overview');
  const [searchTerm, setSearchTerm] = useState("");

  const filteredConflicts = useMemo(() => {
    if (!searchTerm) return result.conflicts;
    const lower = searchTerm.toLowerCase();
    return result.conflicts.filter(c => 
      c.message.toLowerCase().includes(lower) || 
      c.procedures.some(p => p.patientName.toLowerCase().includes(lower) || p.doctorName.toLowerCase().includes(lower))
    );
  }, [result.conflicts, searchTerm]);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 pb-12">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 px-6 py-4 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-blue-600 rounded-lg text-white">
              <ActivityIcon className="w-5 h-5" />
            </div>
            <h1 className="text-xl font-bold tracking-tight">Check Thủ thuật YHCT Bảo Lộc</h1>
          </div>
          <button 
            onClick={onReset}
            className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors"
          >
            <RefreshCcw className="w-4 h-4" />
            Phân tích file khác
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto mt-8 px-6">
        {/* Navigation */}
        <div className="flex space-x-1 mb-8 bg-slate-200/50 p-1 rounded-xl w-fit">
          <TabButton active={activeTab === 'overview'} onClick={() => setActiveTab('overview')}>
            Tổng quan
          </TabButton>
          <TabButton active={activeTab === 'conflicts'} onClick={() => setActiveTab('conflicts')}>
            <span className="flex items-center gap-2">
              Danh sách cảnh báo 
              {result.conflicts.length > 0 && (
                <span className="bg-red-500 text-white text-[10px] px-2 py-0.5 rounded-full font-bold">
                  {result.conflicts.length}
                </span>
              )}
            </span>
          </TabButton>
          <TabButton active={activeTab === 'data'} onClick={() => setActiveTab('data')}>
            Dữ liệu gốc
          </TabButton>
        </div>

        {/* Views */}
        {activeTab === 'overview' && <OverviewView result={result} />}
        {activeTab === 'conflicts' && (
          <ConflictsView 
            conflicts={filteredConflicts} 
            searchTerm={searchTerm} 
            setSearchTerm={setSearchTerm} 
            totalConflicts={result.conflicts.length}
          />
        )}
        {activeTab === 'data' && <DataView procedures={result.procedures} />}
      </main>
    </div>
  );
}

function TabButton({ active, onClick, children }: { active: boolean; onClick: () => void; children: React.ReactNode }) {
  return (
    <button
      onClick={onClick}
      className={`px-6 py-2.5 rounded-lg text-sm font-medium transition-all ${
        active ? 'bg-white text-blue-700 shadow-sm' : 'text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
      }`}
    >
      {children}
    </button>
  );
}

function OverviewView({ result }: { result: AnalysisResult }) {
  const patientConflicts = result.conflicts.filter(c => c.type === 'PATIENT_TIME_OVERLAP').length;
  const doctorConflicts = result.conflicts.filter(c => c.type === 'DOCTOR_TIME_OVERLAP').length;
  const machineOverlaps = result.conflicts.filter(c => c.type === 'MACHINE_TIME_OVERLAP').length;
  const missingMachines = result.conflicts.filter(c => c.type === 'MISSING_MACHINE_ID').length;
  const invalidTimeConflicts = result.conflicts.filter(c => c.type === 'INVALID_TIME').length;
  const clsConflicts = result.conflicts.filter(c => c.type === 'CLS_TIME_OVERLAP').length;

  return (
    <div className="space-y-6 animate-in slide-in-from-bottom-4 fade-in duration-500">
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
        <StatCard title="Tổng số bệnh nhân" value={result.totalPatients.toString()} icon={<User className="text-blue-500" />} />
        <StatCard title="Tổng số thủ thuật" value={result.procedures.length.toString()} icon={<ActivityIcon className="text-indigo-500" />} />
        <StatCard 
          title="Tổng cảnh báo" 
          value={result.conflicts.length.toString()} 
          icon={<AlertTriangle className={result.conflicts.length > 0 ? "text-red-500" : "text-emerald-500"} />} 
          isDanger={result.conflicts.length > 0}
        />
        <StatCard title="Tổng lỗi dữ liệu" value={result.errors.length.toString()} icon={<FileWarning className="text-amber-500" />} />
      </div>

      {result.errors.length > 0 && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-5 mb-8">
          <h3 className="text-amber-800 font-bold flex items-center gap-2 mb-3">
            <AlertTriangle className="w-5 h-5" />
            Lỗi phân tích dữ liệu ({result.errors.length})
          </h3>
          <ul className="text-sm text-amber-700 space-y-1 list-disc pl-5">
            {result.errors.slice(0, 5).map((err, i) => (
              <li key={i}>{err}</li>
            ))}
            {result.errors.length > 5 && (
              <li className="font-semibold italic">...và thêm {result.errors.length - 5} lỗi khác.</li>
            )}
          </ul>
        </div>
      )}

      {result.conflicts.length === 0 ? (
        <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-10 text-center">
          <div className="mx-auto w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mb-4">
            <svg className="w-8 h-8" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
            </svg>
          </div>
          <h3 className="text-xl font-bold text-emerald-800 mb-2">Phân tích hoàn tất! Dữ liệu của bạn rất chuẩn Xác.</h3>
          <p className="text-emerald-600">Không tìm thấy bất kỳ sự trùng lặp lịch nào giữa các bệnh nhân và bác sĩ.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div className="lg:col-span-2 space-y-6">
            <h2 className="text-lg font-bold text-slate-800">Tóm tắt loại cảnh báo</h2>
            <div className="grid gap-4 sm:grid-cols-2">
              <ConflictSummaryCard 
                type="Bệnh nhân trùng lịch" 
                count={patientConflicts} 
                icon={<UserX className="w-5 h-5" />} 
                description="Một bệnh nhân có 2 hoặc nhiều thủ thuật diễn ra cùng lúc."
                color="red"
              />
              <ConflictSummaryCard 
                type="Bác sĩ trùng lịch" 
                count={doctorConflicts} 
                icon={<Clock className="w-5 h-5" />} 
                description="Một nhân viên y tế thực hiện thủ thuật trên nhiều bệnh nhân cùng lúc."
                color="orange"
              />
              {machineOverlaps > 0 && (
                <ConflictSummaryCard 
                  type="Trùng mã máy" 
                  count={machineOverlaps} 
                  icon={<AlertTriangle className="w-5 h-5" />} 
                  description="Một máy được sử dụng cho nhiều bệnh nhân hoặc thủ thuật cùng lúc."
                  color="purple"
                />
              )}
              {clsConflicts > 0 && (
                <ConflictSummaryCard 
                  type="Trùng lịch CLS" 
                  count={clsConflicts} 
                  icon={<FileWarning className="w-5 h-5" />} 
                  description="Bệnh nhân có dịch vụ Cận Lâm Sàng trùng thời gian."
                  color="blue"
                />
              )}
              {missingMachines > 0 && (
                <ConflictSummaryCard 
                  type="Thiếu mã máy" 
                  count={missingMachines} 
                  icon={<FileWarning className="w-5 h-5" />} 
                  description="Các thủ thuật bắt buộc nhưng không có mã máy."
                  color="amber"
                />
              )}
              {invalidTimeConflicts > 0 && (
                <ConflictSummaryCard 
                  type="Lỗi thời gian" 
                  count={invalidTimeConflicts} 
                  icon={<FileWarning className="w-5 h-5" />} 
                  description="Dữ liệu thời gian không hợp lệ hoặc logic thời gian sai."
                  color="amber"
                />
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function ConflictsView({ conflicts, searchTerm, setSearchTerm, totalConflicts }: { conflicts: Conflict[], searchTerm: string, setSearchTerm: (s:string) => void, totalConflicts: number }) {
  const exportToExcel = () => {
    if (conflicts.length === 0) return;

    const data: any[] = [];
    conflicts.forEach(conflict => {
      conflict.procedures.forEach(proc => {
        data.push({
          "Loại Cảnh Báo": conflict.type === 'PATIENT_TIME_OVERLAP' ? 'Bệnh nhân trùng lịch' : 
                           conflict.type === 'DOCTOR_TIME_OVERLAP' ? 'Bác sĩ trùng lịch' : 
                           conflict.type === 'MACHINE_TIME_OVERLAP' ? 'Trùng mã máy' : 
                           conflict.type === 'CLS_TIME_OVERLAP' ? 'Trùng lịch CLS' : 
                           conflict.type === 'MISSING_MACHINE_ID' ? 'Thiếu mã máy' : 'Lỗi thời gian',
          "Nội dung": conflict.message.replace(/\*\*/g, ''), // remove markdown bold syntax
          "Thủ thuật": proc.procedureName,
          "Mã máy": proc.machineId || '',
          "Bệnh nhân": proc.patientName,
          "Nhân viên": proc.doctorName,
          "Từ thời gian": isValid(proc.startTime) ? formatDateTime(proc.startTime) : 'Lỗi',
          "Đến thời gian": isValid(proc.endTime) ? formatDateTime(proc.endTime) : 'Lỗi',
          "Dòng Excel": proc.rowNumber
        });
      });
    });

    const worksheet = XLSX.utils.json_to_sheet(data);
    
    // Auto-size columns appropriately
    const maxWidths: number[] = [];
    data.forEach(row => {
      Object.keys(row).forEach((key, i) => {
        const val = row[key as keyof typeof row] ? String(row[key as keyof typeof row]) : "";
        maxWidths[i] = Math.max(maxWidths[i] || key.length, val.length);
      });
    });
    worksheet['!cols'] = maxWidths.map(w => ({ wch: Math.min(w + 2, 50) }));

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "DanhSachCanhBao");
    
    // Generate file and trigger download
    XLSX.writeFile(workbook, "DanhSachCanhBao.xlsx");
  };

  return (
    <div className="animate-in slide-in-from-bottom-4 fade-in duration-500">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-6 gap-4">
        <div className="flex items-center gap-4">
          <h2 className="text-xl font-bold text-slate-800">Danh sách cảnh báo</h2>
          {conflicts.length > 0 && (
            <button 
              onClick={exportToExcel}
              className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-lg transition-colors"
            >
              <Download className="w-4 h-4" />
              Tải Excel
            </button>
          )}
        </div>
        <div className="relative">
          <Search className="w-5 h-5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input 
            type="text" 
            placeholder="Tìm kiếm theo tên bệnh nhân, thủ thuật..."
            className="pl-10 pr-4 py-2 border border-slate-300 rounded-lg w-full sm:w-80 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
      </div>

      {conflicts.length === 0 ? (
        <div className="py-12 text-center text-slate-500">
          {totalConflicts === 0 ? "Không có cảnh báo nào." : "Không tìm thấy kết quả phù hợp."}
        </div>
      ) : (
        <div className="space-y-4">
          {conflicts.map(conflict => {
            let borderColor = '#f59e0b';
            if (conflict.type === 'PATIENT_TIME_OVERLAP') borderColor = '#ef4444';
            else if (conflict.type === 'DOCTOR_TIME_OVERLAP') borderColor = '#f97316';
            else if (conflict.type === 'MACHINE_TIME_OVERLAP') borderColor = '#a855f7';
            else if (conflict.type === 'CLS_TIME_OVERLAP') borderColor = '#3b82f6';

            let iconBg = 'bg-amber-50 text-amber-500';
            let iconElement = <FileWarning className="w-5 h-5" />;
            if (conflict.type === 'PATIENT_TIME_OVERLAP') {
              iconBg = 'bg-red-50 text-red-500';
              iconElement = <UserX className="w-5 h-5" />;
            } else if (conflict.type === 'DOCTOR_TIME_OVERLAP') {
              iconBg = 'bg-orange-50 text-orange-500';
              iconElement = <Clock className="w-5 h-5" />;
            } else if (conflict.type === 'MACHINE_TIME_OVERLAP') {
              iconBg = 'bg-purple-50 text-purple-500';
              iconElement = <AlertTriangle className="w-5 h-5" />;
            } else if (conflict.type === 'CLS_TIME_OVERLAP') {
              iconBg = 'bg-blue-50 text-blue-500';
              iconElement = <FileWarning className="w-5 h-5" />;
            }

            return (
            <div key={conflict.id} className="bg-white border-l-4 border rounded-lg shadow-sm overflow-hidden" 
              style={{ borderLeftColor: borderColor }}>
              <div className="p-4 sm:p-5">
                <div className="flex items-start gap-4">
                  <div className={`mt-0.5 p-2 rounded-full flex-shrink-0 ${iconBg}`}>
                    {iconElement}
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="font-medium text-slate-800 leading-snug">
                      <MarkdownText text={conflict.message} />
                    </p>
                    <div className="mt-4 bg-slate-50 rounded-lg border border-slate-100 overflow-hidden">
                      <table className="w-full text-sm text-left">
                        <thead className="bg-slate-100 text-slate-600 text-xs uppercase">
                          <tr>
                            <th className="px-4 py-2">Thủ thuật</th>
                            <th className="px-4 py-2">Mã máy</th>
                            <th className="px-4 py-2">Bệnh nhân</th>
                            <th className="px-4 py-2">Nhân viên</th>
                            <th className="px-4 py-2">Thời gian</th>
                            <th className="px-4 py-2 whitespace-nowrap">Dòng Excel</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {conflict.procedures.map((proc, i) => (
                            <tr key={i} className="hover:bg-slate-100/50">
                              <td className="px-4 py-3 font-medium text-slate-700">{proc.procedureName}</td>
                              <td className="px-4 py-3 text-slate-500">{proc.machineId || '-'}</td>
                              <td className="px-4 py-3 text-slate-600">{proc.patientName}</td>
                              <td className="px-4 py-3 text-slate-600">{proc.doctorName}</td>
                              <td className="px-4 py-3 text-slate-600 whitespace-nowrap">
                                {formatDateTime(proc.startTime)}<br/><span className="text-slate-400">đến</span> {formatDateTime(proc.endTime)}
                              </td>
                              <td className="px-4 py-3 text-slate-500">#{proc.rowNumber}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          );
        })}
        </div>
      )}
    </div>
  );
}

function DataView({ procedures }: { procedures: ProcessedProcedure[] }) {
  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden animate-in slide-in-from-bottom-4 fade-in duration-500">
      <div className="overflow-x-auto">
        <table className="w-full text-sm text-left text-slate-600">
          <thead className="text-xs text-slate-700 uppercase bg-slate-50 border-b border-slate-200">
            <tr>
              <th className="px-4 py-4 w-16">Dòng</th>
              <th className="px-4 py-4">Mã BN</th>
              <th className="px-4 py-4">Họ tên</th>
              <th className="px-4 py-4">Thủ thuật</th>
              <th className="px-4 py-4">Mã máy</th>
              <th className="px-4 py-4">Người thực hiện</th>
              <th className="px-4 py-4">Từ thời gian</th>
              <th className="px-4 py-4">Đến thời gian</th>
              <th className="px-4 py-4 text-center">Thời gian (phút)</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {procedures.map((p, i) => {
              const hasValidTime = isValid(p.startTime) && isValid(p.endTime);
              const duration = hasValidTime ? differenceInMinutes(p.endTime, p.startTime) : '-';
              return (
                <tr key={i} className="hover:bg-slate-50 transition-colors">
                  <td className="px-4 py-3 font-medium text-slate-400">#{p.rowNumber}</td>
                  <td className="px-4 py-3 font-medium text-slate-900">{p.patientId}</td>
                  <td className="px-4 py-3">{p.patientName}</td>
                  <td className="px-4 py-3 text-blue-600">{p.procedureName}</td>
                  <td className="px-4 py-3 text-slate-500">{p.machineId || '-'}</td>
                  <td className="px-4 py-3">{p.doctorName}</td>
                  <td className="px-4 py-3 whitespace-nowrap">{formatDateTime(p.startTime)}</td>
                  <td className="px-4 py-3 whitespace-nowrap">{formatDateTime(p.endTime)}</td>
                  <td className="px-4 py-3 text-center font-medium">
                    {hasValidTime ? `${duration}` : '-'}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function StatCard({ title, value, icon, isDanger }: { title: string; value: string; icon: React.ReactNode; isDanger?: boolean }) {
  return (
    <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex items-center justify-between">
      <div>
        <p className="text-sm font-medium text-slate-500 mb-1">{title}</p>
        <p className={`text-3xl font-bold ${isDanger ? 'text-red-600' : 'text-slate-800'}`}>{value}</p>
      </div>
      <div className={`p-3 rounded-xl ${isDanger ? 'bg-red-50' : 'bg-slate-50'}`}>
        {icon}
      </div>
    </div>
  );
}

function ConflictSummaryCard({ type, count, icon, description, color }: { type: string, count: number, icon: React.ReactNode, description: string, color: 'red' | 'orange' | 'amber' | 'purple' | 'blue' }) {
  const colorMap = {
    red: 'bg-red-50 text-red-600 border-red-200',
    orange: 'bg-orange-50 text-orange-600 border-orange-200',
    amber: 'bg-amber-50 text-amber-600 border-amber-200',
    purple: 'bg-purple-50 text-purple-600 border-purple-200',
    blue: 'bg-blue-50 text-blue-600 border-blue-200',
  };

  return (
    <div className={`p-5 rounded-xl border ${colorMap[color]} flex flex-col`}>
      <div className="flex items-center gap-3 mb-3">
        <div className="p-2 bg-white rounded-lg shadow-sm">
          {icon}
        </div>
        <h4 className="font-bold text-lg">{type}</h4>
      </div>
      <p className="text-sm opacity-90 mb-4">{description}</p>
      <div className="mt-auto text-3xl font-black">
        {count} <span className="text-sm font-medium opacity-80">cảnh báo</span>
      </div>
    </div>
  );
}

const ActivityIcon = (props: React.ComponentProps<'svg'>) => (
  <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...props}>
    <path d="M22 12h-2.48a2 2 0 0 0-1.93 1.46l-2.35 8.36a.25.25 0 0 1-.48 0L9.24 2.18a.25.25 0 0 0-.48 0l-2.35 8.36A2 2 0 0 1 4.48 12H2" />
  </svg>
);

// Simple markdown formatter to bold text surrounded by **
function MarkdownText({ text }: { text: string }) {
  const parts = text.split(/(\*\*.*?\*\*)/g);
  return (
    <>
      {parts.map((part, i) => {
        if (part.startsWith('**') && part.endsWith('**')) {
          return <span key={i} className="font-bold">{part.slice(2, -2)}</span>;
        }
        return part;
      })}
    </>
  );
}
