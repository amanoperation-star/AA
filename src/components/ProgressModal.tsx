import React from 'react';

interface ProgressModalProps {
  isOpen: boolean;
  onClose?: () => void;
  title?: string;
  fileName?: string;
  subtitle?: string;
  percent: number;
  stepText?: string;
  countText?: string;
  isComplete?: boolean;
}

export const ProgressModal: React.FC<ProgressModalProps> = ({
  isOpen,
  onClose,
  title = 'قراءة',
  fileName = 'data_report.xlsx',
  subtitle = 'جاري قراءة ومعالجة البيانات من الملف المرفق، يرجى الانتظار حتى اكتمال المعالجة...',
  percent = 25,
  stepText = 'جاري استخراج السجلات وإعداد الجداول...',
  countText,
  isComplete = false,
}) => {
  if (!isOpen) return null;

  const roundedPercent = Math.min(100, Math.max(0, Math.round(percent)));

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-[#070a12]/80 backdrop-blur-md p-4 font-['Cairo'] text-right select-none transition-all duration-300">
      {/* Background Subtle Animated Grid */}
      <div className="absolute inset-0 modal-bg-grid pointer-events-none z-0 opacity-80" />

      {/* Radial ambient glow gradients */}
      <div className="absolute w-96 h-96 bg-blue-600/15 rounded-full blur-3xl pointer-events-none -top-10 -right-10" />
      <div className="absolute w-96 h-96 bg-indigo-600/15 rounded-full blur-3xl pointer-events-none -bottom-10 -left-10" />

      {/* Main Modal Container matching exact HTML / CSS */}
      <div className="relative z-10 w-full max-w-[640px] bg-[#0f172a]/85 backdrop-blur-xl border border-white/[0.08] rounded-[24px] p-7 sm:p-8 shadow-[0_25px_50px_-12px_rgba(0,0,0,0.6),0_0_40px_0_rgba(59,130,246,0.12),inset_0_1px_1px_0_rgba(255,255,255,0.1)] overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        
        {/* Top Header Navigation */}
        <div className="flex items-center justify-between gap-4 mb-5">
          
          {/* Right Info & File (In RTL this is on the right) */}
          <div className="flex flex-col gap-1 order-1">
            <div className="flex items-center gap-2.5 flex-wrap">
              <h2 className="text-2xl font-extrabold text-[#f8fafc] leading-tight">
                {title}
              </h2>
              <div className="inline-flex items-center gap-1.5 font-['Fira_Code',monospace] text-[13px] text-[#94a3b8] bg-white/[0.04] border border-white/[0.08] px-2.5 py-1 rounded-lg" dir="ltr">
                <i className="fa-solid fa-file-excel text-[#38bdf8] text-xs"></i>
                <span className="truncate max-w-[180px] sm:max-w-[240px]" title={fileName}>
                  {fileName || 'data_report.xlsx'}
                </span>
              </div>
            </div>
          </div>

          {/* Center Gear Animation */}
          <div className="relative w-[52px] h-[52px] rounded-2xl bg-gradient-to-br from-[#3b82f6]/20 to-[#6366f1]/10 border border-[#3b82f6]/35 flex items-center justify-center shadow-[0_0_20px_rgba(59,130,246,0.2)] shrink-0 order-2">
            <i className={`fa-solid fa-gear text-[22px] text-[#38bdf8] ${!isComplete ? 'animate-spin' : ''}`}></i>
          </div>

          {/* Left Controls: Close Button & Top Percentage (In RTL this is on the left) */}
          <div className="flex items-center gap-2.5 order-3 shrink-0">
            <span className="font-['Fira_Code',monospace] text-[13px] font-semibold text-[#38bdf8] bg-[#38bdf8]/10 border border-[#38bdf8]/25 px-3 py-1.5 rounded-full min-w-[52px] text-center" dir="ltr">
              {roundedPercent}%
            </span>
            {onClose && (
              <button
                type="button"
                onClick={onClose}
                className="inline-flex items-center gap-1.5 bg-[#e11d48]/15 hover:bg-[#e11d48]/30 border border-[#e11d48]/35 text-[#fda4af] hover:text-white px-3.5 sm:px-4 py-1.5 rounded-xl text-[13px] font-bold cursor-pointer transition-all duration-200 active:scale-95 shadow-sm"
              >
                <i className="fa-solid fa-xmark text-xs"></i>
                <span>{isComplete ? 'إغلاق' : 'إلغاء الأمر'}</span>
              </button>
            )}
          </div>
        </div>

        {/* Description Paragraph */}
        <p className="text-[#94a3b8] text-sm leading-relaxed mb-6 font-normal">
          {subtitle}
        </p>

        {/* Progress Bar Section */}
        <div className="mb-5">
          <div className="relative h-2.5 w-full bg-[#0f172a]/90 rounded-full border border-white/[0.05] shadow-[inset_0_2px_4px_rgba(0,0,0,0.6)] overflow-hidden p-[2px]">
            <div
              className="h-full rounded-full bg-gradient-to-r from-[#6366f1] via-[#3b82f6] to-[#f59e0b] shadow-[0_0_12px_rgba(56,189,248,0.5)] progress-bar-animated-fill transition-all duration-300 ease-out"
              style={{ width: `${roundedPercent}%` }}
            />
          </div>
        </div>

        {/* Bottom Footer Status */}
        <div className="flex items-center justify-between gap-3 text-[13px]">
          <div className="text-[#94a3b8] flex items-center gap-2 font-medium">
            <span className={`w-2 h-2 rounded-full ${isComplete ? 'bg-emerald-400 shadow-[0_0_8px_#10b981]' : 'bg-[#38bdf8] shadow-[0_0_8px_#38bdf8] status-dot-pulse'}`} />
            <span className="truncate max-w-[280px] sm:max-w-[420px]">
              {stepText || 'جاري استخراج السجلات وإعداد الجداول...'}
            </span>
            {countText && (
              <span className="text-[11px] text-[#38bdf8] font-mono font-bold bg-[#38bdf8]/10 px-2 py-0.5 rounded-md">
                {countText}
              </span>
            )}
          </div>

          <span className="font-['Fira_Code',monospace] font-semibold text-[#6366f1]" dir="ltr">
            {roundedPercent}%
          </span>
        </div>

        {/* Completed notification pill if isComplete */}
        {isComplete && (
          <div className="mt-4 p-2.5 rounded-xl bg-emerald-500/10 border border-emerald-500/30 text-emerald-400 text-xs font-bold text-center flex items-center justify-center gap-2 animate-in fade-in">
            <i className="fa-solid fa-circle-check text-emerald-400"></i>
            <span>اكتملت المعالجة بنجاح! يتم الآن تحديث البيانات...</span>
          </div>
        )}
      </div>
    </div>
  );
};
