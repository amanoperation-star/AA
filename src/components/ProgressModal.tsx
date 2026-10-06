import React from 'react';
import { Settings, CheckCircle2, FileSpreadsheet, X } from 'lucide-react';

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
  fileName = 'نموذج_شيت_ربط_المناديب_والحسابات.csv',
  subtitle = 'جاري تحليل خلايا وسجلات ملف الإكسل...',
  percent = 60,
  stepText = 'تحويل ورقة العمل إلى هيكل بيانات JSON...',
  countText,
  isComplete = false,
}) => {
  if (!isOpen) return null;

  const roundedPercent = Math.min(100, Math.max(0, Math.round(percent)));

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-[#030712]/85 backdrop-blur-md transition-all duration-300 p-4 font-['Cairo']">
      {/* Exact Floating Card matching image.png */}
      <div className="bg-[#0c1326] border border-slate-800/90 rounded-3xl p-6 sm:p-7 max-w-xl w-full mx-auto shadow-[0_25px_60px_-15px_rgba(0,0,0,0.9)] relative overflow-hidden text-right select-none animate-in fade-in zoom-in-95 duration-200">
        
        {/* Subtle Ambient Glow */}
        <div className="absolute top-0 right-1/4 w-48 h-48 bg-blue-600/10 rounded-full blur-3xl pointer-events-none" />
        <div className="absolute bottom-0 left-1/4 w-48 h-48 bg-indigo-600/10 rounded-full blur-3xl pointer-events-none" />

        {/* 1. TOP HEADER SECTION */}
        <div className="flex items-start justify-between gap-4 mb-4">
          {/* Top Left in RTL: Percentage Capsule Badge + Close Button */}
          <div className="order-2 sm:order-1 shrink-0 flex items-center gap-2">
            <span className="px-3.5 py-1 rounded-full text-xs font-mono font-black bg-[#0d1c3a] text-blue-400 border border-blue-500/40 shadow-inner inline-block">
              {roundedPercent}%
            </span>
            {onClose && (
              <button
                onClick={onClose}
                className="w-8 h-8 rounded-xl bg-slate-800/80 hover:bg-slate-700 text-slate-400 hover:text-white flex items-center justify-center transition border border-slate-700/60 shrink-0 cursor-pointer shadow-sm"
                title="إغلاق النافذة"
              >
                <X className="w-4 h-4" />
              </button>
            )}
          </div>

          {/* Top Right in RTL: Spinning Gear Icon + Title + File Name + Subtitle */}
          <div className="flex items-start gap-3.5 order-1 sm:order-2 flex-1">
            {/* Spinning Gear Icon in Circular Dark Blue Badge */}
            <div className="w-12 h-12 rounded-2xl bg-blue-600/20 text-blue-400 border border-blue-500/30 flex items-center justify-center shrink-0 shadow-sm">
              <Settings className="w-6 h-6 animate-spin text-blue-400" />
            </div>

            {/* Texts */}
            <div className="space-y-0.5">
              <h3 className="text-base sm:text-lg font-black text-white leading-tight">
                {title}
              </h3>
              <p className="text-sm font-bold text-slate-100 font-mono tracking-wide break-all" dir="ltr">
                {fileName}
              </p>
              <p className="text-xs text-slate-400 font-medium pt-0.5">
                {subtitle}
              </p>
            </div>
          </div>
        </div>

        {/* 2. PROGRESS BAR TRACK */}
        <div className="my-4">
          <div className="w-full h-3 bg-[#080d1a] rounded-full overflow-hidden p-0.5 border border-slate-700/60 shadow-inner">
            <div
              className="h-full bg-gradient-to-r from-indigo-500 via-blue-400 to-amber-400 rounded-full transition-all duration-200 ease-out shadow-[0_0_12px_rgba(59,130,246,0.6)]"
              style={{ width: `${roundedPercent}%` }}
            />
          </div>
        </div>

        {/* 3. BOTTOM FOOTER SECTION */}
        <div className="flex items-center justify-between text-xs pt-1">
          {/* Left in RTL: Percentage number in purple/blue font */}
          <span className="font-mono font-bold text-indigo-400 order-2 sm:order-1 text-xs">
            {roundedPercent}%
          </span>

          {/* Right in RTL: Step details / status description */}
          <span className="text-slate-400 font-medium order-1 sm:order-2 text-xs truncate max-w-[80%]">
            {stepText || 'تحويل ورقة العمل إلى هيكل بيانات JSON...'}
          </span>
        </div>

        {/* Complete State Badge */}
        {isComplete && (
          <div className="mt-4 p-2.5 rounded-xl bg-emerald-500/10 border border-emerald-500/30 text-emerald-400 text-xs font-bold text-center flex items-center justify-center gap-2 animate-in fade-in">
            <CheckCircle2 className="w-4 h-4 text-emerald-400" />
            <span>اكتملت المعالجة بنجاح! يتم تحديث الجداول...</span>
          </div>
        )}
      </div>
    </div>
  );
};
