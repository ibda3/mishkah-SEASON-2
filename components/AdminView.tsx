// ... داخل AdminView ...

// هذا الجزء هو المسؤول عن رسم المربعات في الأسفل
<div className="grid grid-cols-4 md:grid-cols-8 gap-4">
    {(status as any).playerAnswers?.map((ans: any, i: number) => {
        // تنظيف الإجابة (إزالة المسافات وتحويلها لنص)
        const cleanAns = String(ans || "").trim();
        const hasAnswer = cleanAns !== "";
        
        // تحديد حالة الإجابة (هل هي صحيحة أم خاطئة؟)
        // ملاحظة: لا نظهر الصح/الخطأ إلا إذا ضغط الأدمن على "إظهار الإجابة" (showAnswer)
        const isRevealMode = (status as any).showAnswer;
        const correctAnswer = String((status as any).correctAnswer || "").trim();
        
        const isCorrect = isRevealMode && hasAnswer && cleanAns === correctAnswer;
        const isWrong = isRevealMode && hasAnswer && cleanAns !== correctAnswer;

        // تحديد ستايل المربع
        let boxClass = "bg-slate-900 border-slate-800 text-slate-700"; // افتراضي (لا إجابة)
        
        if (!isRevealMode && hasAnswer) {
            // حالة: المتسابق أجاب، لكن لم نكشف النتيجة بعد (لون أزرق غامق للتوثيق)
            boxClass = "bg-indigo-900/50 border-indigo-500/50 text-indigo-400 shadow-lg shadow-indigo-500/10";
        } else if (isCorrect) {
            // حالة: إجابة صحيحة (أخضر ساطع)
            boxClass = "bg-emerald-500 border-emerald-400 text-white scale-110 shadow-xl shadow-emerald-500/40 ring-4 ring-emerald-500/20 z-10";
        } else if (isWrong) {
            // حالة: إجابة خاطئة (رمادي باهت)
            boxClass = "bg-slate-700 border-slate-600 text-slate-400 opacity-60 grayscale";
        }

        return (
            <div key={i} className="flex flex-col items-center gap-3 group">
                <div className={`w-16 h-16 rounded-2xl flex items-center justify-center text-3xl font-black border-2 transition-all duration-500 ${boxClass}`}>
                    {/* نظهر الحرف دائماً إذا كان موجوداً */}
                    {hasAnswer ? cleanAns : '-'}
                </div>
                <div className="flex flex-col items-center">
                    <span className={`text-[10px] font-black mb-1 transition-colors ${hasAnswer ? 'text-indigo-400' : 'text-slate-600'}`}>P{i+1}</span>
                    {hasAnswer && !isRevealMode && <span className="w-1.5 h-1.5 bg-indigo-500 rounded-full animate-ping"></span>}
                </div>
            </div>
        );
    })}
</div>