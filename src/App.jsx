import React, { useState, useEffect, useMemo } from 'react';
import { 
  Users, Activity, PlusCircle, FileSpreadsheet, Save, Trash2, Home, Upload, Eye, X, Lock, Unlock, AlertTriangle, BarChart2, Cloud as CloudIcon
} from 'lucide-react';
import { initializeApp } from 'firebase/app';
import { getAuth, signInAnonymously, onAuthStateChanged, signInWithCustomToken } from 'firebase/auth';
import { getFirestore, collection, doc, setDoc, deleteDoc, onSnapshot } from 'firebase/firestore';

let app, auth, db, appId;
try {
  const firebaseConfig = typeof __firebase_config !== 'undefined' ? JSON.parse(__firebase_config) : null;
  if (firebaseConfig && Object.keys(firebaseConfig).length > 0) {
    app = initializeApp(firebaseConfig);
    auth = getAuth(app);
    db = getFirestore(app);
    appId = typeof __app_id !== 'undefined' ? __app_id : 'default-app-id';
  }
} catch (e) {
  console.error("Firebase init failed, running locally:", e);
}

const calcMean = (arr) => {
  const valid = arr.filter(n => n !== null && n !== '' && !isNaN(n));
  if (valid.length === 0) return null;
  const sum = valid.reduce((a, b) => a + Number(b), 0);
  return (sum / valid.length).toFixed(2);
};

const calcSum = (arr) => {
  const valid = arr.filter(n => n !== null && n !== '' && !isNaN(n));
  if (valid.length === 0) return null;
  return valid.reduce((a, b) => a + Number(b), 0).toFixed(2);
};

const calcSD = (arr) => {
  const valid = arr.filter(n => n !== null && n !== '' && !isNaN(n));
  if (valid.length < 2) return null;
  const mean = valid.reduce((a, b) => a + Number(b), 0) / valid.length;
  const variance = valid.reduce((a, b) => a + Math.pow(Number(b) - mean, 2), 0) / (valid.length - 1);
  return Math.sqrt(variance).toFixed(2);
};

const calcTotalROM = (irMean, erMean) => {
  if (irMean && erMean) return (Number(irMean) + Number(erMean)).toFixed(2);
  return null;
};

const calculateANOVA = (groups) => {
  const validGroups = groups.filter(g => g.length > 0);
  if (validGroups.length < 2) return null;

  let totalSum = 0;
  let totalCount = 0;
  validGroups.forEach(g => {
    totalSum += g.reduce((a,b)=>a+b, 0);
    totalCount += g.length;
  });
  const overallMean = totalSum / totalCount;

  let ssb = 0; 
  let ssw = 0; 

  validGroups.forEach(g => {
    const groupMean = g.reduce((a,b)=>a+b, 0) / g.length;
    ssb += g.length * Math.pow(groupMean - overallMean, 2);
    g.forEach(val => {
      ssw += Math.pow(val - groupMean, 2);
    });
  });

  const dfB = validGroups.length - 1;
  const dfW = totalCount - validGroups.length;

  if (dfW === 0 || dfB === 0 || ssw === 0) return null;

  const msb = ssb / dfB;
  const msw = ssw / dfW;
  const fScore = msb / msw;

  if (window.jStat) {
    try {
      return 1 - window.jStat.centralF.cdf(fScore, dfB, dfW);
    } catch(e) { return null; }
  }
  return null;
};

const kpullMetricsList = [
  "Max Value (kg)", "Average Value (kg)", "RFD To Max (kg/s)", "RFD From Max (kg/s)",
  "RFD 20% - 80% Fmax (kg/s)", "Average RFD (kg/s)", "Time To Max (sec)",
  "Max Torque (kg•m)", "Average Torque (kg•m)", "Segment Length (cm)",
  "Starting Threshold (kg)", "RFD 50-100ms (kg/s)", "RFD 100-150ms (kg/s)",
  "RFD 150-200ms (kg/s)", "RFD 0-50ms (kg/s)", "RFD 0-100ms (kg/s)",
  "RFD 0-150ms (kg/s)", "RFD 0-200ms (kg/s)", "Impulse 0-50ms (kg•s)",
  "Impulse 50-100ms (kg•s)", "Impulse 100-150ms (kg•s)", "Impulse 150-200ms (kg•s)",
  "Impulse 0-100ms (kg•s)", "Impulse 0-150ms (kg•s)", "Impulse 0-200ms (kg•s)",
  "Force at 50ms (kg)", "Force at 100ms (kg)", "Force at 250ms (kg)"
];

const initialRomState = () => ({
  activeRightIR: ['', '', ''], passiveRightIR: ['', '', ''],
  activeRightER: ['', '', ''], passiveRightER: ['', '', ''],
  activeLeftIR: ['', '', ''], passiveLeftIR: ['', '', ''],
  activeLeftER: ['', '', ''], passiveLeftER: ['', '', ''],
  thoracicLeft: ['', '', ''], thoracicRight: ['', '', '']
});

const initialYbtState = () => ({
  medial: ['', '', ''], il: ['', '', ''], sl: ['', '', '']
});

const initialKPullState = () => {
  const state = {};
  kpullMetricsList.forEach(metric => {
    state[`ER_${metric}`] = ['', '', ''];
    state[`IR_${metric}`] = ['', '', ''];
  });
  return state;
};

const emptySubject = {
  id: '', subjectNo: '', name: '', gender: '男性', age: '', height: '', weight: '', 
  dominantHand: '右', tennisAge: '', trainingHours: '', ntrp: '', armLength: '',
  group: '健康組', painDuration: '', painVas: '', notes: '',
  rom: initialRomState(),
  smrt: {
    result: '陰性',
    checklist: {
      scapularAnterior: false, dyspnea: false, difficulty: false,
      noReach60: false, humeralAnterior: false, fatigue: false, feedbackNeeded: false, helpNeeded: false
    }
  },
  ybt: initialYbtState(),
  setDuration: '',
  kpull: initialKPullState()
};

const getQuestions = (subject) => {
  const q = [
    { id: 'subjectNo', domain: 'basic', type: 'text', label: '受試者編號', placeholder: '' },
    { id: 'name', domain: 'basic', type: 'text', label: '姓名', placeholder: '' },
    { id: 'gender', domain: 'basic', type: 'select', label: '性別', options: ['男性', '女性'] },
    { id: 'age', domain: 'basic', type: 'number', label: '年齡 (yr)', placeholder: '' },
    { id: 'height', domain: 'basic', type: 'number', label: '身高 (cm)', step: '0.1', placeholder: '' },
    { id: 'weight', domain: 'basic', type: 'number', label: '體重 (kg)', step: '0.1', placeholder: '' },
    { id: 'armLength', domain: 'basic', type: 'number', label: '上肢長度 (cm)', step: '0.1', placeholder: '' },
    { id: 'dominantHand', domain: 'basic', type: 'select', label: '慣用側', options: ['右', '左'] },
    { id: 'tennisAge', domain: 'basic', type: 'number', label: '球齡 (yr)', step: '0.1', placeholder: '' },
    { id: 'trainingHours', domain: 'basic', type: 'number', label: '每週訓練量 (hr/wk)', step: '0.1', placeholder: '' },
    { id: 'ntrp', domain: 'basic', type: 'select', label: 'NTRP 分級', options: ['', 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6, 6.5, 7] },
    { id: 'group', domain: 'basic', type: 'select', label: '組別設定', options: ['健康組', '肩峰下疼痛組', '肩痛病史組'] },
    { id: 'notes', domain: 'basic', type: 'textarea', label: '備註', placeholder: '' },
  ];

  if (subject.group === '肩峰下疼痛組') {
    q.push({ id: 'painDuration', domain: 'basic', type: 'text', label: '疼痛時長', placeholder: '' });
    q.push({ id: 'painVas', domain: 'basic', type: 'number', label: '最近一週最大疼痛程度 (VAS 1-10)', min: 1, max: 10, placeholder: '' });
  }

  q.push(
    { id: 'activeRightIR', domain: 'rom', type: 'trials', label: '主動肩關節: 右側內轉 (R-IR)', path: 'rom' },
    { id: 'passiveRightIR', domain: 'rom', type: 'trials', label: '被動肩關節: 右側內轉 (R-IR)', path: 'rom' },
    { id: 'activeRightER', domain: 'rom', type: 'trials', label: '主動肩關節: 右側外轉 (R-ER)', path: 'rom' },
    { id: 'passiveRightER', domain: 'rom', type: 'trials', label: '被動肩關節: 右側外轉 (R-ER)', path: 'rom' },
    
    { id: 'activeLeftIR', domain: 'rom', type: 'trials', label: '主動肩關節: 左側內轉 (L-IR)', path: 'rom' },
    { id: 'passiveLeftIR', domain: 'rom', type: 'trials', label: '被動肩關節: 左側內轉 (L-IR)', path: 'rom' },
    { id: 'activeLeftER', domain: 'rom', type: 'trials', label: '主動肩關節: 左側外轉 (L-ER)', path: 'rom' },
    { id: 'passiveLeftER', domain: 'rom', type: 'trials', label: '被動肩關節: 左側外轉 (L-ER)', path: 'rom' },
    
    { id: 'thoracicRight', domain: 'rom', type: 'trials', label: '胸椎: 向右旋轉', path: 'rom' },
    { id: 'thoracicLeft', domain: 'rom', type: 'trials', label: '胸椎: 向左旋轉', path: 'rom' }
  );

  q.push(
    { id: 'smrt', domain: 'smrt', type: 'smrt', label: '肩膀內轉動作控制 (SMRT)' }
  );

  q.push(
    { id: 'medial', domain: 'functional', type: 'trials', label: 'YBT-UQ: Medial', path: 'ybt' },
    { id: 'sl', domain: 'functional', type: 'trials', label: 'YBT-UQ: Superolateral (SL)', path: 'ybt' },
    { id: 'il', domain: 'functional', type: 'trials', label: 'YBT-UQ: Inferolateral (IL)', path: 'ybt' },
    { id: 'setDuration', domain: 'functional', type: 'number', label: '後側肩膀肌肉耐力測試 (SET) 持續秒數', step: '0.1' }
  );

  q.push(
    { id: 'kpull_upload', domain: 'kpull', type: 'kpull_upload', label: 'K-Pull 數據上傳' }
  );

  return q;
};

const domains = [
  { id: 'basic', label: '👤 基本資料' },
  { id: 'rom', label: '📐 ROM' },
  { id: 'smrt', label: '💪 SMRT' },
  { id: 'functional', label: '⚖️ 功能測試' },
  { id: 'kpull', label: '📈 K-Pull' }
];

const Dashboard = ({ subjects, setView, handleStartNew }) => {
  const total = subjects.length;
  const groupCount = subjects.reduce((acc, s) => {
    acc[s.group] = (acc[s.group] || 0) + 1;
    return acc;
  }, {});

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <h1 className="text-2xl font-bold text-gray-800 flex items-center gap-2">
        <Activity className="text-blue-600" /> 臨床實驗數據管理
      </h1>
      
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 flex flex-col items-center justify-center">
          <p className="text-gray-500 text-sm font-medium">已收集總人數</p>
          <p className="text-4xl font-bold text-blue-600 mt-2">{total}</p>
        </div>
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 col-span-1 md:col-span-2">
          <p className="text-gray-500 text-sm font-medium mb-3">組別分佈</p>
          <div className="flex gap-4 flex-wrap">
            {['健康組', '肩峰下疼痛組', '肩痛病史組'].map(g => (
              <div key={g} className="bg-blue-50 px-4 py-2 rounded-lg flex-1 min-w-[120px] text-center">
                <span className="block text-xs text-blue-600 mb-1">{g}</span>
                <span className="block text-xl font-bold text-gray-800">{groupCount[g] || 0}</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="flex flex-col sm:flex-row gap-4 mt-8">
        <button 
          onClick={handleStartNew}
          className="flex-1 bg-blue-600 hover:bg-blue-700 text-white p-4 rounded-xl font-medium flex items-center justify-center gap-2 transition-colors"
        >
          <PlusCircle size={20} /> 新增受試者
        </button>
        <button 
          onClick={() => setView('list')}
          className="flex-1 bg-white hover:bg-gray-50 text-gray-700 border border-gray-200 p-4 rounded-xl font-medium flex items-center justify-center gap-2 transition-colors"
        >
          <Users size={20} /> 個案資料管理
        </button>
        <button 
          onClick={() => setView('analysis')}
          disabled={total < 2}
          className={`flex-1 p-4 rounded-xl font-medium flex items-center justify-center gap-2 transition-colors ${total < 2 ? 'bg-gray-100 text-gray-400 cursor-not-allowed' : 'bg-indigo-600 hover:bg-indigo-700 text-white'}`}
        >
          <BarChart2 size={20} /> 組間顯著性分析
        </button>
      </div>
    </div>
  );
};

const AnalysisView = ({ subjects, setView }) => {
  const [analyzedData, setAnalyzedData] = useState([]);

  useEffect(() => {
    const extractors = [
      { name: '年齡 (yr)', get: s => s.age },
      { name: '身高 (cm)', get: s => s.height },
      { name: '體重 (kg)', get: s => s.weight },
      { name: 'ROM 主動右側總和', get: s => calcTotalROM(calcMean(s.rom.activeRightIR), calcMean(s.rom.activeRightER)) },
      { name: 'ROM 主動左側總和', get: s => calcTotalROM(calcMean(s.rom.activeLeftIR), calcMean(s.rom.activeLeftER)) },
      { name: 'YBT 綜合標準化分數', get: s => {
          if (!s.armLength) return null;
          const m = calcMean(s.ybt.medial), il = calcMean(s.ybt.il), sl = calcMean(s.ybt.sl);
          if(m && il && sl) return ((Number(m) + Number(il) + Number(sl)) / (3 * Number(s.armLength))) * 100;
          return null;
        } 
      },
      { name: 'K-Pull ER Max 標準化 %BW', get: s => {
          const mean = calcMean(s.kpull['ER_Max Value (kg)']);
          return (mean && s.weight) ? (Number(mean) / Number(s.weight)) * 100 : null;
        } 
      },
      { name: 'K-Pull IR Max 標準化 %BW', get: s => {
          const mean = calcMean(s.kpull['IR_Max Value (kg)']);
          return (mean && s.weight) ? (Number(mean) / Number(s.weight)) * 100 : null;
        } 
      }
    ];

    const results = extractors.map(ext => {
      const gHealth = [], gPain = [], gHistory = [];
      subjects.forEach(sub => {
        const val = ext.get(sub);
        if (val !== null && val !== '' && !isNaN(val)) {
          if (sub.group === '健康組') gHealth.push(Number(val));
          else if (sub.group === '肩峰下疼痛組') gPain.push(Number(val));
          else if (sub.group === '肩痛病史組') gHistory.push(Number(val));
        }
      });

      const meanSdStr = (arr) => arr.length > 0 ? `${calcMean(arr)} ± ${calcSD(arr) || '0'}` : '-';
      const pValue = calculateANOVA([gHealth, gPain, gHistory]);

      return {
        name: ext.name,
        healthStr: meanSdStr(gHealth),
        painStr: meanSdStr(gPain),
        historyStr: meanSdStr(gHistory),
        pValue: pValue !== null ? pValue : null
      };
    });

    setAnalyzedData(results);
  }, [subjects]);

  return (
    <div className="space-y-6 animate-in fade-in">
      <div className="flex justify-between items-center gap-4">
        <h2 className="text-2xl font-bold text-gray-800 flex items-center gap-2">
          <BarChart2 className="text-indigo-600" /> 組間顯著性分析 (One-Way ANOVA)
        </h2>
        <button onClick={() => setView('dashboard')} className="p-2 border rounded-lg text-gray-600 hover:bg-gray-50">
          <Home size={20} />
        </button>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-x-auto">
        <table className="w-full text-left border-collapse min-w-[600px]">
          <thead>
            <tr className="bg-gray-50 text-gray-600 text-sm border-b">
              <th className="p-4">比較參數</th>
              <th className="p-4">健康組 (Mean ± SD)</th>
              <th className="p-4">肩峰下疼痛組</th>
              <th className="p-4">肩痛病史組</th>
              <th className="p-4 font-bold text-indigo-700">P-value</th>
            </tr>
          </thead>
          <tbody>
            {analyzedData.map((row, idx) => {
              const isSignificant = row.pValue !== null && row.pValue < 0.05;
              return (
                <tr key={idx} className={`border-b last:border-0 hover:bg-gray-50 transition ${isSignificant ? 'bg-red-50/50' : ''}`}>
                  <td className="p-4 font-medium text-gray-900">{row.name}</td>
                  <td className="p-4 text-sm text-gray-700">{row.healthStr}</td>
                  <td className="p-4 text-sm text-gray-700">{row.painStr}</td>
                  <td className="p-4 text-sm text-gray-700">{row.historyStr}</td>
                  <td className={`p-4 font-bold ${isSignificant ? 'text-red-600' : 'text-gray-500'}`}>
                    {row.pValue !== null ? row.pValue.toFixed(4) : '-'}
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
      <p className="text-sm text-gray-500">※ 採用 One-Way ANOVA 檢定，紅色標示代表該參數在三組間具備統計上的顯著差異 ($p &lt; 0.05$)。</p>
    </div>
  );
};

const SubjectList = ({ subjects, setView, handleEdit, handleDelete, exportToExcel, isExporting }) => {
  const [previewData, setPreviewData] = useState(null);
  const [deleteId, setDeleteId] = useState(null);

  const sortedSubjects = [...subjects].sort((a, b) => (a.subjectNo || '').localeCompare(b.subjectNo || '', undefined, { numeric: true }));

  const confirmDelete = () => {
    handleDelete(deleteId);
    setDeleteId(null);
  };

  const PreviewModal = ({ subject, onClose }) => {
    const kpull = subject.kpull;

    const renderRomRow = (label, irArr, erArr) => (
      <div className="bg-white p-2.5 rounded-lg border border-gray-100 shadow-sm">
        <div className="font-bold text-gray-700 mb-2">{label}</div>
        <div className="flex justify-between text-xs text-gray-600">
          <span>IR Mean: <strong className="text-blue-600 text-sm">{calcMean(irArr) || '-'}</strong></span>
          <span>ER Mean: <strong className="text-blue-600 text-sm">{calcMean(erArr) || '-'}</strong></span>
          <span>Total: <strong className="text-green-600 text-sm">{calcTotalROM(calcMean(irArr), calcMean(erArr)) || '-'}</strong></span>
        </div>
      </div>
    );

    return (
      <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-50 p-4 animate-in fade-in">
        <div className="bg-white rounded-2xl shadow-xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden animate-in zoom-in-95 duration-200">
          <div className="px-6 py-4 border-b flex justify-between items-center bg-gray-50 shrink-0">
            <h3 className="text-xl font-bold text-gray-800 flex items-center gap-2">
              <Eye className="text-blue-600"/> 資料預覽 - {subject.name}
            </h3>
            <button onClick={onClose} className="p-2 text-gray-400 hover:bg-gray-200 hover:text-gray-700 rounded-full transition-colors">
              <X size={20}/>
            </button>
          </div>
          
          <div className="p-6 overflow-y-auto space-y-6 flex-1">
            <section>
              <h4 className="font-bold text-gray-500 border-b pb-1 mb-3 text-sm uppercase tracking-wider">基本資料</h4>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-y-3 gap-x-4 text-sm">
                <div><span className="text-gray-400 block text-xs">編號</span> <strong className="text-gray-800">{subject.subjectNo}</strong></div>
                <div><span className="text-gray-400 block text-xs">性別</span> <strong className="text-gray-800">{subject.gender || '-'}</strong></div>
                <div className="col-span-2 sm:col-span-2"><span className="text-gray-400 block text-xs">組別</span> <strong className="text-blue-700 bg-blue-50 px-2 py-0.5 rounded">{subject.group}</strong></div>
                <div className="col-span-2 sm:col-span-4"><span className="text-gray-400 block text-xs">年齡/身高/體重</span> <strong className="text-gray-800">{subject.age}歲 / {subject.height}cm / {subject.weight}kg</strong></div>
                <div className="col-span-2 sm:col-span-4"><span className="text-gray-400 block text-xs">備註</span> <strong className="text-gray-800 whitespace-pre-wrap">{subject.notes || '-'}</strong></div>
              </div>
            </section>

            <section>
              <h4 className="font-bold text-gray-500 border-b pb-1 mb-3 text-sm uppercase tracking-wider">關節活動度 (ROM)</h4>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-sm">
                {renderRomRow('主動肩關節 (右側)', subject.rom.activeRightIR, subject.rom.activeRightER)}
                {renderRomRow('主動肩關節 (左側)', subject.rom.activeLeftIR, subject.rom.activeLeftER)}
                {renderRomRow('被動肩關節 (右側)', subject.rom.passiveRightIR, subject.rom.passiveRightER)}
                {renderRomRow('被動肩關節 (左側)', subject.rom.passiveLeftIR, subject.rom.passiveLeftER)}
                
                <div className="bg-white p-2.5 rounded-lg border border-gray-100 shadow-sm md:col-span-2">
                  <div className="font-bold text-gray-700 mb-2">胸椎旋轉 (Thoracic)</div>
                  <div className="flex justify-start gap-8 text-xs text-gray-600">
                    <span>向右 Mean: <strong className="text-blue-600 text-sm">{calcMean(subject.rom.thoracicRight) || '-'}</strong></span>
                    <span>向左 Mean: <strong className="text-blue-600 text-sm">{calcMean(subject.rom.thoracicLeft) || '-'}</strong></span>
                  </div>
                </div>
              </div>
            </section>

            <section>
              <h4 className="font-bold text-gray-500 border-b pb-1 mb-3 text-sm uppercase tracking-wider">SMRT 動作控制</h4>
              <div className="text-sm bg-gray-50 p-3 rounded-lg border border-gray-100">
                判定結果：<strong className={subject.smrt.result === '陽性' ? 'text-red-600 text-lg' : 'text-green-600 text-lg'}>{subject.smrt.result}</strong>
                <span className="text-gray-500 ml-2">({Object.values(subject.smrt.checklist).filter(Boolean).length} 項異常)</span>
              </div>
            </section>

            <section>
              <h4 className="font-bold text-gray-500 border-b pb-1 mb-3 text-sm uppercase tracking-wider">功能測試 (Functional)</h4>
              <div className="bg-white p-3 rounded-lg border border-gray-100 shadow-sm space-y-3 text-sm">
                <div>
                  <span className="font-bold text-gray-700 block mb-2">YBT-UQ 平均距離與標準化分數</span>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 text-xs text-gray-600 bg-gray-50 p-2 rounded">
                    <div className="flex flex-col">
                      <span>Medial Mean: <strong className="text-blue-600 text-sm">{calcMean(subject.ybt.medial) || '-'}</strong> cm</span>
                      <span className="text-green-600 font-medium">方向標準化: {subject.armLength && calcSum(subject.ybt.medial) ? ((calcSum(subject.ybt.medial) / subject.armLength) * 100).toFixed(2) : '-'} %</span>
                    </div>
                    <div className="flex flex-col">
                      <span>SL Mean: <strong className="text-blue-600 text-sm">{calcMean(subject.ybt.sl) || '-'}</strong> cm</span>
                      <span className="text-green-600 font-medium">方向標準化: {subject.armLength && calcSum(subject.ybt.sl) ? ((calcSum(subject.ybt.sl) / subject.armLength) * 100).toFixed(2) : '-'} %</span>
                    </div>
                    <div className="flex flex-col">
                      <span>IL Mean: <strong className="text-blue-600 text-sm">{calcMean(subject.ybt.il) || '-'}</strong> cm</span>
                      <span className="text-green-600 font-medium">方向標準化: {subject.armLength && calcSum(subject.ybt.il) ? ((calcSum(subject.ybt.il) / subject.armLength) * 100).toFixed(2) : '-'} %</span>
                    </div>
                  </div>
                  <div className="mt-2 p-2 bg-green-50 text-green-800 rounded flex justify-between items-center border border-green-100">
                    <span className="font-bold">綜合標準化總分 (Composite Score)</span>
                    <strong className="text-lg">
                      {(subject.armLength && calcMean(subject.ybt.medial) && calcMean(subject.ybt.sl) && calcMean(subject.ybt.il)) 
                        ? (((Number(calcMean(subject.ybt.medial)) + Number(calcMean(subject.ybt.sl)) + Number(calcMean(subject.ybt.il))) / (3 * subject.armLength)) * 100).toFixed(2) + ' %' 
                        : '-'}
                    </strong>
                  </div>
                </div>
                <div className="border-t border-gray-100 pt-3">
                  <span className="font-bold text-gray-700">SET (後側肩膀肌肉耐力): </span>
                  <strong className="text-blue-600 text-lg">{subject.setDuration || '-'}</strong> 秒
                </div>
              </div>
            </section>

            <section className="bg-blue-50/50 p-4 rounded-xl border border-blue-100">
              <h4 className="font-bold text-blue-800 mb-4 flex items-center gap-2">
                <Activity size={18}/> K-Pull 測試結果完整預覽 (28項指標)
              </h4>
              
              <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
                <div className="bg-white p-3 rounded-lg border border-orange-200 shadow-sm overflow-hidden flex flex-col">
                  <div className="text-orange-900 font-bold mb-2 bg-orange-50 p-2 rounded">外轉 (ER) 數據</div>
                  <div className="overflow-x-auto overflow-y-auto max-h-[35vh]">
                    <table className="w-full text-xs text-left min-w-[400px]">
                      <thead className="bg-gray-50 sticky top-0 shadow-sm">
                        <tr><th className="p-2">指標</th><th className="p-2">T1</th><th className="p-2">T2</th><th className="p-2">T3</th><th className="p-2 border-l border-gray-200">Mean</th><th className="p-2">SD</th></tr>
                      </thead>
                      <tbody>
                        {kpullMetricsList.map(m => {
                          const arr = kpull[`ER_${m}`] || ['', '', ''];
                          const mean = calcMean(arr);
                          const sd = calcSD(arr);
                          return (
                            <tr key={m} className="border-t border-gray-100 hover:bg-gray-50">
                              <td className="p-2 truncate max-w-[140px] font-medium" title={m}>{m}</td>
                              <td className="p-2">{arr[0] || '-'}</td>
                              <td className="p-2">{arr[1] || '-'}</td>
                              <td className="p-2">{arr[2] || '-'}</td>
                              <td className="p-2 font-bold text-orange-700 border-l border-gray-100">{mean || '-'}</td>
                              <td className="p-2 font-bold text-orange-500">{sd || '-'}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div className="bg-white p-3 rounded-lg border border-blue-200 shadow-sm overflow-hidden flex flex-col">
                  <div className="text-blue-900 font-bold mb-2 bg-blue-50 p-2 rounded">內轉 (IR) 數據</div>
                  <div className="overflow-x-auto overflow-y-auto max-h-[35vh]">
                    <table className="w-full text-xs text-left min-w-[400px]">
                      <thead className="bg-gray-50 sticky top-0 shadow-sm">
                        <tr><th className="p-2">指標</th><th className="p-2">T1</th><th className="p-2">T2</th><th className="p-2">T3</th><th className="p-2 border-l border-gray-200">Mean</th><th className="p-2">SD</th></tr>
                      </thead>
                      <tbody>
                        {kpullMetricsList.map(m => {
                          const arr = kpull[`IR_${m}`] || ['', '', ''];
                          const mean = calcMean(arr);
                          const sd = calcSD(arr);
                          return (
                            <tr key={m} className="border-t border-gray-100 hover:bg-gray-50">
                              <td className="p-2 truncate max-w-[140px] font-medium" title={m}>{m}</td>
                              <td className="p-2">{arr[0] || '-'}</td>
                              <td className="p-2">{arr[1] || '-'}</td>
                              <td className="p-2">{arr[2] || '-'}</td>
                              <td className="p-2 font-bold text-blue-700 border-l border-gray-100">{mean || '-'}</td>
                              <td className="p-2 font-bold text-blue-500">{sd || '-'}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            </section>
          </div>
          
          <div className="px-6 py-4 border-t bg-gray-50 text-right shrink-0">
            <button onClick={onClose} className="px-6 py-2 bg-gray-200 text-gray-800 font-bold rounded-lg hover:bg-gray-300 transition-colors">
              關閉預覽
            </button>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="space-y-6 animate-in fade-in">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
        <h2 className="text-2xl font-bold text-gray-800 flex items-center gap-2">
          <Users className="text-blue-600" /> 個案資料管理
        </h2>
        <div className="flex gap-2 w-full sm:w-auto">
          <button onClick={() => setView('dashboard')} className="p-2 border rounded-lg text-gray-600 hover:bg-gray-50">
            <Home size={20} />
          </button>
          <button 
            onClick={exportToExcel} 
            disabled={subjects.length === 0 || isExporting}
            className={`flex-1 sm:flex-none flex items-center justify-center gap-2 text-white px-4 py-2 rounded-lg text-sm font-medium transition ${subjects.length === 0 || isExporting ? 'bg-gray-400 cursor-not-allowed' : 'bg-green-600 hover:bg-green-700'}`}
          >
            <FileSpreadsheet size={16} className={isExporting ? "animate-pulse" : ""} /> 
            {isExporting ? '匯出中...' : '匯出'}
          </button>
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-x-auto">
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="bg-gray-50 text-gray-600 text-sm border-b">
              <th className="p-4">編號</th>
              <th className="p-4">姓名</th>
              <th className="p-4">性別</th>
              <th className="p-4">組別</th>
              <th className="p-4">年齡/身高/體重</th>
              <th className="p-4 text-center">操作</th>
            </tr>
          </thead>
          <tbody>
            {sortedSubjects.length === 0 && (
              <tr><td colSpan="6" className="p-8 text-center text-gray-400">尚無資料</td></tr>
            )}
            {sortedSubjects.map(sub => (
              <tr key={sub.id} className="border-b last:border-0 hover:bg-blue-50/50 transition">
                <td className="p-4 font-medium text-gray-900">{sub.subjectNo || '未填'}</td>
                <td className="p-4">{sub.name}</td>
                <td className="p-4 text-sm text-gray-700">{sub.gender || '-'}</td>
                <td className="p-4">
                  <span className={`px-2 py-1 rounded text-xs font-medium ${sub.group === '健康組' ? 'bg-green-100 text-green-700' : 'bg-orange-100 text-orange-700'}`}>
                    {sub.group}
                  </span>
                </td>
                <td className="p-4 text-sm text-gray-600">{sub.age}歲 / {sub.height}cm / {sub.weight}kg</td>
                <td className="p-4 text-center flex justify-center gap-2">
                  <button onClick={() => setPreviewData(sub)} className="text-gray-500 hover:bg-gray-100 hover:text-gray-800 p-2 rounded flex items-center gap-1 text-sm font-medium transition-colors">
                    <Eye size={16} /> 預覽
                  </button>
                  <button onClick={() => handleEdit(sub)} className="text-blue-600 hover:bg-blue-100 p-2 rounded text-sm font-medium transition-colors">
                    編輯
                  </button>
                  <button onClick={() => setDeleteId(sub.id)} className="text-red-500 hover:bg-red-50 p-2 rounded transition-colors">
                    <Trash2 size={18} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      
      {previewData && <PreviewModal subject={previewData} onClose={() => setPreviewData(null)} />}
      
      {deleteId && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-50 p-4 animate-in fade-in">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-sm overflow-hidden animate-in zoom-in-95 duration-200">
            <div className="p-6 text-center">
              <div className="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4 text-red-600">
                <AlertTriangle size={32} />
              </div>
              <h3 className="text-xl font-bold text-gray-800 mb-2">確定要刪除嗎？</h3>
              <p className="text-gray-500 text-sm">此操作無法復原，您確定要刪除這筆受試者的所有測量數據嗎？</p>
            </div>
            <div className="px-6 py-4 bg-gray-50 flex gap-3">
              <button onClick={() => setDeleteId(null)} className="flex-1 py-2 bg-white border border-gray-300 text-gray-700 font-bold rounded-xl hover:bg-gray-50 transition-colors">
                取消
              </button>
              <button onClick={confirmDelete} className="flex-1 py-2 bg-red-600 text-white font-bold rounded-xl hover:bg-red-700 shadow-sm transition-colors">
                確認刪除
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

const DomainViewer = ({ currentSubject, setCurrentSubject, handleSave, handleCancel, handleAutoSave }) => {
  const [activeTab, setActiveTab] = useState('basic');
  const questions = React.useMemo(() => getQuestions(currentSubject), [currentSubject.group]);
  const currentDomainQuestions = questions.filter(q => q.domain === activeTab);

  const isBasicInfoComplete = (currentSubject.name || '').trim() !== '' && (currentSubject.subjectNo || '').trim() !== '';

  const updateField = (id, val) => {
    setCurrentSubject(prev => {
      const next = { ...prev, [id]: val };
      if (handleAutoSave) handleAutoSave(next);
      return next;
    });
  };
  
  const updateTrial = (path, id, idx, val) => {
    setCurrentSubject(prev => {
      const newArr = [...prev[path][id]];
      newArr[idx] = val;
      const next = { ...prev, [path]: { ...prev[path], [id]: newArr } };
      if (handleAutoSave) handleAutoSave(next);
      return next;
    });
  };

  const handleSmrtCheck = (key) => {
    setCurrentSubject(prev => {
      const updatedChecklist = { ...prev.smrt.checklist, [key]: !prev.smrt.checklist[key] };
      const part1Score = updatedChecklist.scapularAnterior ? 1 : 0;
      const part2Keys = ['dyspnea', 'difficulty', 'noReach60', 'humeralAnterior', 'fatigue', 'feedbackNeeded', 'helpNeeded'];
      const part2Score = part2Keys.filter(k => updatedChecklist[k]).length;
      const newResult = (part1Score === 1 || part2Score > 3) ? '陽性' : '陰性';
      const next = { ...prev, smrt: { ...prev.smrt, checklist: updatedChecklist, result: newResult } };
      if (handleAutoSave) handleAutoSave(next);
      return next;
    });
  };

  const handleCsvUpload = async (e, direction) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const parsedData = {};
    kpullMetricsList.forEach(m => parsedData[m] = []);
    
    try {
      if (file.name.toLowerCase().endsWith('.pdf')) {
        // --- 處理 PDF 解析 ---
        if (!window.pdfjsLib) {
          await new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js';
            script.onload = () => {
              window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
              resolve();
            };
            script.onerror = reject;
            document.head.appendChild(script);
          });
        }

        const arrayBuffer = await file.arrayBuffer();
        const pdf = await window.pdfjsLib.getDocument(arrayBuffer).promise;
        let fullText = "";

        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const textContent = await page.getTextContent();
          const pageText = textContent.items.map(item => item.str).join(" ");
          fullText += pageText + "\n";
        }

        kpullMetricsList.forEach(metric => {
          const safeMetric = metric.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const regex = new RegExp(`${safeMetric}\\s*([-+]?[0-9]*\\.?[0-9]+)`, 'g');
          
          let match;
          while ((match = regex.exec(fullText)) !== null) {
              if (match[1] && parsedData[metric].length < 3) {
                  parsedData[metric].push(match[1]);
              }
          }
        });

      } else {
        // --- 處理 CSV 解析 ---
        const text = await file.text();
        const lines = text.split('\n').map(l => l.split(','));
        for (let line of lines) {
          if (!line[0]) continue;
          const key = line[0].trim();
          if (kpullMetricsList.includes(key)) {
            const vals = line.slice(1).map(v => v.trim()).filter(v => v !== '' && !isNaN(parseFloat(v)));
            parsedData[key].push(...vals);
          }
        }
      }
    } catch (err) {
      console.error("Error reading file:", err);
      alert("檔案讀取失敗，請確認檔案格式是否正確。");
    }
    
    kpullMetricsList.forEach(m => {
      parsedData[m] = parsedData[m].slice(0, 3);
      while(parsedData[m].length < 3) parsedData[m].push('');
    });

    setCurrentSubject(prev => {
      const newKpull = { ...prev.kpull };
      kpullMetricsList.forEach(m => {
        newKpull[`${direction}_${m}`] = parsedData[m];
      });
      const next = { ...prev, kpull: newKpull };
      if (handleAutoSave) handleAutoSave(next);
      return next;
    });
    
    e.target.value = null; 
  };

  const part1Checked = currentSubject.smrt.checklist.scapularAnterior ? 1 : 0;
  const part2Keys = ['dyspnea', 'difficulty', 'noReach60', 'humeralAnterior', 'fatigue', 'feedbackNeeded', 'helpNeeded'];
  const part2Count = part2Keys.filter(k => currentSubject.smrt.checklist[k]).length;

  return (
    <div className="flex flex-col flex-1 h-[calc(100dvh-73px)] md:h-auto bg-white md:rounded-2xl shadow-sm border-gray-200 overflow-hidden">
      
      {/* 頂部 Domain 橫向滑動選單 */}
      <div className="flex overflow-x-auto border-b bg-white sticky top-0 z-20 scrollbar-hide shrink-0">
        {domains.map(tab => {
          const isDisabled = tab.id !== 'basic' && !isBasicInfoComplete;
          return (
            <button
              key={tab.id}
              disabled={isDisabled}
              onClick={() => setActiveTab(tab.id)}
              className={`whitespace-nowrap px-5 py-4 font-bold border-b-4 transition-colors text-sm sm:text-base flex items-center gap-1
                ${activeTab === tab.id ? 'border-blue-600 text-blue-700 bg-blue-50/50' : ''}
                ${isDisabled ? 'opacity-40 cursor-not-allowed border-transparent text-gray-500' : ''}
                ${!isDisabled && activeTab !== tab.id ? 'border-transparent text-gray-500 hover:bg-gray-50' : ''}
              `}
            >
              {isDisabled && <Lock size={14} className="mb-0.5" />}
              {tab.label}
            </button>
          );
        })}
      </div>

      {/* 中間內容區塊 */}
      <div className="flex-1 overflow-y-auto p-4 md:p-6 bg-gray-50/30">
        <div className="max-w-xl mx-auto space-y-8 animate-in fade-in duration-300 pb-8">
          
          {currentDomainQuestions.map((q) => (
            <div key={q.id} className="bg-white p-5 rounded-2xl border border-gray-100 shadow-sm">
              <h2 className="text-lg md:text-xl font-bold text-gray-800 mb-4">{q.label}</h2>
              
              {/* 一般文字/數字/多行輸入 */}
              {(q.type === 'text' || q.type === 'number' || q.type === 'textarea') && (
                q.type === 'textarea' ? (
                  <textarea
                    placeholder={q.placeholder}
                    className="w-full text-lg p-3 border-2 rounded-xl focus:border-blue-500 outline-none transition-colors font-medium bg-gray-50/50 focus:bg-white border-gray-200 text-blue-900 resize-y"
                    value={currentSubject[q.id] || ''} 
                    onChange={(e) => updateField(q.id, e.target.value)} 
                    rows={3}
                  />
                ) : (
                  <input 
                    type={q.type} 
                    step={q.step} 
                    min={q.min} 
                    max={q.max}
                    placeholder={q.placeholder}
                    className={`w-full text-lg p-3 border-2 rounded-xl focus:border-blue-500 outline-none transition-colors font-medium bg-gray-50/50 focus:bg-white
                      ${(q.id === 'subjectNo' || q.id === 'name') && !currentSubject[q.id] ? 'border-orange-300 focus:border-orange-500 bg-orange-50/20' : 'border-gray-200 text-blue-900'}
                    `}
                    value={currentSubject[q.id] || ''} 
                    onChange={(e) => updateField(q.id, e.target.value)} 
                  />
                )
              )}

              {/* 下拉選單 */}
              {q.type === 'select' && (
                <select 
                  className="w-full text-lg p-3 border-2 border-gray-200 rounded-xl bg-gray-50/50 focus:bg-white focus:border-blue-500 outline-none text-gray-800 font-medium transition-colors"
                  value={currentSubject[q.id]} 
                  onChange={(e) => updateField(q.id, e.target.value)}
                >
                  {q.options.map(opt => <option key={opt} value={opt}>{opt || '請選擇...'}</option>)}
                </select>
              )}

              {/* 測量三次專用介面 (Trials) */}
              {q.type === 'trials' && (
                <div className="space-y-3">
                  <div className="flex gap-2 md:gap-4">
                    {[0, 1, 2].map(i => (
                      <div key={i} className="flex-1">
                        <div className="text-xs text-center text-gray-500 mb-1 font-bold">T{i+1}</div>
                        <input 
                          type="number" step="0.1" 
                          className="w-full text-center text-lg p-2 bg-gray-50 border-2 border-gray-200 rounded-lg focus:border-blue-500 focus:bg-white outline-none text-blue-900 font-bold transition-colors" 
                          placeholder="-" 
                          value={currentSubject[q.path][q.id][i]} 
                          onChange={e => updateTrial(q.path, q.id, i, e.target.value)} 
                        />
                      </div>
                    ))}
                  </div>
                  <div className="mt-3 px-3 py-2 bg-blue-50/80 rounded-lg border border-blue-100 flex justify-between text-blue-800 text-sm">
                    <span>Mean: <strong className="text-base">{calcMean(currentSubject[q.path][q.id]) ?? '-'}</strong></span>
                    <span>SD: <strong className="text-base">{calcSD(currentSubject[q.path][q.id]) ?? '-'}</strong></span>
                  </div>
                </div>
              )}

              {/* SMRT 複合介面 */}
              {q.type === 'smrt' && (
                <div className="space-y-4">
                  <div className={`p-4 border-2 rounded-xl font-bold flex items-center justify-between transition-colors ${currentSubject.smrt.result === '陽性' ? 'bg-red-50 text-red-700 border-red-200' : 'bg-green-50 text-green-700 border-green-200'}`}>
                    <span className="flex items-center gap-2 text-lg">
                      判定結果: {currentSubject.smrt.result}
                    </span>
                    <span className="text-sm font-normal bg-white/60 px-2 py-1 rounded-md shadow-sm">
                      得分: {part1Checked + part2Count} 分
                    </span>
                  </div>

                  {/* 第一部分 */}
                  <div className="bg-gray-50 p-4 rounded-xl border border-gray-200">
                    <h4 className="font-bold text-gray-700 mb-3 flex justify-between">
                      <span>第一部分 (1項，有勾即陽性)</span>
                      <span className={part1Checked ? 'text-red-600' : 'text-gray-400'}>{part1Checked} / 1</span>
                    </h4>
                    <label className="flex items-center gap-3 p-3 bg-white rounded-xl cursor-pointer shadow-sm border border-gray-100 active:bg-blue-50">
                      <input type="checkbox" className="w-5 h-5 rounded text-blue-600" checked={currentSubject.smrt.checklist.scapularAnterior} onChange={() => handleSmrtCheck('scapularAnterior')} />
                      <span className="text-gray-800 font-medium">肩胛骨過度前移</span>
                    </label>
                  </div>

                  {/* 第二部分 */}
                  <div className="bg-gray-50 p-4 rounded-xl border border-gray-200">
                    <h4 className="font-bold text-gray-700 mb-3 flex justify-between">
                      <span>第二部分 (7項，{'>'}3 分為陽性)</span>
                      <span className={part2Count > 3 ? 'text-red-600' : 'text-blue-600'}>{part2Count} / 7</span>
                    </h4>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                      {[
                        {k: 'dyspnea', l: '呼吸困難'},
                        {k: 'difficulty', l: '動作困難'},
                        {k: 'noReach60', l: '無法到達肩內轉 60∘'},
                        {k: 'humeralAnterior', l: '肱骨過度前移'},
                        {k: 'fatigue', l: '疲勞'},
                        {k: 'feedbackNeeded', l: '需要外在回饋'},
                        {k: 'helpNeeded', l: '需要外在幫助'}
                      ].map(c => (
                        <label key={c.k} className="flex items-start gap-3 p-3 bg-white rounded-xl cursor-pointer shadow-sm border border-gray-100 active:bg-blue-50">
                          <input type="checkbox" className="w-5 h-5 rounded text-blue-600 mt-0.5 shrink-0" checked={currentSubject.smrt.checklist[c.k]} onChange={() => handleSmrtCheck(c.k)} />
                          <span className="text-gray-800 font-medium text-sm md:text-base leading-tight">{c.l}</span>
                        </label>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {/* 極簡版 K-Pull 純上傳介面 */}
              {q.type === 'kpull_upload' && (
                <div className="space-y-6">
                  {/* ER 上傳 */}
                  <div className="bg-orange-50 border-2 border-dashed border-orange-300 p-6 md:p-8 rounded-2xl text-center">
                    <Upload className="mx-auto text-orange-400 mb-4" size={40} />
                    <h4 className="text-orange-800 font-bold text-xl mb-6">上傳 外轉 (ER) 數據檔案</h4>
                    <input type="file" accept=".csv, .pdf" onChange={(e) => handleCsvUpload(e, 'ER')} className="hidden" id="csv-upload-er" />
                    <label htmlFor="csv-upload-er" className="cursor-pointer bg-orange-600 text-white px-8 py-3 rounded-xl shadow-md hover:bg-orange-700 transition font-bold inline-block text-lg">
                      選擇 CSV 或 PDF 檔案
                    </label>
                    {currentSubject.kpull['ER_Max Value (kg)'].filter(v=>v !== '').length > 0 && (
                      <div className="mt-5 p-4 bg-white rounded-xl border border-orange-200 shadow-sm animate-in slide-in-from-bottom-2 text-left text-sm space-y-1">
                        <div className="text-green-600 font-bold mb-2 text-center">✓ 上傳成功</div>
                        <div className="text-gray-600 flex justify-between"><span>Max Value Mean:</span> <strong className="text-orange-700">{calcMean(currentSubject.kpull['ER_Max Value (kg)'])} kg</strong></div>
                        <div className="text-gray-600 flex justify-between"><span>Max Value SD:</span> <strong className="text-orange-700">{calcSD(currentSubject.kpull['ER_Max Value (kg)'])} kg</strong></div>
                      </div>
                    )}
                  </div>

                  {/* IR 上傳 */}
                  <div className="bg-blue-50 border-2 border-dashed border-blue-300 p-6 md:p-8 rounded-2xl text-center">
                    <Upload className="mx-auto text-blue-400 mb-4" size={40} />
                    <h4 className="text-blue-800 font-bold text-xl mb-6">上傳 內轉 (IR) 數據檔案</h4>
                    <input type="file" accept=".csv, .pdf" onChange={(e) => handleCsvUpload(e, 'IR')} className="hidden" id="csv-upload-ir" />
                    <label htmlFor="csv-upload-ir" className="cursor-pointer bg-blue-600 text-white px-8 py-3 rounded-xl shadow-md hover:bg-blue-700 transition font-bold inline-block text-lg">
                      選擇 CSV 或 PDF 檔案
                    </label>
                    {currentSubject.kpull['IR_Max Value (kg)'].filter(v=>v !== '').length > 0 && (
                      <div className="mt-5 p-4 bg-white rounded-xl border border-blue-200 shadow-sm animate-in slide-in-from-bottom-2 text-left text-sm space-y-1">
                        <div className="text-green-600 font-bold mb-2 text-center">✓ 上傳成功</div>
                        <div className="text-gray-600 flex justify-between"><span>Max Value Mean:</span> <strong className="text-blue-700">{calcMean(currentSubject.kpull['IR_Max Value (kg)'])} kg</strong></div>
                        <div className="text-gray-600 flex justify-between"><span>Max Value SD:</span> <strong className="text-blue-700">{calcSD(currentSubject.kpull['IR_Max Value (kg)'])} kg</strong></div>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>
          ))}

        </div>

        {/* 功能測試 (Functional) 的即時標準化分析區塊 */}
        {activeTab === 'functional' && (
          <div className="max-w-xl mx-auto mt-6 bg-blue-50 p-5 rounded-2xl border border-blue-200 shadow-sm mb-8 animate-in fade-in">
            <h3 className="text-lg font-bold text-blue-800 mb-4 flex items-center gap-2">
              <Activity size={20} /> YBT-UQ 即時標準化分析
            </h3>
            {currentSubject.armLength ? (
              <div className="space-y-4 text-sm font-medium">
                <div className="flex justify-between items-center bg-white p-2 rounded border border-blue-100">
                  <span className="text-gray-600">Medial 方向標準化:</span>
                  <span className="text-blue-700">{calcSum(currentSubject.ybt.medial) ? ((calcSum(currentSubject.ybt.medial) / currentSubject.armLength) * 100).toFixed(2) + '%' : '-'}</span>
                </div>
                <div className="flex justify-between items-center bg-white p-2 rounded border border-blue-100">
                  <span className="text-gray-600">SL 方向標準化:</span>
                  <span className="text-blue-700">{calcSum(currentSubject.ybt.sl) ? ((calcSum(currentSubject.ybt.sl) / currentSubject.armLength) * 100).toFixed(2) + '%' : '-'}</span>
                </div>
                <div className="flex justify-between items-center bg-white p-2 rounded border border-blue-100">
                  <span className="text-gray-600">IL 方向標準化:</span>
                  <span className="text-blue-700">{calcSum(currentSubject.ybt.il) ? ((calcSum(currentSubject.ybt.il) / currentSubject.armLength) * 100).toFixed(2) + '%' : '-'}</span>
                </div>
                <div className="flex justify-between items-center bg-blue-600 text-white p-3 rounded-lg shadow-sm">
                  <span className="font-bold">綜合標準化總分:</span>
                  <span className="font-bold text-xl">
                    {(calcMean(currentSubject.ybt.medial) && calcMean(currentSubject.ybt.sl) && calcMean(currentSubject.ybt.il))
                      ? (((Number(calcMean(currentSubject.ybt.medial)) + Number(calcMean(currentSubject.ybt.sl)) + Number(calcMean(currentSubject.ybt.il))) / (3 * currentSubject.armLength)) * 100).toFixed(2) + '%'
                      : '-'}
                  </span>
                </div>
              </div>
            ) : (
              <p className="text-orange-600 text-sm font-bold bg-orange-50 p-3 rounded-lg">
                ⚠️ 請先至「基本資料」填寫「上肢長度 (cm)」以進行標準化計算。
              </p>
            )}
          </div>
        )}

      </div>

      {/* 底部操作區 */}
      <div className="shrink-0 p-4 border-t flex gap-4 bg-white z-30 shadow-[0_-5px_15px_rgba(0,0,0,0.03)]">
        <button 
          onClick={handleCancel}
          className="flex-[1] py-3 sm:py-4 rounded-xl font-bold flex items-center justify-center gap-2 bg-white border border-gray-200 text-gray-700 hover:bg-gray-50 transition-colors"
        >
          返回列表
        </button>
        <button 
          onClick={handleSave}
          className="flex-[2] py-3 sm:py-4 rounded-xl font-bold flex items-center justify-center gap-2 bg-green-600 text-white hover:bg-green-700 shadow-md transition-transform active:scale-[0.98]"
        >
          <Save size={20}/> 完成 (已即時儲存)
        </button>
      </div>
    </div>
  );
};

export default function App() {
  const [view, setView] = useState('dashboard');
  const [subjects, setSubjects] = useState([]);
  const [currentSubject, setCurrentSubject] = useState(emptySubject);
  const [isExporting, setIsExporting] = useState(false);
  
  const [user, setUser] = useState(null);
  const [isCloudReady, setIsCloudReady] = useState(false);

  useEffect(() => {
    if (!auth) return;
    const initAuth = async () => {
      try {
        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
          await signInWithCustomToken(auth, __initial_auth_token);
        } else {
          await signInAnonymously(auth);
        }
      } catch(e) { console.error("Firebase Auth Error:", e); }
    };
    initAuth();
    const unsubscribe = onAuthStateChanged(auth, (u) => setUser(u));
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (!window.jStat) {
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jstat/1.9.6/jstat.min.js';
      document.head.appendChild(script);
    }

    if (user && db) {
      const colRef = collection(db, 'artifacts', appId, 'users', user.uid, 'subjects');
      const unsubscribe = onSnapshot(colRef, (snapshot) => {
        const cloudData = snapshot.docs.map(doc => doc.data());
        setSubjects(cloudData);
        setIsCloudReady(true);
      }, (error) => console.error("Firestore sync error:", error));
      return () => unsubscribe();
    } else {
      const saved = localStorage.getItem('tennis_study_data');
      if (saved) setSubjects(JSON.parse(saved));
      setIsCloudReady(false);
    }
  }, [user]);

  useEffect(() => {
    localStorage.setItem('tennis_study_data', JSON.stringify(subjects));
  }, [subjects]);

  useEffect(() => {
    if (view === 'entry' && currentSubject.id) {
      const cloudSubject = subjects.find(s => s.id === currentSubject.id);
      if (cloudSubject && JSON.stringify(cloudSubject) !== JSON.stringify(currentSubject)) {
        setCurrentSubject(cloudSubject);
      }
    }
  }, [subjects, view]);

  const handleStartNew = async () => {
    let maxNum = 0;
    let prefix = '';
    subjects.forEach(sub => {
      const match = String(sub.subjectNo || '').match(/^(\D*)(\d+)$/);
      if (match) {
        const currentPrefix = match[1];
        const num = parseInt(match[2], 10);
        if (num > maxNum) {
          maxNum = num;
          prefix = currentPrefix;
        }
      }
    });

    const nextNum = maxNum + 1;
    const paddedNum = String(nextNum).padStart(3, '0');
    const nextSubjectNo = `${prefix}${paddedNum}`;
    const newSubject = { ...emptySubject, id: Date.now().toString(), subjectNo: nextSubjectNo };
    setCurrentSubject(newSubject);
    
    if (user && db) {
      try {
        const docRef = doc(db, 'artifacts', appId, 'users', user.uid, 'subjects', newSubject.id);
        await setDoc(docRef, newSubject);
      } catch (e) { console.error("Init document failed", e); }
    } else {
      setSubjects(prev => [...prev, newSubject]);
    }
    setView('entry');
  };

  const handleEdit = (sub) => {
    if (sub.smrt && sub.smrt.checklist && sub.smrt.checklist.helpNeeded === undefined) sub.smrt.checklist.helpNeeded = false;
    if (sub.kpull && !sub.kpull['ER_Max Value (kg)']) {
      kpullMetricsList.forEach(m => {
        if (!sub.kpull[`ER_${m}`]) sub.kpull[`ER_${m}`] = ['', '', ''];
        if (!sub.kpull[`IR_${m}`]) sub.kpull[`IR_${m}`] = ['', '', ''];
      });
      if (sub.kpull.peakForceER) sub.kpull['ER_Max Value (kg)'] = sub.kpull.peakForceER;
      if (sub.kpull.peakForceIR) sub.kpull['IR_Max Value (kg)'] = sub.kpull.peakForceIR;
    }
    if (!sub.gender) sub.gender = '男性';
    setCurrentSubject(sub);
    setView('entry');
  };

  const handleDelete = async (id) => {
    try {
      if (user && db) {
        const docRef = doc(db, 'artifacts', appId, 'users', user.uid, 'subjects', id);
        await deleteDoc(docRef);
      } else {
        setSubjects(subjects.filter(s => s.id !== id));
      }
    } catch (error) {
      console.error("Delete failed, fallback to local:", error);
      setSubjects(subjects.filter(s => s.id !== id));
    }
  };

  const handleAutoSave = async (updatedSubject) => {
    if (user && db) {
      try {
        const docRef = doc(db, 'artifacts', appId, 'users', user.uid, 'subjects', updatedSubject.id);
        await setDoc(docRef, updatedSubject, { merge: true });
      } catch (error) { console.error("AutoSave failed:", error); }
    } else {
      setSubjects(prev => {
        const exists = prev.find(s => s.id === updatedSubject.id);
        if (exists) return prev.map(s => s.id === updatedSubject.id ? updatedSubject : s);
        return [...prev, updatedSubject];
      });
    }
  };

  const saveSubject = async () => {
    try {
      if (user && db) {
        const docRef = doc(db, 'artifacts', appId, 'users', user.uid, 'subjects', currentSubject.id);
        await setDoc(docRef, currentSubject);
      } else {
        setSubjects(prev => {
          const exists = prev.find(s => s.id === currentSubject.id);
          if (exists) return prev.map(s => s.id === currentSubject.id ? currentSubject : s);
          return [...prev, currentSubject];
        });
      }
    } catch (error) {
      console.error("Save failed, fallback to local:", error);
      setSubjects(prev => {
        const exists = prev.find(s => s.id === currentSubject.id);
        if (exists) return prev.map(s => s.id === currentSubject.id ? currentSubject : s);
        return [...prev, currentSubject];
      });
    } finally {
      setView('dashboard');
    }
  };

  const exportToExcel = async () => {
    if (subjects.length === 0) return;
    setIsExporting(true);

    try {
      if (!window.XLSX) {
        await new Promise((resolve, reject) => {
          const script = document.createElement('script');
          script.src = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
          script.onload = () => resolve();
          script.onerror = reject;
          document.body.appendChild(script);
        });
      }

      const XLSX = window.XLSX;
      const dateStr = new Date().toISOString().split('T')[0];
      const sanitizeHeader = (str) => str.replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
      const translateGender = (g) => g === '女性' ? 0 : 1;
      const translateHand = (h) => h === '左' ? 1 : 0;
      const translateGroup = (g) => {
        if (g === '肩峰下疼痛組') return 1;
        if (g === '肩痛病史組') return 3;
        return 0; 
      };

      const flattenSubject = (sub) => {
        const flat = {
          'SystemID': sub.id,
          'SubjectID': sub.subjectNo,
          'Name': sub.name,
          'Gender': translateGender(sub.gender),
          'Age': sub.age,
          'Height_cm': sub.height,
          'Weight_kg': sub.weight,
          'DominantHand': translateHand(sub.dominantHand),
          'TennisAge_yr': sub.tennisAge,
          'Training_hr': sub.trainingHours,
          'NTRP': sub.ntrp,
          'ArmLength_cm': sub.armLength,
          'Group': translateGroup(sub.group),
          'PainDuration': sub.painDuration,
          'PainVAS': sub.painVas,
          'Notes': sub.notes || '',
        };

        const flattenTrials = (prefix, arr) => {
          arr.forEach((val, idx) => flat[`${prefix}_T${idx + 1}`] = val);
          flat[`${prefix}_Mean`] = calcMean(arr) || '';
          flat[`${prefix}_SD`] = calcSD(arr) || '';
        };

        Object.entries(sub.rom).forEach(([key, arr]) => flattenTrials(`ROM_${key}`, arr));
        flat['ROM_Total_ActiveLeft'] = calcTotalROM(calcMean(sub.rom.activeLeftIR), calcMean(sub.rom.activeLeftER)) || '';
        flat['ROM_Total_ActiveRight'] = calcTotalROM(calcMean(sub.rom.activeRightIR), calcMean(sub.rom.activeRightER)) || '';
        
        flat['SMRT_Result'] = sub.smrt.result === '陽性' ? 1 : 0;
        Object.entries(sub.smrt.checklist).forEach(([key, val]) => flat[`SMRT_${key}`] = val ? 1 : 0);

        const armLen = Number(sub.armLength);
        Object.entries(sub.ybt).forEach(([key, arr]) => {
          flattenTrials(`YBT_${key}`, arr);
          if (armLen > 0) {
            const sum = calcSum(arr);
            flat[`YBT_${key}_NormPct`] = sum ? ((Number(sum) / armLen) * 100).toFixed(2) : '';
          }
        });
        if (armLen > 0) {
          const mMean = calcMean(sub.ybt.medial);
          const ilMean = calcMean(sub.ybt.il);
          const slMean = calcMean(sub.ybt.sl);
          if (mMean && ilMean && slMean) {
            flat['YBT_CompositeScorePct'] = (((Number(mMean) + Number(ilMean) + Number(slMean)) / (3 * armLen)) * 100).toFixed(2);
          }
        }

        flat['SET_Duration_sec'] = sub.setDuration;

        kpullMetricsList.forEach(m => {
          const cleanM = sanitizeHeader(m);
          const erArr = sub.kpull[`ER_${m}`] || ['', '', ''];
          const irArr = sub.kpull[`IR_${m}`] || ['', '', ''];
          flattenTrials(`KPull_ER_${cleanM}`, erArr);
          flattenTrials(`KPull_IR_${cleanM}`, irArr);
        });
        
        const weight = Number(sub.weight);
        const erMaxArr = sub.kpull['ER_Max Value (kg)'] || sub.kpull.peakForceER || ['', '', ''];
        const irMaxArr = sub.kpull['IR_Max Value (kg)'] || sub.kpull.peakForceIR || ['', '', ''];
        const erMean = calcMean(erMaxArr);
        const irMean = calcMean(irMaxArr);
        
        if (weight > 0) {
          flat['KPull_ER_Max_NormBW'] = erMean ? ((Number(erMean) / weight) * 100).toFixed(2) : '';
          flat['KPull_IR_Max_NormBW'] = irMean ? ((Number(irMean) / weight) * 100).toFixed(2) : '';
        }
        flat['KPull_ER_IR_MaxRatio'] = (erMean && irMean && Number(irMean) !== 0) ? (Number(erMean) / Number(irMean)).toFixed(2) : '';

        return flat;
      };

      const flatData = subjects.map(flattenSubject);
      const headers = Object.keys(flatData[0]);
      const csvRows = [headers.join(',')];

      for (const row of flatData) {
        const values = headers.map(header => {
          const escaped = ('' + (row[header] ?? '')).replace(/"/g, '""');
          return `"${escaped}"`;
        });
        csvRows.push(values.join(','));
      }

      const csvContent = '\uFEFF' + csvRows.join('\n');
      const csvBlob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const csvLink = document.createElement('a');
      csvLink.href = URL.createObjectURL(csvBlob);
      csvLink.download = `Tennis_Clinical_Data_${dateStr}.csv`;
      csvLink.click();

      // Excel (Multi-Sheet)
      const wb = XLSX.utils.book_new();

      const addTrialsToRow = (obj, prefix, arr) => {
        arr.forEach((val, idx) => obj[`${prefix}_T${idx + 1}`] = val);
        obj[`${prefix}_Mean`] = calcMean(arr) || '';
        obj[`${prefix}_SD`] = calcSD(arr) || '';
      };

      const basicSheetData = subjects.map(sub => ({
        'SubjectID': sub.subjectNo, 'Name': sub.name, 'Gender': translateGender(sub.gender),
        'Age': sub.age, 'Height_cm': sub.height, 'Weight_kg': sub.weight,
        'DominantHand': translateHand(sub.dominantHand), 'TennisAge_yr': sub.tennisAge, 'Training_hr': sub.trainingHours,
        'NTRP': sub.ntrp, 'ArmLength_cm': sub.armLength, 'Group': translateGroup(sub.group),
        'PainDuration': sub.painDuration, 'PainVAS': sub.painVas, 'Notes': sub.notes || ''
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(basicSheetData), "BasicInfo");

      const romSheetData = subjects.map(sub => {
        const row = { 'SubjectID': sub.subjectNo, 'Name': sub.name };
        Object.entries(sub.rom).forEach(([key, arr]) => addTrialsToRow(row, `ROM_${key}`, arr));
        row['ROM_Total_ActiveLeft'] = calcTotalROM(calcMean(sub.rom.activeLeftIR), calcMean(sub.rom.activeLeftER)) || '';
        row['ROM_Total_ActiveRight'] = calcTotalROM(calcMean(sub.rom.activeRightIR), calcMean(sub.rom.activeRightER)) || '';
        return row;
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(romSheetData), "ROM");

      const smrtSheetData = subjects.map(sub => {
        const row = { 'SubjectID': sub.subjectNo, 'Name': sub.name };
        row['SMRT_Result'] = sub.smrt.result === '陽性' ? 1 : 0;
        Object.entries(sub.smrt.checklist).forEach(([key, val]) => row[`SMRT_${key}`] = val ? 1 : 0);
        return row;
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(smrtSheetData), "SMRT");

      const funcSheetData = subjects.map(sub => {
        const row = { 'SubjectID': sub.subjectNo, 'Name': sub.name };
        const armLen = Number(sub.armLength);
        Object.entries(sub.ybt).forEach(([key, arr]) => {
          addTrialsToRow(row, `YBT_${key}`, arr);
          if (armLen > 0) {
            const sum = calcSum(arr);
            row[`YBT_${key}_NormPct`] = sum ? ((Number(sum) / armLen) * 100).toFixed(2) : '';
          }
        });
        if (armLen > 0) {
          const mMean = calcMean(sub.ybt.medial);
          const ilMean = calcMean(sub.ybt.il);
          const slMean = calcMean(sub.ybt.sl);
          if (mMean && ilMean && slMean) {
            row['YBT_CompositeScorePct'] = (((Number(mMean) + Number(ilMean) + Number(slMean)) / (3 * armLen)) * 100).toFixed(2);
          }
        }
        row['SET_Duration_sec'] = sub.setDuration;
        return row;
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(funcSheetData), "Functional");

      const kpullSheetData = subjects.map(sub => {
        const row = { 'SubjectID': sub.subjectNo, 'Name': sub.name, 'Weight_kg': sub.weight };
        kpullMetricsList.forEach(m => {
          const cleanM = sanitizeHeader(m);
          addTrialsToRow(row, `KPull_ER_${cleanM}`, sub.kpull[`ER_${m}`] || ['', '', '']);
          addTrialsToRow(row, `KPull_IR_${cleanM}`, sub.kpull[`IR_${m}`] || ['', '', '']);
        });
        const weight = Number(sub.weight);
        const erMaxArr = sub.kpull['ER_Max Value (kg)'] || sub.kpull.peakForceER || ['', '', ''];
        const irMaxArr = sub.kpull['IR_Max Value (kg)'] || sub.kpull.peakForceIR || ['', '', ''];
        const erMean = calcMean(erMaxArr);
        const irMean = calcMean(irMaxArr);
        
        if (weight > 0) {
          row['KPull_ER_Max_NormBW'] = erMean ? ((Number(erMean) / weight) * 100).toFixed(2) : '';
          row['KPull_IR_Max_NormBW'] = irMean ? ((Number(irMean) / weight) * 100).toFixed(2) : '';
        }
        row['KPull_ER_IR_MaxRatio'] = (erMean && irMean && Number(irMean) !== 0) ? (Number(erMean) / Number(irMean)).toFixed(2) : '';
        return row;
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(kpullSheetData), "KPull");

      const styleSheetWithColors = (ws) => {
        if (!ws['!ref']) return;
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const headerCellRef = XLSX.utils.encode_cell({ r: 0, c: C });
          const headerCell = ws[headerCellRef];
          if (!headerCell || !headerCell.v) continue;
          
          const headerStr = String(headerCell.v);
          let bgColor = null;
          
          if (headerStr.includes('_Mean')) bgColor = 'D9EAD3';
          else if (headerStr.includes('_SD')) bgColor = 'F4CCCC';

          if (bgColor) {
            for (let R = 0; R <= range.e.r; ++R) {
              const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
              if (!ws[cellRef]) continue;
              ws[cellRef].s = {
                fill: { fgColor: { rgb: bgColor } },
                font: { bold: R === 0 }
              };
            }
          } else {
            if (ws[headerCellRef]) ws[headerCellRef].s = { font: { bold: true } };
          }
        }
      };

      ['BasicInfo', 'ROM', 'SMRT', 'Functional', 'KPull'].forEach(sheetName => {
        if (wb.Sheets[sheetName]) styleSheetWithColors(wb.Sheets[sheetName]);
      });

      setTimeout(() => {
        XLSX.writeFile(wb, `Tennis_Clinical_Data_${dateStr}.xlsx`);
      }, 800);

    } catch (error) {
      console.error("Export failed:", error);
      alert("匯出失敗，請檢查網路連線或稍後再試。");
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="min-h-[100dvh] bg-gray-100 text-gray-900 font-sans p-0 md:p-8 flex flex-col">
      <div className="max-w-5xl mx-auto w-full flex-1 flex flex-col min-h-0">
        
        {view !== 'entry' && isCloudReady && (
          <div className="text-xs font-bold text-blue-600 bg-blue-100 py-1.5 px-3 rounded-b-lg mb-2 inline-flex items-center gap-1.5 self-end">
            <CloudIcon size={14} /> 雲端即時同步中
          </div>
        )}

        {view === 'entry' && (
          <div className="bg-white p-4 shadow-sm border-b md:rounded-t-2xl flex justify-between items-center shrink-0">
            <span className="font-bold text-gray-800 bg-gray-100 px-3 py-1.5 rounded-full text-sm flex items-center gap-2">
              <Users size={16} className="text-blue-600"/>
              {currentSubject.name || '新受試者'} ({currentSubject.subjectNo || '未編號'})
            </span>
            {isCloudReady && <CloudIcon size={20} className="text-blue-500 opacity-70" />}
          </div>
        )}

        <div className={`flex-1 flex flex-col md:bg-white md:shadow-md md:rounded-b-2xl md:border md:border-t-0 min-h-0 ${view === 'entry' ? 'p-0 md:p-8' : 'p-4 md:p-8'}`}>
          {view === 'dashboard' && <Dashboard subjects={subjects} setView={setView} handleStartNew={handleStartNew} />}
          {view === 'list' && <SubjectList subjects={subjects} setView={setView} handleEdit={handleEdit} handleDelete={handleDelete} exportToExcel={exportToExcel} isExporting={isExporting} />}
          {view === 'analysis' && <AnalysisView subjects={subjects} setView={setView} />}
          
          {view === 'entry' && (
            <DomainViewer
              currentSubject={currentSubject}
              setCurrentSubject={setCurrentSubject}
              handleSave={saveSubject}
              handleCancel={() => setView('dashboard')}
              handleAutoSave={handleAutoSave}
            />
          )}
        </div>
      </div>
    </div>
  );
}
