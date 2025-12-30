
import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Task, OperationStatus, User } from '../types';
import { SharePointService } from '../services/sharepointService';
import { parseExcelContentToTasks } from '../services/geminiService';
import { 
  Maximize2, Minimize2, Loader2, Database, 
  ShieldCheck, RefreshCw, CheckCircle,
  Activity, PaintBucket, HelpCircle, X, LogOut, 
  ChevronDown, ChevronRight, RotateCcw, Save, 
  UserCheck, Upload, Sparkles, Send
} from 'lucide-react';

const STATUS_CONFIG: Record<string, { label: string, color: string, next: OperationStatus, shortcut: string, desc: string }> = {
  'OK': { label: 'OK', color: 'bg-emerald-500 text-white border-emerald-600 shadow-emerald-500/20 dark:bg-emerald-600', next: 'EA', shortcut: '1', desc: 'Concluído' },
  'EA': { label: 'EA', color: 'bg-amber-400 text-slate-900 border-amber-500 shadow-amber-400/20 dark:bg-amber-500 dark:text-slate-950', next: 'AR', shortcut: '2', desc: 'Em Andamento' },
  'ATT': { label: 'ATT', color: 'bg-blue-800 text-white border-blue-900 shadow-blue-800/20 dark:bg-blue-900', next: 'AT', shortcut: '3', desc: 'Atualizar' },
  'AR': { label: 'AR', color: 'bg-orange-500 text-white border-orange-600 shadow-orange-500/20 dark:bg-orange-600', next: 'ATT', shortcut: '4', desc: 'Aguardando Retorno' },
  'AT': { label: 'AT', color: 'bg-rose-600 text-white border-rose-700 shadow-rose-600/20 dark:bg-rose-700', next: 'PR', shortcut: '5', desc: 'Atrasado' },
  'PR': { label: 'PR', color: 'bg-slate-400 text-white border-slate-500 shadow-slate-400/20 dark:bg-slate-500', next: 'OK', shortcut: '6', desc: 'Pendente' },
};

interface TaskManagerProps {
  tasks: Task[];
  setTasks: React.Dispatch<React.SetStateAction<Task[]>>;
  locations: string[];
  setLocations: any;
  onUserSwitch: any;
  collapsedCategories: string[];
  setCollapsedCategories: any;
  currentUser: User;
  onLogout: () => void;
}

const TaskManager: React.FC<TaskManagerProps> = ({ 
  tasks, 
  setTasks, 
  locations, 
  collapsedCategories,
  setCollapsedCategories,
  onUserSwitch, 
  currentUser,
  onLogout
}) => {
  const [activeTool, setActiveTool] = useState<OperationStatus | null>(null);
  const [isUpdating, setIsUpdating] = useState(false);
  const [compact, setCompact] = useState(true);
  
  const [isResetModalOpen, setIsResetModalOpen] = useState(false);
  const [resetResponsible, setResetResponsible] = useState('');

  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [importText, setImportText] = useState('');
  const [isProcessingImport, setIsProcessingImport] = useState(false);

  const [isDragging, setIsDragging] = useState(false);
  const paintedThisDrag = useRef<Set<string>>(new Set());
  
  const autoCollapsedSessionRef = useRef<Set<string>>(new Set());
  const manuallyOpenedRef = useRef<Set<string>>(new Set());

  const getCategoryStats = (category: string) => {
    const catTasks = tasks.filter(t => (t.category || 'Geral') === category);
    if (catTasks.length === 0) return { percent: 0, isComplete: false };
    let totalCells = 0, okCells = 0;
    catTasks.forEach(t => {
      locations.forEach(loc => {
        totalCells++;
        if (t.operations[loc] === 'OK') okCells++;
      });
    });
    const percent = totalCells === 0 ? 0 : Math.round((okCells / totalCells) * 100);
    return { percent, isComplete: percent === 100 };
  };

  useEffect(() => {
    const categories = Array.from(new Set<string>(tasks.map(t => t.category || 'Geral')));
    categories.forEach((cat: string) => {
        const stats = getCategoryStats(cat);
        if (stats.isComplete && !collapsedCategories.includes(cat) && !manuallyOpenedRef.current.has(cat)) {
            setCollapsedCategories((prev: string[]) => [...prev, cat]);
            autoCollapsedSessionRef.current.add(cat);
        }
    });
  }, [tasks]);

  useEffect(() => {
    const handleMouseUp = () => { setIsDragging(false); paintedThisDrag.current.clear(); };
    window.addEventListener('mouseup', handleMouseUp);
    return () => window.removeEventListener('mouseup', handleMouseUp);
  }, []);

  const handleUpdateStatus = async (taskId: string, location: string, status: OperationStatus) => {
    if (!currentUser.accessToken) return;
    const originalTasks = [...tasks];
    setTasks(prev => prev.map(t => t.id === taskId ? { 
      ...t, operations: { ...t.operations, [location]: status } 
    } : t));
    
    setIsUpdating(true);
    try {
      const today = new Date().toISOString().split('T')[0];
      const todayKey = today.replace(/-/g, '');
      const uniqueKey = `${todayKey}_${taskId}_${location}`;
      await SharePointService.updateStatus(currentUser.accessToken, {
        DataReferencia: today, TarefaID: taskId, OperacaoSigla: location, Status: status,
        Usuario: currentUser.name, Title: uniqueKey
      });
    } catch (err: any) {
      setTasks(originalTasks);
      alert(`Falha ao salvar: ${err.message}`);
    } finally {
      setIsUpdating(false);
    }
  };

  const handleResetChecklist = async () => {
    if (!resetResponsible.trim() || !currentUser.accessToken) return;
    
    setIsUpdating(true);
    try {
        // Salva histórico
        await SharePointService.saveHistory(currentUser.accessToken, {
            id: Date.now().toString(),
            timestamp: new Date().toISOString(),
            tasks: tasks,
            resetBy: resetResponsible,
            email: currentUser.email
        });
        
        const today = new Date().toISOString().split('T')[0];
        const todayKey = today.replace(/-/g, '');
        
        // Reset em lote
        const promises = [];
        for (const task of tasks) {
            for (const loc of locations) {
                const uniqueKey = `${todayKey}_${task.id}_${loc}`;
                promises.push(SharePointService.updateStatus(currentUser.accessToken!, {
                    DataReferencia: today, TarefaID: task.id, OperacaoSigla: loc,
                    Status: 'PR', Usuario: resetResponsible, Title: uniqueKey
                }));
            }
        }
        await Promise.all(promises);

        setTasks(prev => prev.map(t => ({
            ...t, operations: locations.reduce((acc, loc) => ({ ...acc, [loc]: 'PR' }), {})
        })));

        setIsResetModalOpen(false);
        setResetResponsible('');
        alert("Checklist resetado e salvo no histórico!");
    } catch (error: any) {
        alert(`Erro ao resetar: ${error.message}`);
    } finally {
        setIsUpdating(false);
    }
  };

  const handleImportChecklist = async () => {
    if (!importText.trim() || !currentUser.accessToken) return;
    setIsProcessingImport(true);
    try {
      const parsedTasks = await parseExcelContentToTasks(importText);
      if (parsedTasks.length === 0) { alert("Nenhuma tarefa identificada."); return; }

      if (confirm(`Importar ${parsedTasks.length} tarefas para o SharePoint?`)) {
        for (const spTask of parsedTasks) {
          await SharePointService.createTask(currentUser.accessToken, spTask);
        }
        alert("Importação concluída! Recarregando dados...");
        onUserSwitch();
        setIsImportModalOpen(false);
        setImportText('');
      }
    } catch (error: any) {
      alert("Erro na importação: " + error.message);
    } finally {
      setIsProcessingImport(false);
    }
  };

  const toggleCategory = (cat: string) => {
    const isCurrentlyCollapsed = collapsedCategories.includes(cat);
    if (isCurrentlyCollapsed) {
      setCollapsedCategories((prev: string[]) => prev.filter(c => c !== cat));
      manuallyOpenedRef.current.add(cat);
    } else {
      setCollapsedCategories((prev: string[]) => [...prev, cat]);
      manuallyOpenedRef.current.delete(cat);
    }
  };

  const groupedTasks = useMemo(() => tasks.reduce((acc, t) => {
    const cat = t.category || 'Geral';
    if (!acc[cat]) acc[cat] = [];
    acc[cat].push(t);
    return acc;
  }, {} as Record<string, Task[]>), [tasks]);

  return (
    <div className="flex flex-col h-full bg-white dark:bg-slate-900 rounded-2xl border dark:border-slate-800 shadow-sm overflow-hidden relative font-sans">
      {/* HEADER / TOOLBAR */}
      <div className="px-4 py-3 border-b dark:border-slate-800 flex flex-col xl:flex-row justify-between items-center bg-gray-50/80 dark:bg-slate-800/80 backdrop-blur-md gap-3 shrink-0 z-50">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2">
            <div className="p-2 bg-blue-600 rounded-lg text-white shadow-lg shadow-blue-500/20">
              <Activity size={20} />
            </div>
            <h2 className="text-lg font-bold text-gray-800 dark:text-white whitespace-nowrap">Checklist CCO</h2>
          </div>
          <div className="h-6 w-px bg-gray-300 dark:bg-slate-700 hidden md:block" />
          {isUpdating ? (
            <div className="flex items-center gap-2 text-[10px] text-blue-500 animate-pulse font-black uppercase tracking-widest">
              <Loader2 size={12} className="animate-spin"/> Gravando
            </div>
          ) : (
            <div className="flex items-center gap-2 text-[10px] text-green-500 font-bold uppercase tracking-widest">
              <ShieldCheck size={12}/> Protegido
            </div>
          )}
        </div>

        <div className="flex flex-wrap items-center justify-center gap-3">
          {/* PINCEL DE STATUS */}
          <div className="flex items-center gap-2 bg-white dark:bg-slate-950 px-3 py-1.5 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm">
            <div className="flex items-center gap-1">
              {(Object.entries(STATUS_CONFIG) as [string, any][]).map(([key, cfg]) => (
                <button 
                  key={key} 
                  onClick={() => setActiveTool(activeTool === key ? null : key as OperationStatus)} 
                  className={`w-8 h-8 rounded-lg font-black text-[11px] transition-all border flex items-center justify-center relative group shadow-sm ${cfg.color} ${activeTool === key ? 'ring-2 ring-offset-2 ring-blue-500 scale-110 z-10' : 'opacity-85 hover:opacity-100 hover:scale-105'}`}
                  title={`${cfg.desc}`}
                >
                  {cfg.label}
                </button>
              ))}
            </div>
          </div>

          <div className="flex items-center gap-1">
            <button onClick={() => setIsImportModalOpen(true)} className="flex items-center gap-2 px-3 py-2 bg-emerald-50 dark:bg-emerald-900/30 text-emerald-600 dark:text-emerald-400 rounded-xl hover:bg-emerald-100 transition-all border border-emerald-100 dark:border-emerald-800">
              <Upload size={18} />
              <span className="text-xs font-bold hidden sm:inline">Importar Excel</span>
            </button>
            <button onClick={() => { setResetResponsible(''); setIsResetModalOpen(true); }} className="flex items-center gap-2 px-3 py-2 bg-amber-50 dark:bg-amber-900/30 text-amber-600 dark:text-amber-400 rounded-xl hover:bg-amber-100 transition-all border border-amber-100 dark:border-amber-800">
              <RotateCcw size={18} />
              <span className="text-xs font-bold hidden sm:inline">Resetar</span>
            </button>
            <button onClick={onLogout} className="flex items-center gap-2 px-3 py-1.5 bg-red-50 dark:bg-red-900/30 text-red-600 dark:text-red-400 rounded-xl hover:bg-red-100 transition-all font-bold text-xs">
                <LogOut size={16}/> Sair
            </button>
          </div>
        </div>
      </div>

      {/* MODAL DE IMPORTAÇÃO */}
      {isImportModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
             <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden border dark:border-slate-700">
                <div className="bg-emerald-800 text-white p-4 flex justify-between items-center">
                    <h3 className="font-bold">Importar Tarefas Excel</h3>
                    <button onClick={() => setIsImportModalOpen(false)}><X size={24} /></button>
                </div>
                <div className="p-6">
                    <textarea 
                        value={importText} onChange={(e) => setImportText(e.target.value)}
                        className="w-full h-64 p-4 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-xs font-mono dark:text-white mb-4 outline-none"
                        placeholder="Cole aqui os dados copiados do seu Excel..."
                    />
                    <div className="flex gap-3">
                        <button onClick={() => setIsImportModalOpen(false)} className="flex-1 py-3 bg-gray-200 dark:bg-slate-700 rounded-xl font-bold">Cancelar</button>
                        <button onClick={handleImportChecklist} disabled={isProcessingImport} className="flex-[2] py-3 bg-emerald-600 text-white font-bold rounded-xl flex items-center justify-center gap-2">
                            {isProcessingImport ? <Loader2 size={20} className="animate-spin" /> : <Sparkles size={20} />}
                            Processar com IA
                        </button>
                    </div>
                </div>
             </div>
        </div>
      )}

      {/* MODAL DE RESET (SIMPLIFICADO E SEM TRAVAS) */}
      {isResetModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
             <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-md overflow-hidden border dark:border-slate-700 animate-in zoom-in duration-200">
                <div className="bg-amber-600 text-white p-6">
                    <div className="flex items-center gap-3 mb-2">
                        <RotateCcw size={24} />
                        <h3 className="font-black text-xl uppercase tracking-tight">Resetar Checklist</h3>
                    </div>
                    <p className="text-amber-100 text-xs font-medium">Os dados atuais serão movidos para o histórico permanente no SharePoint.</p>
                </div>
                
                <div className="p-8 bg-gray-50 dark:bg-slate-900">
                    <div className="mb-8">
                        <label className="block text-[10px] font-black text-gray-400 uppercase mb-3 tracking-widest">Quem está realizando o reset?</label>
                        <div className="relative group">
                            <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none text-amber-500">
                                <UserCheck size={20} />
                            </div>
                            <input 
                                type="text"
                                value={resetResponsible}
                                onChange={(e) => setResetResponsible(e.target.value)}
                                className="w-full pl-12 p-4 border-2 border-slate-200 dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-lg font-bold dark:text-white focus:border-amber-500 focus:ring-4 focus:ring-amber-500/10 outline-none transition-all placeholder:text-slate-300"
                                placeholder="Digite seu nome completo..."
                                autoFocus
                            />
                        </div>
                    </div>

                    <div className="flex gap-3">
                        <button 
                            onClick={() => setIsResetModalOpen(false)} 
                            className="flex-1 py-4 bg-slate-200 dark:bg-slate-700 text-slate-600 dark:text-slate-200 font-black rounded-2xl hover:bg-slate-300 transition-all uppercase text-xs"
                        >
                            Cancelar
                        </button>
                        <button 
                            onClick={handleResetChecklist} 
                            disabled={!resetResponsible.trim() || isUpdating} 
                            className="flex-[2] py-4 bg-amber-600 text-white font-black rounded-2xl shadow-xl flex items-center justify-center gap-3 transition-all hover:bg-amber-700 active:scale-95 disabled:opacity-50 uppercase text-xs"
                        >
                            {isUpdating ? <Loader2 size={20} className="animate-spin" /> : <Send size={20} />}
                            Confirmar Reset
                        </button>
                    </div>
                </div>
             </div>
        </div>
      )}

      {/* TABELA PRINCIPAL */}
      <div className="flex-1 overflow-auto bg-slate-100 dark:bg-slate-950 transition-colors scrollbar-thin">
        {tasks.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center p-12 text-center text-slate-400">
                <Database size={48} className="mb-6 opacity-20"/>
                <h3 className="text-lg font-black dark:text-white">Nenhuma tarefa no sistema</h3>
                <p className="text-sm">Use o botão "Importar Excel" para migrar seus dados.</p>
            </div>
        ) : (
            <table className={`min-w-full border-separate border-spacing-0 select-none ${compact ? 'text-[10px]' : 'text-[11px]'}`}>
              <thead className="sticky top-0 z-[40]">
                <tr className="bg-blue-900 text-white shadow-xl">
                  <th className="p-3 border-r border-blue-800 text-left sticky left-0 bg-blue-900 z-[45] min-w-[350px] font-black uppercase tracking-widest text-[9px]">Ação / Tarefa</th>
                  {locations.map(loc => (
                    <th key={loc} className="p-3 border-r border-blue-800 w-24 text-center font-bold">{loc.replace('LAT-', '').replace('ITA-', '')}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {(Object.entries(groupedTasks) as [string, Task[]][]).map(([cat, catTasks]) => {
                  const isCollapsed = collapsedCategories.includes(cat);
                  const { percent, isComplete } = getCategoryStats(cat);
                  return (
                    <React.Fragment key={cat}>
                      <tr 
                        className={`bg-blue-600 text-white h-10 cursor-pointer hover:bg-blue-700 transition-all`} 
                        onClick={() => toggleCategory(cat)}
                      >
                        <td colSpan={locations.length + 1} className="p-0 border-y border-blue-700 sticky left-0 z-30 overflow-hidden">
                          <div className={`absolute inset-y-0 left-0 transition-all duration-1000 bg-green-500/30`} style={{ width: `${percent}%` }} />
                          <div className="absolute inset-0 px-4 flex items-center justify-between z-10">
                            <div className="flex items-center gap-3">
                              {isCollapsed ? <ChevronRight size={14}/> : <ChevronDown size={14}/>}
                              <span className="text-[10px] font-black uppercase tracking-widest">{cat}</span>
                            </div>
                            <span className="text-[9px] font-black bg-black/20 px-2 py-0.5 rounded-lg">{percent}%</span>
                          </div>
                        </td>
                      </tr>
                      {!isCollapsed && catTasks.map(task => (
                        <tr key={task.id} className="bg-white dark:bg-slate-900 border-b dark:border-slate-800/50 hover:bg-blue-50/30 transition-colors group">
                          <td className="p-4 border-r dark:border-slate-800 sticky left-0 bg-inherit z-30 shadow-[4px_0_8px_-4px_rgba(0,0,0,0.1)]">
                            <div className="font-bold text-slate-800 dark:text-slate-100 text-[13px]">
                                <span className="text-[9px] text-blue-500 font-mono mr-2">[{task.timeRange}]</span>
                                {task.title}
                            </div>
                            {task.description && <div className="text-[11px] text-slate-500 dark:text-slate-400 mt-1">{task.description}</div>}
                          </td>
                          {locations.map(loc => {
                            const status = task.operations[loc] || 'PR';
                            const cfg = STATUS_CONFIG[status];
                            return (
                              <td 
                                key={loc} 
                                className="p-0 border-r dark:border-slate-800 h-12 relative cursor-pointer" 
                                onMouseDown={() => {
                                    const next = activeTool || cfg.next;
                                    handleUpdateStatus(task.id, loc, next);
                                }}
                              >
                                <div className={`absolute inset-[3px] rounded-lg flex items-center justify-center font-black text-[12px] ${cfg.color} hover:brightness-110 active:scale-90 shadow transition-all border`}>
                                  {cfg.label}
                                </div>
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
        )}
      </div>
    </div>
  );
};

export default TaskManager;
