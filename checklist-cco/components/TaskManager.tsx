
import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Task, OperationStatus, User } from '../types';
import { SharePointService } from '../services/sharepointService';
import { 
  Loader2, Database, 
  ShieldCheck,
  Activity, X, LogOut, 
  ChevronDown, ChevronRight, RotateCcw,
  UserCheck, Send, PaintBucket, Maximize2, Minimize2, HelpCircle
} from 'lucide-react';

const STATUS_CONFIG: Record<string, { label: string, color: string, next: OperationStatus, shortcut: string, desc: string }> = {
  'OK': { label: 'OK', color: 'bg-emerald-600 text-white border-emerald-500 shadow-emerald-900/40', next: 'EA', shortcut: '1', desc: 'Concluído' },
  'EA': { label: 'EA', color: 'bg-amber-500 text-slate-900 border-amber-400 shadow-amber-900/40', next: 'AR', shortcut: '2', desc: 'Em Andamento' },
  'ATT': { label: 'ATT', color: 'bg-blue-800 text-white border-blue-700 shadow-blue-950/40', next: 'AT', shortcut: '3', desc: 'Atualizar' },
  'AR': { label: 'AR', color: 'bg-[#5c3e2f] text-white border-[#4d3428] shadow-black/40', next: 'ATT', shortcut: '4', desc: 'Aguardando Retorno' },
  'AT': { label: 'AT', color: 'bg-rose-700 text-white border-rose-600 shadow-rose-950/40', next: 'PR', shortcut: '5', desc: 'Atrasado' },
  'PR': { label: 'PR', color: 'bg-slate-600 text-white border-slate-500 shadow-black/40', next: 'OK', shortcut: '6', desc: 'Pendente' },
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
  teamMembers?: string[];
}

const TaskManager: React.FC<TaskManagerProps> = ({ 
  tasks, 
  setTasks, 
  locations, 
  collapsedCategories,
  setCollapsedCategories,
  currentUser,
  onLogout,
  teamMembers = []
}) => {
  const [activeTool, setActiveTool] = useState<OperationStatus | null>(null);
  const [isUpdating, setIsUpdating] = useState(false);
  const [compact, setCompact] = useState(true);
  
  const [isResetModalOpen, setIsResetModalOpen] = useState(false);
  const [resetResponsible, setResetResponsible] = useState('');

  const [isDragging, setIsDragging] = useState(false);
  const paintedThisDrag = useRef<Set<string>>(new Set());
  
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

  // Fix: Adicionada a função handleResetChecklist para processar o arquivamento e reset dos status
  const handleResetChecklist = async () => {
    if (!resetResponsible || !currentUser.accessToken) return;
    
    setIsUpdating(true);
    try {
      // 1. Arquivar estado atual no histórico do SharePoint
      await SharePointService.saveHistory(currentUser.accessToken, {
        id: Date.now().toString(),
        timestamp: new Date().toISOString(),
        tasks: tasks,
        resetBy: resetResponsible,
        email: currentUser.email
      });

      // 2. Resetar estados locais para 'PR' (Pendente)
      const resetTasks = tasks.map(t => {
        const resetOps: Record<string, OperationStatus> = {};
        locations.forEach(loc => { resetOps[loc] = 'PR'; });
        return { ...t, operations: resetOps };
      });

      // 3. Sincronizar o reset com o SharePoint para o dia atual
      const today = new Date().toISOString().split('T')[0];
      const todayKey = today.replace(/-/g, '');
      
      // Itera sobre as tarefas e operações para resetar no banco de dados
      for (const task of tasks) {
        for (const loc of locations) {
          // Apenas envia update se o status atual não for PR para otimizar
          if (task.operations[loc] !== 'PR') {
            const uniqueKey = `${todayKey}_${task.id}_${loc}`;
            await SharePointService.updateStatus(currentUser.accessToken, {
              DataReferencia: today,
              TarefaID: task.id,
              OperacaoSigla: loc,
              Status: 'PR',
              Usuario: currentUser.name,
              Title: uniqueKey
            });
          }
        }
      }

      setTasks(resetTasks);
      setIsResetModalOpen(false);
      setResetResponsible('');
      alert("Checklist arquivado e resetado com sucesso!");
    } catch (err: any) {
      console.error("Erro ao resetar checklist:", err);
      alert(`Falha ao resetar: ${err.message}`);
    } finally {
      setIsUpdating(false);
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
    <div className="flex flex-col h-full bg-[#020617] text-white rounded-2xl border border-slate-800 shadow-2xl overflow-hidden relative font-sans">
      {/* HEADER / TOOLBAR */}
      <div className="px-6 py-4 border-b border-slate-800 flex flex-col xl:flex-row justify-between items-center bg-[#0f172a] gap-4 shrink-0 z-50 shadow-lg">
        <div className="flex items-center gap-6">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-blue-600 rounded-lg text-white shadow-lg shadow-blue-500/30">
              <Activity size={24} />
            </div>
            <h2 className="text-xl font-black tracking-tight whitespace-nowrap">Checklist CCO</h2>
          </div>
          
          <div className="flex items-center gap-2 px-3 py-1 bg-slate-900/50 rounded-full border border-slate-700/50">
            {isUpdating ? (
              <div className="flex items-center gap-2 text-[10px] text-blue-400 animate-pulse font-black uppercase tracking-widest">
                <Loader2 size={12} className="animate-spin"/> Gravando
              </div>
            ) : (
              <div className="flex items-center gap-2 text-[10px] text-emerald-500 font-black uppercase tracking-widest">
                <ShieldCheck size={12}/> Protegido
              </div>
            )}
          </div>
        </div>

        <div className="flex flex-wrap items-center justify-center gap-4">
          <div className="flex items-center gap-3 bg-black/40 px-4 py-2 rounded-2xl border border-slate-700 shadow-inner">
            <div className="flex items-center gap-2 text-[10px] font-black text-slate-500 uppercase tracking-widest">
              <PaintBucket size={16} className={activeTool ? 'text-blue-500' : 'text-slate-600'} />
              Pincel
            </div>
            <div className="flex items-center gap-1.5">
              {(Object.entries(STATUS_CONFIG) as [string, any][]).map(([key, cfg]) => (
                <button 
                  key={key} 
                  onClick={() => setActiveTool(activeTool === key ? null : key as OperationStatus)} 
                  className={`w-9 h-9 rounded-xl font-black text-[10px] transition-all border-2 flex items-center justify-center shadow-lg ${cfg.color} ${activeTool === key ? 'ring-4 ring-blue-500/30 border-white scale-110 z-10' : 'opacity-80 hover:opacity-100 hover:scale-105 border-transparent'}`}
                >
                  {cfg.label}
                </button>
              ))}
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button onClick={() => setIsResetModalOpen(true)} className="flex items-center gap-2 px-5 py-2.5 bg-amber-600/20 text-amber-500 rounded-2xl hover:bg-amber-600 hover:text-white transition-all border border-amber-600/50 font-bold text-sm shadow-lg shadow-amber-900/20">
              <RotateCcw size={18} />
              Resetar
            </button>
            <button onClick={() => setCompact(!compact)} className="p-2.5 bg-slate-800 text-slate-400 rounded-2xl hover:bg-slate-700 transition-all border border-slate-700">
              {compact ? <Maximize2 size={20}/> : <Minimize2 size={20}/>}
            </button>
            <button onClick={onLogout} className="flex items-center gap-2 px-4 py-2.5 bg-red-600/20 text-red-500 rounded-2xl hover:bg-red-600 hover:text-white transition-all font-bold text-sm border border-red-600/30">
                <LogOut size={18}/> Sair
            </button>
          </div>
        </div>
      </div>

      <div className="px-6 py-2 bg-[#1e293b]/50 text-[10px] font-black uppercase text-slate-400 tracking-widest flex items-center gap-2 shadow-inner">
          Ação / Descrição da Tarefa
      </div>

      {/* TABELA PRINCIPAL */}
      <div className="flex-1 overflow-auto bg-[#020617] transition-colors scrollbar-thin">
        {tasks.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center p-12 text-center text-slate-600">
                <Database size={64} className="mb-6 opacity-10"/>
                <h3 className="text-xl font-black">Nenhuma tarefa disponível</h3>
                <p className="text-sm">Contate o administrador para vincular operações ao seu e-mail.</p>
            </div>
        ) : (
            <table className={`min-w-full border-separate border-spacing-0 select-none ${compact ? 'text-[11px]' : 'text-[13px]'}`}>
              <thead className="sticky top-0 z-[40]">
                {/* Headers moved to a visual block above the table for the task column, locations stay as columns */}
              </thead>
              <tbody>
                {(Object.entries(groupedTasks) as [string, Task[]][]).map(([cat, catTasks]) => {
                  const isCollapsed = collapsedCategories.includes(cat);
                  const { percent } = getCategoryStats(cat);
                  return (
                    <React.Fragment key={cat}>
                      <tr 
                        className={`bg-[#1e3a8a] text-white h-12 cursor-pointer hover:bg-blue-800 transition-all group sticky top-0 z-[35]`} 
                        onClick={() => toggleCategory(cat)}
                      >
                        <td colSpan={locations.length + 1} className="p-0 border-y border-blue-900/50">
                          <div className={`absolute inset-y-0 left-0 transition-all duration-1000 bg-white/5`} style={{ width: `${percent}%` }} />
                          <div className="relative px-6 flex items-center justify-between z-10 w-full h-full">
                            <div className="flex items-center gap-3">
                              {isCollapsed ? <ChevronRight size={16} strokeWidth={3}/> : <ChevronDown size={16} strokeWidth={3}/>}
                              <span className="text-[11px] font-black uppercase tracking-[0.2em]">{cat}</span>
                            </div>
                            <span className="text-[10px] font-black bg-black/40 px-3 py-1 rounded-full border border-white/10">{percent}%</span>
                          </div>
                        </td>
                      </tr>
                      {!isCollapsed && catTasks.map(task => (
                        <tr key={task.id} className="bg-slate-900/20 border-b border-slate-800 hover:bg-slate-800/30 transition-colors group">
                          <td className="p-6 border-r border-slate-800 sticky left-0 bg-[#020617] z-30 min-w-[400px]">
                            <div className="flex flex-col gap-2">
                                <div className="font-bold text-slate-100 text-[15px] leading-snug">
                                    {task.title}
                                </div>
                                {task.description && (
                                    <div className="text-[12px] text-slate-500 leading-relaxed max-w-2xl">
                                        {task.description}
                                    </div>
                                )}
                            </div>
                          </td>
                          {locations.map(loc => {
                            const status = task.operations[loc] || 'PR';
                            const cfg = STATUS_CONFIG[status];
                            return (
                              <td 
                                key={loc} 
                                className="p-0 border-r border-slate-800 w-28 h-20 relative group/cell" 
                                onMouseDown={() => handleUpdateStatus(task.id, loc, activeTool || cfg.next)}
                              >
                                <div className={`absolute inset-[6px] rounded-2xl flex items-center justify-center font-black text-[11px] transition-all duration-200 cursor-pointer shadow-xl border border-white/5 active:scale-95 ${cfg.color}`}>
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

      <div className="px-6 py-3 bg-[#0f172a] border-t border-slate-800 flex justify-between items-center text-[10px] text-slate-500 font-black uppercase tracking-widest shrink-0">
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2 px-3 py-1 bg-emerald-500/10 text-emerald-500 rounded-full border border-emerald-500/20">
                <div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></div>
                OPERADOR: {currentUser.name}
            </div>
          </div>
          <div className="flex items-center gap-3">
              <HelpCircle size={14} className="text-blue-500"/>
              ATALHOS: (1-6) PINTAR, (ARRASTE) PINTURA CONTÍNUA, (TÍTULO DA AÇÃO) PINTAR LINHA TODA
          </div>
      </div>

      {/* RESET MODAL */}
      {isResetModalOpen && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-xl z-[100] flex items-center justify-center p-4">
             <div className="bg-[#0f172a] rounded-[2.5rem] shadow-2xl w-full max-w-md overflow-hidden border border-slate-700 animate-in zoom-in duration-300">
                <div className="bg-amber-600 text-white p-8">
                    <div className="flex items-center gap-4 mb-4">
                        <div className="p-3 bg-white/20 rounded-2xl">
                            <RotateCcw size={32} />
                        </div>
                        <h3 className="font-black text-3xl uppercase tracking-tight">Resetar</h3>
                    </div>
                    <p className="text-amber-100 font-medium opacity-80">Arquive os dados atuais e comece uma nova rodada do checklist.</p>
                </div>
                
                <div className="p-10">
                    <div className="mb-10">
                        <label className="block text-xs font-black text-slate-500 uppercase mb-4 tracking-[0.2em]">Responsável</label>
                        <div className="relative group">
                            <div className="absolute inset-y-0 left-0 pl-5 flex items-center pointer-events-none text-amber-500">
                                <UserCheck size={24} />
                            </div>
                            <select 
                                value={resetResponsible}
                                onChange={(e) => setResetResponsible(e.target.value)}
                                className="w-full pl-14 p-5 border-2 border-slate-700 rounded-3xl bg-slate-900 text-xl font-bold text-white focus:border-amber-500 focus:ring-4 focus:ring-amber-500/10 outline-none transition-all appearance-none"
                            >
                                <option value="" disabled>Selecione seu nome...</option>
                                {teamMembers.map(name => (
                                    <option key={name} value={name}>{name}</option>
                                ))}
                            </select>
                            <div className="absolute inset-y-0 right-0 pr-5 flex items-center pointer-events-none text-slate-500">
                                <ChevronDown size={24} />
                            </div>
                        </div>
                    </div>

                    <div className="flex gap-4">
                        <button 
                            onClick={() => setIsResetModalOpen(false)} 
                            className="flex-1 py-5 bg-slate-800 text-slate-400 font-black rounded-3xl hover:bg-slate-700 transition-all uppercase tracking-widest text-xs"
                        >
                            Cancelar
                        </button>
                        <button 
                            onClick={handleResetChecklist} 
                            disabled={!resetResponsible || isUpdating} 
                            className="flex-[2] py-5 bg-amber-600 text-white font-black rounded-3xl shadow-xl shadow-amber-900/40 flex items-center justify-center gap-3 transition-all hover:bg-amber-500 active:scale-95 disabled:opacity-50 uppercase tracking-widest text-xs"
                        >
                            {isUpdating ? <Loader2 size={24} className="animate-spin" /> : <Send size={24} />}
                            Confirmar
                        </button>
                    </div>
                </div>
             </div>
        </div>
      )}
    </div>
  );
};

export default TaskManager;
