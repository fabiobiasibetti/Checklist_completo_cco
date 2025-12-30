
import { HistoryRecord, Task, User as AppUser } from '../types';
import { SharePointService } from '../services/sharepointService';
import { History, Calendar, Clock, ChevronRight, ChevronDown, CheckCircle2, User, Loader2, MapPin, Eye, FileSearch } from 'lucide-react';
import React, { useState, useEffect, useMemo } from 'react';

const STATUS_CONFIG: Record<string, { label: string, color: string }> = {
  'PR': { label: 'PR', color: 'bg-slate-200 text-slate-600 border-slate-300 dark:bg-slate-700 dark:text-slate-300 dark:border-slate-600' },
  'OK': { label: 'OK', color: 'bg-green-200 text-green-800 border-green-300 dark:bg-green-900/60 dark:text-green-300 dark:border-green-800' },
  'EA': { label: 'EA', color: 'bg-yellow-200 text-yellow-800 border-yellow-300 dark:bg-yellow-900/60 dark:text-yellow-300 dark:border-yellow-800' },
  'AR': { label: 'AR', color: 'bg-orange-200 text-orange-800 border-orange-300 dark:bg-orange-900/60 dark:text-orange-300 dark:border-orange-800' },
  'ATT': { label: 'ATT', color: 'bg-blue-200 text-blue-800 border-blue-300 dark:bg-blue-900/60 dark:text-blue-300 dark:border-blue-800' },
  'AT': { label: 'AT', color: 'bg-red-500 text-white border-red-600 dark:bg-red-800 dark:text-white dark:border-red-700' },
};

interface HistoryViewerProps {
    currentUser: AppUser;
}

type EnhancedHistoryRecord = HistoryRecord & {
    hasPartial?: boolean;
    partialRecord?: HistoryRecord;
};

const HistoryViewer: React.FC<HistoryViewerProps> = ({ currentUser }) => {
  const [history, setHistory] = useState<HistoryRecord[]>([]);
  const [selectedRecord, setSelectedRecord] = useState<EnhancedHistoryRecord | null>(null);
  const [viewingPartial, setViewingPartial] = useState(false);
  const [collapsedCategories, setCollapsedCategories] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    const fetchFromSP = async () => {
        const token = currentUser.accessToken || (window as any).__access_token; 
        if (!token) {
            setIsLoading(false);
            return;
        }
        
        setIsLoading(true);
        try {
            const data = await SharePointService.getHistory(token, currentUser.email);
            setHistory(data);
        } catch (e) {
            console.error("Erro ao carregar histórico:", e);
        } finally {
            setIsLoading(false);
        }
    };
    fetchFromSP();
  }, [currentUser]);

  const displayTimestamp = (timestamp: string) => {
    try {
        if (!timestamp) return "--/--/---- --:--";
        if (timestamp.includes('/') && !timestamp.includes('T')) return timestamp; 
        const date = new Date(timestamp);
        return date.toLocaleString('pt-BR');
    } catch(e) {
        return timestamp;
    }
  };

  const processedHistory = useMemo(() => {
    const groupedByDay: Record<string, HistoryRecord[]> = {};
    
    history.forEach(rec => {
        const dateStr = displayTimestamp(rec.timestamp).split(',')[0].trim(); // Get DD/MM/YYYY
        if (!groupedByDay[dateStr]) groupedByDay[dateStr] = [];
        groupedByDay[dateStr].push(rec);
    });

    const finalHistory: EnhancedHistoryRecord[] = [];

    Object.keys(groupedByDay).forEach(date => {
        const dayRecs = groupedByDay[date];
        const partials = dayRecs.filter(r => r.resetBy === 'Salvamento automático (10:00h)');
        const mains = dayRecs.filter(r => r.resetBy !== 'Salvamento automático (10:00h)');

        if (mains.length > 0) {
            // Sort manual resets of the day by time descending (latest first)
            mains.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
            
            mains.forEach((main, idx) => {
                const item: EnhancedHistoryRecord = { ...main };
                // Associate the day's partial (10h) only with the LATEST reset of the day
                if (idx === 0 && partials.length > 0) {
                    item.hasPartial = true;
                    item.partialRecord = partials[0];
                }
                finalHistory.push(item);
            });
        } else {
            // If only partials exist, show them as standalone cards
            finalHistory.push(...partials);
        }
    });

    const sorted = finalHistory.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
    
    if (sorted.length > 0 && !selectedRecord) {
        setSelectedRecord(sorted[0]);
    }

    return sorted;
  }, [history]);

  const toggleCategory = (category: string) => {
    setCollapsedCategories(prev => 
      prev.includes(category) 
        ? prev.filter(c => c !== category) 
        : [...prev, category]
    );
  };

  const getGroupedTasks = (tasks: Task[]) => {
    return tasks.reduce((acc, task) => {
      const cat = task.category || 'Geral';
      if (!acc[cat]) acc[cat] = [];
      acc[cat].push(task);
      return acc;
    }, {} as Record<string, Task[]>);
  };

  const getAllLocations = (tasks: Task[]) => {
      if (!tasks || tasks.length === 0) return [];
      return Array.from(new Set(tasks.flatMap(t => Object.keys(t.operations))));
  };

  const currentTasksToDisplay = viewingPartial && selectedRecord?.partialRecord 
    ? selectedRecord.partialRecord.tasks 
    : selectedRecord?.tasks || [];

  return (
    <div id="history-container" className="flex flex-col md:flex-row h-full gap-4">
      {/* Sidebar - History List */}
      <div className="w-full md:w-80 bg-white dark:bg-slate-900 rounded-xl shadow-sm border border-gray-200 dark:border-slate-800 overflow-hidden flex flex-col h-[300px] md:h-full">
        <div className="p-4 bg-gray-50 dark:bg-slate-800 border-b border-gray-200 dark:border-slate-700">
          <h2 className="font-bold text-gray-700 dark:text-gray-200 flex items-center gap-2">
            <History size={20} className="text-blue-600 dark:text-blue-400"/>
            Histórico Cloud
          </h2>
          <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">{processedHistory.length} snapshots salvos</p>
        </div>
        <div className="flex-1 overflow-y-auto p-2 space-y-2 scrollbar-thin">
          {isLoading ? (
              <div className="py-10 flex flex-col items-center gap-2 text-blue-500">
                  <Loader2 className="animate-spin" size={24}/>
                  <span className="text-[10px] font-bold uppercase">Sincronizando...</span>
              </div>
          ) : processedHistory.length === 0 ? (
            <div className="text-center text-gray-400 dark:text-gray-500 py-8 text-sm">
              Nenhum registro encontrado.
            </div>
          ) : processedHistory.map(record => (
            <button
              key={record.id}
              onClick={() => {
                  setSelectedRecord(record);
                  setViewingPartial(false);
              }}
              className={`w-full text-left p-4 rounded-2xl transition-all border flex flex-col gap-3 relative group
                ${selectedRecord?.id === record.id 
                  ? 'bg-blue-50 dark:bg-blue-900/20 border-blue-200 dark:border-blue-800 shadow-sm' 
                  : 'bg-white dark:bg-slate-900 border-transparent hover:bg-gray-50 dark:hover:bg-slate-800 hover:border-gray-200 dark:hover:border-slate-700'
                }
              `}
            >
              <div className="flex items-center gap-3">
                  <div className={`p-2.5 rounded-xl ${selectedRecord?.id === record.id ? 'bg-blue-600 text-white' : 'bg-gray-100 dark:bg-slate-800 text-gray-400 dark:text-gray-500'}`}>
                    <Calendar size={18} />
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="flex flex-wrap justify-between items-center gap-x-2">
                        <div className={`font-black text-xs ${selectedRecord?.id === record.id ? 'text-blue-900 dark:text-blue-200' : 'text-gray-800 dark:text-gray-200'}`}>
                           {displayTimestamp(record.timestamp).split(',')[0]}
                        </div>
                        {record.hasPartial && (
                            <div 
                                onClick={(e) => {
                                    e.stopPropagation();
                                    setSelectedRecord(record);
                                    setViewingPartial(true);
                                }}
                                className="text-[9px] font-black text-blue-500 hover:text-blue-700 underline whitespace-nowrap cursor-pointer transition-colors animate-pulse"
                            >
                                (Ver salvamento parcial)
                            </div>
                        )}
                    </div>
                    <div className="text-[10px] text-gray-500 dark:text-gray-400 font-bold flex items-center gap-1 mt-0.5">
                       <Clock size={10} />
                       {displayTimestamp(record.timestamp).split(',')[1]}
                    </div>
                  </div>
              </div>
              <div className="flex items-center gap-2 px-3 py-1.5 bg-black/5 dark:bg-white/5 rounded-xl border dark:border-white/5">
                  <User size={12} className="text-slate-400" />
                  <span className="text-[10px] font-black text-slate-600 dark:text-slate-400 truncate uppercase">
                      {record.resetBy || 'Desconhecido'}
                  </span>
              </div>
            </button>
          ))}
        </div>
      </div>

      {/* Main Content - Read Only Table */}
      <div className="flex-1 bg-white dark:bg-slate-900 rounded-xl shadow-sm border border-gray-200 dark:border-slate-800 overflow-hidden flex flex-col">
        {!selectedRecord ? (
           <div className="flex-1 flex flex-col items-center justify-center text-gray-400 dark:text-gray-600">
              <History size={48} className="mb-4 opacity-20"/>
              <p className="font-bold uppercase text-[10px] tracking-widest">Selecione um snapshot para visualizar</p>
           </div>
        ) : (
          <>
            <div className="p-4 bg-gray-50 dark:bg-slate-800 border-b border-gray-200 dark:border-slate-700 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
               <div className="flex items-center gap-4">
                  <div className={`w-12 h-12 ${viewingPartial ? 'bg-amber-500' : 'bg-blue-600'} rounded-2xl flex items-center justify-center text-white shadow-lg transition-all transform hover:scale-105`}>
                      {viewingPartial ? <FileSearch size={24} /> : <User size={24} />}
                  </div>
                  <div>
                    <h3 className="font-black text-gray-800 dark:text-white text-sm uppercase tracking-tight flex items-center gap-2">
                        {viewingPartial ? (
                            <>Visualizando Salvamento Parcial (10:00h)</>
                        ) : (
                            <>Responsável: {selectedRecord.resetBy || 'Não informado'}</>
                        )}
                    </h3>
                    <div className="flex items-center gap-3 mt-1">
                        <span className="text-[10px] flex items-center gap-1.5 text-slate-500 font-black uppercase tracking-wider">
                            <MapPin size={10} className="text-blue-500" /> {selectedRecord.email || 'Logística CCO'}
                        </span>
                        <span className="text-[10px] flex items-center gap-1.5 text-slate-500 font-black uppercase tracking-wider">
                            <Calendar size={10} className="text-blue-500" /> {displayTimestamp(viewingPartial && selectedRecord.partialRecord ? selectedRecord.partialRecord.timestamp : selectedRecord.timestamp)}
                        </span>
                    </div>
                  </div>
               </div>
               
               <div className="flex gap-2">
                  {selectedRecord.hasPartial && (
                      <button 
                        onClick={() => setViewingPartial(!viewingPartial)}
                        className={`px-4 py-2 rounded-xl text-[10px] font-black uppercase border transition-all flex items-center gap-2 shadow-sm ${viewingPartial 
                            ? 'bg-blue-600 text-white border-blue-500 hover:bg-blue-700' 
                            : 'bg-amber-100 dark:bg-amber-900/30 text-amber-600 dark:text-amber-400 border-amber-200 dark:border-amber-800 hover:bg-amber-200 dark:hover:bg-amber-800/50'}`}
                      >
                          {viewingPartial ? (
                              <><Eye size={14}/> Voltar para Snapshot Principal</>
                          ) : (
                              <><Sparkles size={14}/> Ver Salvamento Parcial (10:00h)</>
                          )}
                      </button>
                  )}
                  <div className="px-3 py-1 bg-green-100 dark:bg-green-900/30 text-green-600 dark:text-green-400 rounded-full text-[10px] font-black uppercase border border-green-200 dark:border-green-800 flex items-center gap-1.5">
                      <CheckCircle2 size={12}/> Snapshot Recuperado
                  </div>
               </div>
            </div>

            <div className="flex-1 overflow-auto bg-slate-50 dark:bg-slate-950 scrollbar-thin">
                <table className="w-full border-collapse bg-white dark:bg-slate-900 text-[10px]">
                  <thead className="sticky top-0 z-20 bg-slate-800 dark:bg-slate-950 text-white shadow-lg">
                    <tr>
                      <th className="p-3 text-left min-w-[200px] w-[30%] border-r border-slate-700 dark:border-slate-800 sticky left-0 bg-slate-800 dark:bg-slate-950 z-30 shadow-[4px_0_8px_-4px_rgba(0,0,0,0.3)] font-black uppercase tracking-widest text-[9px]">AÇÃO / DESCRIÇÃO</th>
                      {getAllLocations(currentTasksToDisplay).map(loc => (
                        <th key={loc} className="p-1 text-center min-w-[50px] border-r border-slate-700 dark:border-slate-800 font-bold uppercase tracking-tighter text-[9px]">
                            {loc.replace('LAT-', '').replace('ITA-', '')}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {(() => {
                        const grouped = getGroupedTasks(currentTasksToDisplay);
                        const locs = getAllLocations(currentTasksToDisplay);
                        const categories = Object.keys(grouped);

                        if (categories.length === 0) {
                            return <tr><td colSpan={locs.length + 1} className="p-12 text-center text-slate-400 font-bold uppercase text-[10px] italic">Nenhum dado disponível neste snapshot.</td></tr>;
                        }

                        return categories.map(category => {
                            const categoryTasks = grouped[category];
                            const isCollapsed = collapsedCategories.includes(category);

                            return (
                                <React.Fragment key={category}>
                                    <tr 
                                        className="bg-slate-600 dark:bg-slate-800 text-white font-black uppercase tracking-widest cursor-pointer hover:bg-slate-700 dark:hover:bg-slate-700 transition-colors text-[11px] h-8"
                                        onClick={() => toggleCategory(category)}
                                    >
                                        <td colSpan={1 + locs.length} className="px-3 border-y border-slate-700 dark:border-slate-900 sticky left-0 z-10 text-[9px]">
                                            <div className="flex items-center gap-2">
                                                {isCollapsed ? <ChevronRight size={14} strokeWidth={3}/> : <ChevronDown size={14} strokeWidth={3}/>}
                                                {category}
                                            </div>
                                        </td>
                                    </tr>

                                    {!isCollapsed && categoryTasks.map(task => (
                                        <tr key={task.id} className="hover:bg-blue-50/30 dark:hover:bg-slate-800/50 transition-colors border-b border-gray-100 dark:border-slate-800 group h-12">
                                            <td className="p-3 border-r border-gray-100 dark:border-slate-800 sticky left-0 bg-white dark:bg-slate-900 group-hover:bg-blue-50/30 dark:group-hover:bg-slate-800/50 z-10 shadow-[4px_0_8px_-4px_rgba(0,0,0,0.1)]">
                                                <div className="font-bold text-gray-800 dark:text-gray-200 text-[11px] leading-tight">{task.title}</div>
                                                {task.description && (
                                                  <div className="text-gray-400 dark:text-gray-500 text-[9px] font-normal leading-snug whitespace-pre-wrap mt-1 opacity-70">{task.description}</div>
                                                )}
                                            </td>
                                            {locs.map(loc => {
                                                const statusKey = task.operations[loc] || 'PR';
                                                const config = STATUS_CONFIG[statusKey] || STATUS_CONFIG['PR'];
                                                return (
                                                    <td key={`${task.id}-${loc}`} className="p-0 border-r border-gray-100 dark:border-slate-800 h-full relative group/cell">
                                                        <div className={`absolute inset-[3px] rounded-lg flex items-center justify-center text-[9px] font-black border transition-all ${config.color} shadow-sm uppercase`}>
                                                            {config.label}
                                                        </div>
                                                    </td>
                                                );
                                            })}
                                        </tr>
                                    ))}
                                </React.Fragment>
                            );
                        });
                    })()}
                  </tbody>
                </table>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

const Sparkles = ({ size = 20, className = "" }) => (
    <svg 
        width={size} height={size} viewBox="0 0 24 24" fill="none" 
        stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" 
        className={className}
    >
        <path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z"/>
        <path d="M5 3v4"/><path d="M19 17v4"/><path d="M3 5h4"/><path d="M17 19h4"/>
    </svg>
);

export default HistoryViewer;
