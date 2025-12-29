
import { HistoryRecord, Task, User as AppUser } from '../types';
import { SharePointService } from '../services/sharepointService';
import { History, Calendar, Clock, ChevronRight, ChevronDown, CheckCircle2, User, Loader2, MapPin } from 'lucide-react';
import React, { useState, useEffect } from 'react';

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

const HistoryViewer: React.FC<HistoryViewerProps> = ({ currentUser }) => {
  const [history, setHistory] = useState<HistoryRecord[]>([]);
  const [selectedRecord, setSelectedRecord] = useState<HistoryRecord | null>(null);
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
            if (data.length > 0) setSelectedRecord(data[0]);
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
        if (timestamp.includes('/')) return timestamp; 
        const date = new Date(timestamp);
        return date.toLocaleString('pt-BR');
    } catch(e) {
        return timestamp;
    }
  };

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
      return Array.from(new Set(tasks.flatMap(t => Object.keys(t.operations))));
  };

  const fontSize = 'text-[10px]';
  const categorySize = 'text-[11px] py-1';
  const actionColWidth = 'min-w-[200px] w-[30%]';

  return (
    <div id="history-container" className="flex flex-col md:flex-row h-full gap-4">
      {/* Sidebar - History List */}
      <div className="w-full md:w-80 bg-white dark:bg-slate-900 rounded-xl shadow-sm border border-gray-200 dark:border-slate-800 overflow-hidden flex flex-col h-[300px] md:h-full">
        <div className="p-4 bg-gray-50 dark:bg-slate-800 border-b border-gray-200 dark:border-slate-700">
          <h2 className="font-bold text-gray-700 dark:text-gray-200 flex items-center gap-2">
            <History size={20} className="text-blue-600 dark:text-blue-400"/>
            Histórico Cloud
          </h2>
          <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">{history.length} snapshots salvos</p>
        </div>
        <div className="flex-1 overflow-y-auto p-2 space-y-2 scrollbar-thin">
          {isLoading ? (
              <div className="py-10 flex flex-col items-center gap-2 text-blue-500">
                  <Loader2 className="animate-spin" size={24}/>
                  <span className="text-[10px] font-bold uppercase">Sincronizando...</span>
              </div>
          ) : history.length === 0 ? (
            <div className="text-center text-gray-400 dark:text-gray-500 py-8 text-sm">
              Nenhum registro encontrado.
            </div>
          ) : history.map(record => (
            <button
              key={record.id}
              onClick={() => setSelectedRecord(record)}
              className={`w-full text-left p-3 rounded-xl transition-all border flex flex-col gap-2
                ${selectedRecord?.id === record.id 
                  ? 'bg-blue-50 dark:bg-blue-900/20 border-blue-200 dark:border-blue-800 shadow-sm' 
                  : 'bg-white dark:bg-slate-900 border-transparent hover:bg-gray-50 dark:hover:bg-slate-800 hover:border-gray-200 dark:hover:border-slate-700'
                }
              `}
            >
              <div className="flex items-center gap-3">
                  <div className={`p-2 rounded-lg ${selectedRecord?.id === record.id ? 'bg-blue-600 text-white' : 'bg-gray-100 dark:bg-slate-800 text-gray-400 dark:text-gray-500'}`}>
                    <Calendar size={16} />
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className={`font-bold text-xs ${selectedRecord?.id === record.id ? 'text-blue-900 dark:text-blue-200' : 'text-gray-700 dark:text-gray-300'}`}>
                       {displayTimestamp(record.timestamp).split(' ')[0]}
                    </div>
                    <div className="text-[10px] text-gray-500 dark:text-gray-400 flex items-center gap-1">
                       <Clock size={10} />
                       {displayTimestamp(record.timestamp).split(' ')[1]}
                    </div>
                  </div>
              </div>
              <div className="flex items-center gap-2 px-2 py-1 bg-black/5 dark:bg-white/5 rounded-lg">
                  <User size={12} className="text-slate-400" />
                  <span className="text-[10px] font-bold text-slate-600 dark:text-slate-400 truncate">
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
              <p className="font-bold">Selecione um snapshot para visualizar</p>
           </div>
        ) : (
          <>
            <div className="p-4 bg-gray-50 dark:bg-slate-800 border-b border-gray-200 dark:border-slate-700 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
               <div className="flex items-center gap-4">
                  <div className="w-12 h-12 bg-blue-600 rounded-2xl flex items-center justify-center text-white shadow-lg shadow-blue-500/20">
                      <User size={24} />
                  </div>
                  <div>
                    <h3 className="font-black text-gray-800 dark:text-white text-sm uppercase tracking-tight">
                        Responsável: {selectedRecord.resetBy || 'Não informado'}
                    </h3>
                    <div className="flex items-center gap-3 mt-1">
                        <span className="text-[10px] flex items-center gap-1 text-slate-500 font-bold uppercase">
                            <MapPin size={10} /> {selectedRecord.email || 'CCO'}
                        </span>
                        <span className="text-[10px] flex items-center gap-1 text-slate-500 font-bold uppercase">
                            <Calendar size={10} /> {displayTimestamp(selectedRecord.timestamp)}
                        </span>
                    </div>
                  </div>
               </div>
               <div className="px-3 py-1 bg-green-100 dark:bg-green-900/30 text-green-600 dark:text-green-400 rounded-full text-[10px] font-black uppercase border border-green-200 dark:border-green-800">
                   Snapshot Recuperado
               </div>
            </div>

            <div className="flex-1 overflow-auto bg-slate-50 dark:bg-slate-950 scrollbar-thin">
                <table className={`w-full border-collapse bg-white dark:bg-slate-900 ${fontSize}`}>
                  <thead className="sticky top-0 z-20 bg-slate-800 dark:bg-slate-950 text-white shadow-lg">
                    <tr>
                      <th className={`p-2 text-left ${actionColWidth} border-r border-slate-700 dark:border-slate-800 sticky left-0 bg-slate-800 dark:bg-slate-950 z-30 shadow-[4px_0_8px_-4px_rgba(0,0,0,0.3)] font-black uppercase tracking-widest text-[9px]`}>AÇÃO / TAREFA</th>
                      {getAllLocations(selectedRecord.tasks).map(loc => (
                        <th key={loc} className={`p-1 text-center min-w-[45px] border-r border-slate-700 dark:border-slate-800 font-bold ${fontSize}`}>
                            {loc.replace('LAT-', '').replace('ITA-', '')}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {(() => {
                        const grouped = getGroupedTasks(selectedRecord.tasks);
                        const locs = getAllLocations(selectedRecord.tasks);
                        const categories = Object.keys(grouped);

                        return categories.map(category => {
                            const categoryTasks = grouped[category];
                            const isCollapsed = collapsedCategories.includes(category);

                            return (
                                <React.Fragment key={category}>
                                    <tr 
                                        className={`bg-slate-600 dark:bg-slate-800 text-white font-bold uppercase tracking-widest cursor-pointer hover:bg-slate-700 dark:hover:bg-slate-700 transition-colors ${categorySize}`}
                                        onClick={() => toggleCategory(category)}
                                    >
                                        <td colSpan={1 + locs.length} className={`px-3 border-y border-slate-700 dark:border-slate-900 sticky left-0 z-10 text-[9px]`}>
                                            <div className="flex items-center gap-2">
                                                {isCollapsed ? <ChevronRight size={14}/> : <ChevronDown size={14}/>}
                                                {category}
                                            </div>
                                        </td>
                                    </tr>

                                    {!isCollapsed && categoryTasks.map(task => (
                                        <tr key={task.id} className="hover:bg-blue-50/30 dark:hover:bg-slate-800/50 transition-colors border-b border-gray-100 dark:border-slate-800 group">
                                            <td className={`p-3 border-r border-gray-100 dark:border-slate-800 sticky left-0 bg-white dark:bg-slate-900 group-hover:bg-blue-50/30 dark:group-hover:bg-slate-800/50 z-10 shadow-[4px_0_8px_-4px_rgba(0,0,0,0.1)]`}>
                                                <div className={`font-bold text-gray-800 dark:text-gray-200 text-[11px] leading-tight`}>{task.title}</div>
                                                {task.description && (
                                                  <div className="text-gray-400 dark:text-gray-500 text-[9px] font-normal leading-snug whitespace-pre-wrap mt-1 opacity-80">{task.description}</div>
                                                )}
                                            </td>
                                            {locs.map(loc => {
                                                const statusKey = task.operations[loc] || 'PR';
                                                const config = STATUS_CONFIG[statusKey] || STATUS_CONFIG['PR'];
                                                return (
                                                    <td key={`${task.id}-${loc}`} className="p-0 border-r border-gray-100 dark:border-slate-800 h-full relative">
                                                        <div className={`absolute inset-[2px] rounded flex items-center justify-center text-[8px] font-black border ${config.color} shadow-sm uppercase`}>
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

export default HistoryViewer;
