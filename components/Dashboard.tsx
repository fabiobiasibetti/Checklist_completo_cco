import React, { useMemo } from 'react';
import { Task } from '../types';
import { ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip, CartesianGrid, Legend, Cell } from 'recharts';
import { CheckCircle2, Clock, AlertTriangle, Activity } from 'lucide-react';

interface DashboardProps {
  tasks: Task[];
}

const Dashboard: React.FC<DashboardProps> = ({ tasks }) => {
  
  // Calculate aggregate stats across all cells (Tasks * Locations)
  const stats = useMemo(() => {
    let totalCells = 0;
    let ok = 0;
    let pending = 0;
    let issues = 0; // EA, AR, ATT, AT

    tasks.forEach(task => {
        Object.values(task.operations).forEach(status => {
            totalCells++;
            if (status === 'OK') ok++;
            else if (status === 'PR' || !status) pending++;
            else issues++;
        });
    });

    return { totalCells, ok, pending, issues };
  }, [tasks]);

  // Prepare data for the Bar Chart (Status per Location)
  const chartData = useMemo(() => {
    // 1. Identify all unique locations present in the tasks
    const allLocations: string[] = Array.from(new Set(tasks.flatMap(t => Object.keys(t.operations))));
    
    // 2. Build stats per location
    const data = allLocations.map(loc => {
        let locOk = 0;
        let locPr = 0;
        let locIssue = 0;

        tasks.forEach(t => {
            const s = t.operations[loc];
            if (s === 'OK') locOk++;
            else if (s === 'PR' || !s) locPr++;
            else locIssue++;
        });

        return {
            name: loc.replace('LAT-', '').replace('ITA-', ''), // Shorten name for chart
            fullKey: loc,
            OK: locOk,
            Pendentes: locPr,
            Problemas: locIssue // EA, AR, ATT, AT
        };
    });

    // Sort by "Problemas" descending to highlight attention points, then by alphabetical
    return data.sort((a, b) => b.Problemas - a.Problemas || a.name.localeCompare(b.name));
  }, [tasks]);

  const progress = stats.totalCells > 0 ? Math.round((stats.ok / stats.totalCells) * 100) : 0;

  return (
    <div className="space-y-6 animate-fade-in pb-10">
      {/* Stats Cards */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <div className="bg-white dark:bg-slate-900 p-6 rounded-xl shadow-sm border border-gray-100 dark:border-slate-800 flex items-center space-x-4">
          <div className="p-3 bg-blue-100 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 rounded-full">
            <Activity size={24} />
          </div>
          <div>
            <p className="text-sm text-gray-500 dark:text-gray-400 font-medium">Progresso Global</p>
            <h3 className="text-2xl font-bold text-gray-800 dark:text-white">{progress}%</h3>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-900 p-6 rounded-xl shadow-sm border border-gray-100 dark:border-slate-800 flex items-center space-x-4">
          <div className="p-3 bg-green-100 dark:bg-green-900/30 text-green-600 dark:text-green-400 rounded-full">
            <CheckCircle2 size={24} />
          </div>
          <div>
            <p className="text-sm text-gray-500 dark:text-gray-400 font-medium">Finalizados (OK)</p>
            <h3 className="text-2xl font-bold text-gray-800 dark:text-white">{stats.ok}</h3>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-900 p-6 rounded-xl shadow-sm border border-gray-100 dark:border-slate-800 flex items-center space-x-4">
          <div className="p-3 bg-red-100 dark:bg-red-900/30 text-red-600 dark:text-red-400 rounded-full">
            <AlertTriangle size={24} />
          </div>
          <div>
            <p className="text-sm text-gray-500 dark:text-gray-400 font-medium">Atenção/Erros</p>
            <h3 className="text-2xl font-bold text-gray-800 dark:text-white">{stats.issues}</h3>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-900 p-6 rounded-xl shadow-sm border border-gray-100 dark:border-slate-800 flex items-center space-x-4">
          <div className="p-3 bg-gray-100 dark:bg-slate-800 text-gray-600 dark:text-gray-400 rounded-full">
            <Clock size={24} />
          </div>
          <div>
            <p className="text-sm text-gray-500 dark:text-gray-400 font-medium">Pendentes (PR)</p>
            <h3 className="text-2xl font-bold text-gray-800 dark:text-white">{stats.pending}</h3>
          </div>
        </div>
      </div>

      {/* Chart Section */}
      <div className="bg-white dark:bg-slate-900 p-6 rounded-xl shadow-sm border border-gray-100 dark:border-slate-800 flex flex-col h-[500px]">
        <div className="flex items-center justify-between mb-6">
            <h3 className="text-lg font-bold text-gray-800 dark:text-white">Status por Operação</h3>
            <span className="text-sm text-gray-500 dark:text-gray-400">Volume de tarefas por status</span>
        </div>
        
        <div className="flex-1 w-full min-h-0">
            {chartData.length > 0 ? (
                <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                        data={chartData}
                        margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
                        layout="horizontal"
                    >
                        <CartesianGrid strokeDasharray="3 3" vertical={false} strokeOpacity={0.2} />
                        <XAxis 
                            dataKey="name" 
                            tick={{fontSize: 12, fill: '#94a3b8'}} 
                            interval={0} 
                            angle={-45} 
                            textAnchor="end" 
                            height={60}
                            stroke="#cbd5e1"
                        />
                        <YAxis stroke="#cbd5e1" tick={{fill: '#94a3b8'}} />
                        <Tooltip 
                            contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 12px rgba(0,0,0,0.1)', backgroundColor: '#1e293b', color: '#fff' }}
                            cursor={{fill: 'rgba(255,255,255,0.05)'}}
                            itemStyle={{color: '#fff'}}
                        />
                        <Legend wrapperStyle={{ paddingTop: '20px' }}/>
                        <Bar dataKey="OK" stackId="a" fill="#86efac" name="Concluído (OK)" radius={[0, 0, 4, 4]} />
                        <Bar dataKey="Problemas" stackId="a" fill="#f87171" name="Atenção/Atrasado" />
                        <Bar dataKey="Pendentes" stackId="a" fill="#e2e8f0" name="Pendente (PR)" radius={[4, 4, 0, 0]} />
                    </BarChart>
                </ResponsiveContainer>
            ) : (
                <div className="h-full flex items-center justify-center text-gray-400 dark:text-gray-600">
                    <p>Nenhum dado disponível para exibir.</p>
                </div>
            )}
        </div>
      </div>
    </div>
  );
};

export default Dashboard;