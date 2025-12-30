
import React, { useState, useEffect } from 'react';
import { HashRouter as Router, Routes, Route, useNavigate } from 'react-router-dom';
import { CheckSquare, History, Truck, Moon, Sun, LogOut, ChevronLeft, ChevronRight, Loader2, Search } from 'lucide-react';
import TaskManager from './components/TaskManager';
import HistoryViewer from './components/HistoryViewer';
import RouteDepartureView from './components/RouteDeparture';
import SharePointExplorer from './components/SharePointExplorer';
import Login from './components/Login';
import { SharePointService } from './services/sharepointService';
import { logout as msalLogout } from './services/authService';
import { Task, User } from './types';
import { setCurrentUser as setStorageUser } from './services/storageService';

const SidebarLink = ({ to, icon: Icon, label, active, collapsed }: any) => (
  <a href={`#${to}`} className={`flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${active ? 'bg-blue-600 text-white' : 'text-slate-500 hover:bg-slate-100'} ${collapsed ? 'justify-center' : ''}`}>
    <Icon size={20} />
    {!collapsed && <span className="font-medium whitespace-nowrap">{label}</span>}
  </a>
);

const AppContent = () => {
  const [currentUser, setUser] = useState<User | null>(null);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [locations, setLocations] = useState<string[]>([]);
  const [teamMembers, setTeamMembers] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [syncMessage, setSyncMessage] = useState("Iniciando...");
  const [isDarkMode, setIsDarkMode] = useState(true);
  const [collapsed, setCollapsed] = useState(true);
  const [collapsedCategories, setCollapsedCategories] = useState<string[]>([]);
  
  const navigate = useNavigate();

  const loadDataFromSharePoint = async (user: User) => {
    if (!user.accessToken) return;
    (window as any).__access_token = user.accessToken; 
    setIsLoading(true);
    try {
      setSyncMessage("Carregando Definições...");
      const spTasks = await SharePointService.getTasks(user.accessToken);
      const spOps = await SharePointService.getOperations(user.accessToken, user.email);
      const spMembers = await SharePointService.getTeamMembers(user.accessToken);
      setTeamMembers(spMembers);
      
      setSyncMessage("Sincronizando Matriz 1:1...");
      await SharePointService.ensureMatrix(user.accessToken, spTasks, spOps);

      setSyncMessage("Recuperando Status...");
      const today = new Date().toISOString().split('T')[0];
      const spStatus = await SharePointService.getStatusByDate(user.accessToken, today);

      const opSiglas = spOps.map(o => o.Title);
      setLocations(opSiglas);

      const matrixTasks: Task[] = spTasks.map(t => {
        const ops: Record<string, any> = {};
        opSiglas.forEach(sigla => {
          const statusMatch = spStatus.find(s => s.TarefaID === t.id && s.OperacaoSigla === sigla);
          ops[sigla] = statusMatch ? statusMatch.Status : 'PR';
        });

        return {
          id: t.id,
          title: t.Title,
          description: t.Descricao,
          category: t.Categoria,
          timeRange: t.Horario,
          operations: ops,
          createdAt: new Date().toISOString(),
          isDaily: true,
          active: t.Ativa
        };
      });

      setTasks(matrixTasks.filter(t => t.active !== false));
    } catch (err) {
      console.error("Erro ao carregar SharePoint:", err);
      alert("Erro crítico na sincronização: " + (err as Error).message);
    } finally {
      setIsLoading(false);
    }
  };

  // Logic for automatic save after 10:00h - IMPROVED
  useEffect(() => {
    if (!currentUser || tasks.length === 0) return;

    const checkAutoSaveTrigger = async () => {
      const now = new Date();
      const hours = now.getHours();

      // Trigger if it's 10:00 AM or later
      if (hours >= 10) {
        // Usar data local para evitar problemas de fuso horário/UTC
        const localDateStr = now.toLocaleDateString('pt-BR').replace(/\//g, '-');
        const safeEmail = currentUser.email.replace(/[^a-zA-Z0-9]/g, '_');
        const autoSaveFlag = `auto_save_done_${safeEmail}_${localDateStr}`;
        
        if (localStorage.getItem(autoSaveFlag) !== 'true') {
          console.log(`[${now.toLocaleTimeString()}] Iniciando backup automático das 10:00h...`);
          try {
            await SharePointService.saveHistory(currentUser.accessToken!, {
              id: Date.now().toString(),
              timestamp: now.toISOString(),
              tasks: tasks,
              resetBy: 'Salvamento automático (10:00h)',
              email: currentUser.email
            });
            localStorage.setItem(autoSaveFlag, 'true');
            console.log("Backup automático 10:00h concluído com sucesso.");
          } catch (e) {
            console.error("Falha no backup automático:", e);
          }
        }
      }
    };

    // Check every 60 seconds
    const interval = setInterval(checkAutoSaveTrigger, 60000);
    checkAutoSaveTrigger();
    return () => clearInterval(interval);
  }, [currentUser, tasks]);

  const handleLogout = async () => {
    await msalLogout();
    setUser(null);
    setStorageUser(null);
    delete (window as any).__access_token;
    navigate('/');
  };

  useEffect(() => {
    if (isDarkMode) document.documentElement.classList.add('dark');
    else document.documentElement.classList.remove('dark');
  }, [isDarkMode]);

  if (!currentUser) return <Login onLogin={(u) => { setUser(u); loadDataFromSharePoint(u); }} />;

  return (
    <div className="flex h-screen bg-slate-50 dark:bg-slate-950 overflow-hidden">
      <aside className={`bg-white dark:bg-slate-900 border-r dark:border-slate-800 transition-all ${collapsed ? 'w-20' : 'w-64'} p-4 flex flex-col`}>
        <div className="mb-10 flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white font-bold">V</div>
          {!collapsed && <h1 className="font-bold dark:text-white text-sm">CCO Digital</h1>}
        </div>
        <nav className="flex-1 space-y-2">
          <SidebarLink to="/" icon={CheckSquare} label="Checklist" active={window.location.hash === '#/'} collapsed={collapsed} />
          <SidebarLink to="/departures" icon={Truck} label="Saídas" active={window.location.hash === '#/departures'} collapsed={collapsed} />
          <SidebarLink to="/history" icon={History} label="Histórico" active={window.location.hash === '#/history'} collapsed={collapsed} />
          <SidebarLink to="/explorer" icon={Search} label="Explorador" active={window.location.hash === '#/explorer'} collapsed={collapsed} />
        </nav>
        <div className="mt-auto space-y-2 border-t pt-4 dark:border-slate-800">
           <button onClick={() => setIsDarkMode(!isDarkMode)} className="p-2 w-full flex justify-center text-slate-500 hover:bg-slate-100 rounded-lg">
             {isDarkMode ? <Sun size={20}/> : <Moon size={20}/>}
           </button>
           <button onClick={() => setCollapsed(!collapsed)} className="p-2 w-full flex justify-center text-slate-500 hover:bg-slate-100 rounded-lg">
             {collapsed ? <ChevronRight size={20}/> : <ChevronLeft size={20}/>}
           </button>
        </div>
      </aside>
      <main className="flex-1 overflow-hidden p-4">
        {isLoading ? (
          <div className="h-full flex items-center justify-center flex-col gap-4 text-blue-600">
             <Loader2 size={40} className="animate-spin" />
             <p className="font-bold animate-pulse text-sm uppercase tracking-widest">{syncMessage}</p>
          </div>
        ) : (
          <Routes>
            <Route path="/" element={
              <TaskManager 
                tasks={tasks} 
                setTasks={setTasks} 
                locations={locations} 
                setLocations={setLocations} 
                onUserSwitch={() => loadDataFromSharePoint(currentUser)} 
                collapsedCategories={collapsedCategories} 
                setCollapsedCategories={setCollapsedCategories} 
                currentUser={currentUser}
                onLogout={handleLogout}
                teamMembers={teamMembers}
              />
            } />
            <Route path="/departures" element={<RouteDepartureView />} />
            <Route path="/history" element={<HistoryViewer currentUser={currentUser} />} />
            <Route path="/explorer" element={<SharePointExplorer currentUser={currentUser} />} />
          </Routes>
        )}
      </main>
    </div>
  );
};

const App = () => (<Router><AppContent /></Router>);
export default App;
