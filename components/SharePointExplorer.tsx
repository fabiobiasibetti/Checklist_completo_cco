
import React, { useState, useEffect } from 'react';
import { SharePointService } from '../services/sharepointService';
import { User } from '../types';
import { Database, Search, Table, Columns, AlertCircle, CheckCircle2, Loader2, Link } from 'lucide-react';

interface SharePointExplorerProps {
  currentUser: User;
}

const SharePointExplorer: React.FC<SharePointExplorerProps> = ({ currentUser }) => {
  const [metadata, setMetadata] = useState<any[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [selectedList, setSelectedList] = useState<any>(null);

  useEffect(() => {
    const fetchMetadata = async () => {
      const token = currentUser.accessToken || (window as any).__access_token;
      if (!token) {
        setIsLoading(false);
        return;
      }

      setIsLoading(true);
      try {
        const data = await SharePointService.getAllListsMetadata(token);
        setMetadata(data);
        if (data.length > 0) {
            const firstValid = data.find(d => !d.error);
            setSelectedList(firstValid || data[0]);
        }
      } catch (err) {
        console.error("Erro ao explorar SharePoint:", err);
      } finally {
        setIsLoading(false);
      }
    };
    fetchMetadata();
  }, [currentUser]);

  if (isLoading) {
    return (
      <div className="h-full flex flex-col items-center justify-center text-blue-600 gap-4">
        <Loader2 size={40} className="animate-spin" />
        <p className="font-bold animate-pulse uppercase tracking-widest text-sm">Explorando Estrutura SharePoint...</p>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-full gap-4 animate-fade-in">
      <div className="bg-white dark:bg-slate-900 p-4 rounded-xl shadow-sm border border-gray-100 dark:border-slate-800 flex justify-between items-center">
        <div>
          <h2 className="text-xl font-bold text-gray-800 dark:text-white flex items-center gap-2">
            <Search className="text-blue-600" />
            Explorador de Dados
          </h2>
          <p className="text-xs text-gray-500 dark:text-gray-400">Verifique os nomes internos das colunas no SharePoint</p>
        </div>
        <div className="flex items-center gap-2 text-[10px] font-black uppercase text-slate-400">
           Ambiente: <span className="text-blue-500">PRODUÇÃO CCO</span>
        </div>
      </div>

      <div className="flex flex-col lg:flex-row gap-4 flex-1 min-h-0">
        {/* Sidebar - List Picker */}
        <div className="w-full lg:w-72 bg-white dark:bg-slate-900 rounded-xl border dark:border-slate-800 overflow-hidden flex flex-col shrink-0">
          <div className="p-3 bg-gray-50 dark:bg-slate-800 border-b dark:border-slate-700 font-bold text-[10px] uppercase text-slate-500 tracking-widest">
            Listas Identificadas
          </div>
          <div className="p-2 space-y-1 overflow-y-auto flex-1">
            {metadata.map((item, idx) => (
              <button
                key={idx}
                onClick={() => setSelectedList(item)}
                className={`w-full text-left p-3 rounded-xl transition-all flex items-center gap-3 border
                  ${selectedList?.list.displayName === item.list.displayName 
                    ? 'bg-blue-600 text-white border-blue-500 shadow-md scale-[1.02]' 
                    : 'bg-white dark:bg-slate-900 dark:text-slate-300 border-transparent hover:bg-gray-50 dark:hover:bg-slate-800'
                  }
                `}
              >
                <Table size={18} className={selectedList?.list.displayName === item.list.displayName ? 'text-white' : 'text-blue-500'} />
                <div className="flex-1 min-w-0">
                    <div className="font-bold text-xs truncate">{item.list.displayName}</div>
                    <div className={`text-[9px] uppercase font-black opacity-60 ${item.error ? 'text-red-400' : ''}`}>
                        {item.error ? 'Erro de Conexão' : `${item.columns.length} Colunas`}
                    </div>
                </div>
                {item.error ? <AlertCircle size={14} className="text-red-400" /> : <CheckCircle2 size={14} className={selectedList?.list.displayName === item.list.displayName ? 'text-white' : 'text-green-500'} />}
              </button>
            ))}
          </div>
        </div>

        {/* Column Viewer */}
        <div className="flex-1 bg-white dark:bg-slate-900 rounded-xl border dark:border-slate-800 overflow-hidden flex flex-col shadow-sm">
          {!selectedList ? (
            <div className="h-full flex items-center justify-center text-slate-400">Selecione uma lista</div>
          ) : selectedList.error ? (
            <div className="h-full flex flex-col items-center justify-center p-8 text-center text-red-500 bg-red-50 dark:bg-red-900/10">
                <AlertCircle size={48} className="mb-4" />
                <h3 className="text-lg font-bold">Acesso Negado ou Lista Inexistente</h3>
                <p className="text-sm max-w-md mt-2 opacity-80">
                   Não foi possível encontrar a lista <b>{selectedList.list.displayName}</b>. 
                   Certifique-se de que a lista existe no site SharePoint e que seu usuário tem permissões de leitura.
                </p>
            </div>
          ) : (
            <>
              <div className="p-4 bg-gray-50 dark:bg-slate-800 border-b dark:border-slate-700 flex justify-between items-center">
                 <div className="flex items-center gap-3">
                    <div className="p-2 bg-blue-100 dark:bg-blue-900/40 rounded-lg text-blue-600">
                        <Columns size={20} />
                    </div>
                    <div>
                        <h3 className="font-bold text-gray-800 dark:text-white uppercase tracking-tight">{selectedList.list.displayName}</h3>
                        <p className="text-[10px] text-slate-500 font-mono">ID: {selectedList.list.id}</p>
                    </div>
                 </div>
                 <a 
                   href={selectedList.list.webUrl} 
                   target="_blank" 
                   rel="noopener noreferrer"
                   className="flex items-center gap-2 px-3 py-1.5 bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-lg text-xs font-bold hover:bg-slate-300 transition-colors"
                 >
                    <Link size={14} /> Ver no SharePoint
                 </a>
              </div>

              <div className="flex-1 overflow-auto bg-slate-50 dark:bg-slate-950 p-4">
                 <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-3">
                    {selectedList.columns.map((col: any) => (
                      <div 
                        key={col.id} 
                        className="bg-white dark:bg-slate-900 p-3 rounded-xl border dark:border-slate-800 shadow-sm hover:border-blue-300 dark:hover:border-blue-800 transition-colors group"
                      >
                        <div className="flex justify-between items-start mb-1">
                            <span className="text-[10px] font-black uppercase text-blue-600 dark:text-blue-400 bg-blue-50 dark:bg-blue-900/30 px-2 py-0.5 rounded">
                                {col.name === 'Title' ? 'Primária' : col.readOnly ? 'Sistema' : 'Dados'}
                            </span>
                            {col.required && <span className="text-[9px] font-bold text-red-500 uppercase">Obrigatória</span>}
                        </div>
                        <div className="font-bold text-gray-800 dark:text-white text-xs mb-1 truncate">{col.displayName}</div>
                        <div className="text-[10px] font-mono text-slate-500 bg-slate-100 dark:bg-slate-800 p-1.5 rounded select-all group-hover:bg-blue-50 dark:group-hover:bg-blue-900/20 transition-colors">
                           InternalName: <span className="font-bold text-slate-800 dark:text-slate-200">{col.name}</span>
                        </div>
                        <div className="mt-2 text-[9px] text-slate-400 italic">
                            Tipo: {col.text ? 'Texto' : col.dateTime ? 'Data/Hora' : col.number ? 'Número' : col.boolean ? 'Booleano' : col.choice ? 'Escolha' : 'Personalizado'}
                        </div>
                      </div>
                    ))}
                 </div>
              </div>
              
              <div className="p-3 bg-white dark:bg-slate-900 border-t dark:border-slate-800 text-center">
                  <p className="text-[10px] text-slate-400 font-medium">
                      <span className="text-amber-500 font-black">DICA:</span> Use o <b>InternalName</b> em scripts e conexões para maior confiabilidade.
                  </p>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
};

export default SharePointExplorer;
