
import React, { useState, useEffect } from 'react';
import { RouteDeparture } from '../types';
import { getDepartures, saveDepartures } from '../services/storageService';
import { parseRouteDepartures } from '../services/geminiService';
import { Plus, Trash2, Download, Save, AlertTriangle, CheckCircle2, Clock, Maximize2, Minimize2, X, FileText, User, Calendar, MapPin, Upload, Sparkles, Loader2 } from 'lucide-react';

const RouteDepartureView: React.FC = () => {
  const [routes, setRoutes] = useState<RouteDeparture[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [isProcessingImport, setIsProcessingImport] = useState(false);
  const [importText, setImportText] = useState('');
  
  // Form State
  const [formData, setFormData] = useState<Partial<RouteDeparture>>({
    rota: '',
    data: new Date().toISOString().split('T')[0],
    inicio: '00:00:00',
    motorista: '',
    placa: '',
    operacao: '',
    statusOp: 'OK',
    tempo: 'OK',
    observacao: '',
  });

  useEffect(() => {
    const loaded = getDepartures();
    setRoutes(loaded);
    setIsLoading(false);
  }, []);

  const handleSave = (updated: RouteDeparture[]) => {
    setRoutes(updated);
    saveDepartures(updated);
  };

  const calculateWeekString = (dateStr: string) => {
    if (!dateStr) return '';
    const date = new Date(dateStr + 'T12:00:00');
    const monthNames = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
    const month = monthNames[date.getMonth()];
    
    const day = date.getDate();
    const firstDayOfMonth = new Date(date.getFullYear(), date.getMonth(), 1);
    const firstDayWeekday = firstDayOfMonth.getDay(); // 0=Sun, 1=Mon...
    
    // Adjust so Monday is the start of the week logic (Monday=0, Sunday=6)
    const adjustedFirstDayWeekday = firstDayWeekday === 0 ? 6 : firstDayWeekday - 1;
    
    const weekNum = Math.ceil((day + adjustedFirstDayWeekday) / 7);
    return `${month} S${weekNum}`;
  };

  const openModal = () => {
    setFormData({
      ...formData,
      id: undefined,
      data: new Date().toISOString().split('T')[0],
      rota: '',
      motorista: '',
      placa: '',
      operacao: '',
      statusOp: 'OK',
      tempo: 'OK',
      observacao: '',
    });
    setIsModalOpen(true);
  };

  const closeModal = () => {
    setIsModalOpen(false);
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const week = calculateWeekString(formData.data || '');
    
    const newRoute: RouteDeparture = {
      ...formData as RouteDeparture,
      id: Date.now().toString(),
      semana: week,
      saida: '00:00:00',
      motivo: '',
      statusGeral: 'OK',
      aviso: 'NÃO',
      createdAt: new Date().toISOString()
    };
    handleSave([...routes, newRoute]);
    closeModal();
  };

  const handleImport = async () => {
    if (!importText.trim()) return;
    setIsProcessingImport(true);
    try {
        const parsed = await parseRouteDepartures(importText);
        const newRoutes: RouteDeparture[] = parsed.map(p => ({
            ...p,
            id: Date.now().toString() + Math.random().toString(36).substr(2, 9),
            semana: p.semana || calculateWeekString(p.data || ''),
            saida: p.saida || '00:00:00',
            motivo: p.motivo || '',
            statusGeral: p.statusGeral || 'OK',
            aviso: p.aviso || 'NÃO',
            statusOp: p.statusOp || 'OK',
            tempo: p.tempo || 'OK',
            createdAt: new Date().toISOString()
        } as RouteDeparture));

        handleSave([...routes, ...newRoutes]);
        setIsImportModalOpen(false);
        setImportText('');
        alert(`${newRoutes.length} rotas importadas com sucesso!`);
    } catch (error) {
        console.error(error);
        alert("Erro ao processar os dados. Certifique-se de copiar as linhas corretamente.");
    } finally {
        setIsProcessingImport(false);
    }
  };

  const removeRow = (id: string) => {
    if (confirm('Deseja excluir esta rota?')) {
      handleSave(routes.filter(r => r.id !== id));
    }
  };

  const updateCell = (id: string, field: keyof RouteDeparture, value: string) => {
    const updated = routes.map(r => r.id === id ? { ...r, [field]: value } : r);
    handleSave(updated);
  };

  const getRowStyle = (route: RouteDeparture) => {
    if (route.statusOp === 'Atrasado') {
      return 'bg-amber-50 dark:bg-amber-900/10 border-amber-200 dark:border-amber-800/50';
    }
    return 'bg-white dark:bg-slate-900 border-gray-100 dark:border-slate-800';
  };

  const getStatusBadgeStyle = (val: string) => {
    if (val === 'Atrasado') return 'bg-red-500 text-white font-bold';
    if (val === 'OK') return 'bg-green-500 text-white font-bold';
    return 'bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300';
  };

  if (isLoading) return <div className="p-8 text-center text-blue-600">Carregando...</div>;

  return (
    <div className={`flex flex-col h-full animate-fade-in ${isFullscreen ? 'fixed inset-0 z-[60] bg-white dark:bg-slate-950 p-4' : ''}`}>
      <div className="flex justify-between items-center mb-4 shrink-0">
        <div>
          <h2 className="text-xl font-bold text-gray-800 dark:text-white flex items-center gap-2">
            <Clock className="text-blue-600" />
            Saída de Rotas
            {isFullscreen && <span className="ml-2 px-2 py-0.5 bg-blue-100 dark:bg-blue-900 text-blue-700 dark:text-blue-300 text-[10px] rounded uppercase tracking-widest">Tela Cheia</span>}
          </h2>
          <p className="text-xs text-gray-500 dark:text-gray-400">Controle de horários e motivos de atraso</p>
        </div>
        <div className="flex gap-2">
          <button 
            onClick={() => setIsImportModalOpen(true)}
            className="flex items-center gap-2 px-3 py-2 bg-emerald-50 dark:bg-emerald-900/30 text-emerald-600 dark:text-emerald-400 rounded-lg hover:bg-emerald-100 dark:hover:bg-emerald-900/50 transition-all border border-emerald-100 dark:border-emerald-800 shadow-sm"
          >
            <Upload size={18} />
            <span className="text-xs font-bold hidden sm:inline">Importar Excel</span>
          </button>
          <button 
            onClick={() => setIsFullscreen(!isFullscreen)}
            className={`flex items-center gap-2 px-3 py-2 rounded-lg transition-all border shadow-sm ${isFullscreen ? 'bg-blue-600 text-white border-blue-500' : 'bg-gray-100 dark:bg-slate-800 text-gray-600 dark:text-gray-300 border-gray-200 dark:border-slate-700 hover:bg-gray-200 dark:hover:bg-slate-700'}`}
            title={isFullscreen ? "Sair da Tela Cheia" : "Tela Cheia / Ajustar à Tela"}
          >
            {isFullscreen ? <Minimize2 size={18} /> : <Maximize2 size={18} />}
            <span className="text-xs font-bold hidden sm:inline">{isFullscreen ? 'Sair' : 'Ajustar'}</span>
          </button>
          <button 
            onClick={openModal}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-lg transition-all shadow-md active:scale-95"
          >
            <Plus size={18} />
            Nova Rota
          </button>
        </div>
      </div>

      <div className={`flex-1 overflow-auto bg-white dark:bg-slate-900 rounded-xl border border-gray-200 dark:border-slate-800 shadow-sm ${isFullscreen ? 'max-h-full overflow-x-hidden' : ''}`}>
        <table className={`w-full border-collapse text-[10px] ${isFullscreen ? 'min-w-full table-fixed' : 'min-w-[1400px]'}`}>
          <thead className="sticky top-0 z-20 bg-blue-900 text-white uppercase font-bold tracking-wider text-center">
            <tr>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[6%]' : 'w-20'}`}>Semana</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[7%]' : 'w-24'}`}>Rota</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[9%]' : 'w-28'}`}>Data</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[7%]' : 'w-24'}`}>Início</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[15%]' : ''}`}>Motorista</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[7%]' : 'w-24'}`}>Placa</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[7%]' : 'w-24'}`}>Saída</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[8%]' : 'w-32'}`}>Motivo</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[15%]' : 'min-w-[200px]'}`}>Obs</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[4%]' : 'w-16'}`}>Geral</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[4%]' : 'w-16'}`}>Av</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[12%]' : 'w-40'}`}>Operação</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[7%]' : 'w-24'}`}>Status Op</th>
              <th className={`p-2 border border-blue-800 ${isFullscreen ? 'w-[7%]' : 'w-24'}`}>Tempo</th>
              <th className={`p-2 border border-blue-800 w-10 sticky right-0 bg-blue-900`}>#</th>
            </tr>
          </thead>
          <tbody>
            {routes.length === 0 && (
              <tr>
                <td colSpan={15} className="p-10 text-center text-gray-400">
                  Nenhuma rota registrada. Clique em "Nova Rota" ou "Importar" para começar.
                </td>
              </tr>
            )}
            {routes.map((route) => (
              <tr 
                key={route.id} 
                className={`hover:bg-blue-50/50 dark:hover:bg-blue-900/10 transition-colors border-b ${getRowStyle(route)}`}
              >
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.semana} onChange={(e) => updateCell(route.id, 'semana', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center uppercase" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none font-bold text-center" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.inicio} onChange={(e) => updateCell(route.id, 'inicio', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center font-mono" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.motorista} onChange={(e) => updateCell(route.id, 'motorista', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none uppercase truncate" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.placa} onChange={(e) => updateCell(route.id, 'placa', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center font-bold uppercase" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.saida} onChange={(e) => updateCell(route.id, 'saida', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center font-mono" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <select value={route.motivo} onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} className="w-full h-full p-1 bg-transparent outline-none cursor-pointer">
                    <option value="">Nenhum</option>
                    <option value="Manutenção">Manutenção</option>
                    <option value="Mão de obra">Mão de obra</option>
                    <option value="Atraso coleta">Atraso coleta</option>
                    <option value="Atraso carregamento">Atraso carregamento</option>
                    <option value="Outros">Outros</option>
                  </select>
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.observacao} onChange={(e) => updateCell(route.id, 'observacao', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none truncate" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800 text-center">
                   <select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className="w-full h-full p-1 bg-transparent outline-none text-center font-bold">
                    <option value="OK">OK</option>
                    <option value="NOK">NOK</option>
                  </select>
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800 text-center">
                  <select value={route.aviso} onChange={(e) => updateCell(route.id, 'aviso', e.target.value)} className="w-full h-full p-1 bg-transparent outline-none text-center font-bold">
                    <option value="SIM">SIM</option>
                    <option value="NÃO">NÃO</option>
                  </select>
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.operacao} onChange={(e) => updateCell(route.id, 'operacao', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none uppercase font-bold truncate" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800 text-center">
                  <select value={route.statusOp} onChange={(e) => updateCell(route.id, 'statusOp', e.target.value)} className={`w-full h-full p-1 bg-transparent outline-none text-center font-bold ${getStatusBadgeStyle(route.statusOp)}`}>
                    <option value="OK">OK</option>
                    <option value="Atrasado">Atrasado</option>
                  </select>
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                   <input type="text" value={route.tempo} onChange={(e) => updateCell(route.id, 'tempo', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center font-mono" />
                </td>
                <td className="p-1 border border-gray-200 dark:border-slate-800 sticky right-0 bg-white dark:bg-slate-900 text-center">
                  <button onClick={() => removeRow(route.id)} className="text-gray-400 hover:text-red-500 transition-colors p-1">
                    <Trash2 size={14} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* IMPORT MODAL */}
      {isImportModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4 animate-in fade-in duration-200">
             <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden border dark:border-slate-700">
                <div className="bg-emerald-800 dark:bg-slate-800 text-white p-4 flex justify-between items-center">
                    <div className="flex items-center gap-2">
                        <div className="p-2 bg-emerald-600 rounded-lg">
                            <Upload size={20} />
                        </div>
                        <div>
                            <h3 className="font-bold text-lg">Importar do Excel</h3>
                            <p className="text-[10px] text-emerald-200">Cole as linhas da planilha abaixo para processar com IA</p>
                        </div>
                    </div>
                    <button onClick={() => setIsImportModalOpen(false)} className="hover:bg-white/10 p-1 rounded-full transition-colors">
                        <X size={24} />
                    </button>
                </div>
                
                <div className="p-6 bg-gray-50 dark:bg-slate-900">
                    <div className="mb-4">
                        <label className="block text-xs font-bold text-gray-500 dark:text-gray-400 uppercase mb-2">Cole os dados aqui:</label>
                        <textarea 
                            value={importText}
                            onChange={(e) => setImportText(e.target.value)}
                            className="w-full h-64 p-4 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-xs font-mono dark:text-white focus:ring-2 focus:ring-emerald-500 outline-none resize-none"
                            placeholder="Selecione as linhas no Excel, copie (Ctrl+C) e cole aqui (Ctrl+V)..."
                        />
                    </div>

                    <div className="flex flex-col sm:flex-row gap-3">
                        <button 
                            onClick={() => setIsImportModalOpen(false)}
                            className="flex-1 py-3 bg-gray-200 dark:bg-slate-700 text-gray-700 dark:text-gray-200 font-bold rounded-xl hover:bg-gray-300 dark:hover:bg-slate-600 transition-colors"
                        >
                            Cancelar
                        </button>
                        <button 
                            onClick={handleImport}
                            disabled={!importText.trim() || isProcessingImport}
                            className="flex-2 py-3 bg-emerald-600 text-white font-bold rounded-xl hover:bg-emerald-700 transition-all shadow-lg flex items-center justify-center gap-2 active:scale-95 disabled:opacity-50 disabled:cursor-not-allowed"
                            style={{ flex: 2 }}
                        >
                            {isProcessingImport ? (
                                <>
                                    <Loader2 size={20} className="animate-spin" />
                                    Processando com IA...
                                </>
                            ) : (
                                <>
                                    <Sparkles size={20} />
                                    Processar e Importar
                                </>
                            )}
                        </button>
                    </div>
                </div>
             </div>
        </div>
      )}

      {/* ADD ROUTE MODAL */}
      {isModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4 animate-in fade-in duration-200">
          <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden border dark:border-slate-700">
            <div className="bg-blue-900 dark:bg-slate-800 text-white p-4 flex justify-between items-center">
                <div className="flex items-center gap-2">
                    <div className="p-2 bg-blue-600 rounded-lg">
                        <Plus size={20} />
                    </div>
                    <div>
                        <h3 className="font-bold text-lg">Nova Operação</h3>
                        <p className="text-[10px] text-blue-200">Insira os dados da rota abaixo</p>
                    </div>
                </div>
                <button onClick={closeModal} className="hover:bg-white/10 p-1 rounded-full transition-colors">
                    <X size={24} />
                </button>
            </div>
            
            <form onSubmit={handleSubmit} className="p-6 bg-gray-50 dark:bg-slate-900">
                <div className="space-y-4">
                    <div className="grid grid-cols-2 gap-4">
                        <div>
                            <label className="block text-[10px] font-bold text-gray-500 dark:text-gray-400 uppercase mb-1">Data</label>
                            <input 
                                type="date" required
                                value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})}
                                className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm dark:text-white focus:ring-2 focus:ring-blue-500 outline-none"
                            />
                        </div>
                        <div>
                            <label className="block text-[10px] font-bold text-gray-500 dark:text-gray-400 uppercase mb-1">Rota</label>
                            <input 
                                type="text" required
                                value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value.toUpperCase()})}
                                className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm font-bold dark:text-white focus:ring-2 focus:ring-blue-500 outline-none" placeholder="Ex: 24133D"
                            />
                        </div>
                    </div>

                    <div>
                        <label className="block text-[10px] font-bold text-gray-500 dark:text-gray-400 uppercase mb-1">Operação (Cliente)</label>
                        <input 
                            type="text" required
                            value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value.toUpperCase()})}
                            className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm font-bold dark:text-white focus:ring-2 focus:ring-blue-500 outline-none" placeholder="NOME DO CLIENTE"
                        />
                    </div>

                    <div className="grid grid-cols-2 gap-4">
                        <div>
                            <label className="block text-[10px] font-bold text-gray-500 dark:text-gray-400 uppercase mb-1">Motorista</label>
                            <input 
                                type="text" required
                                value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value.toUpperCase()})}
                                className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm dark:text-white focus:ring-2 focus:ring-blue-500 outline-none" placeholder="NOME"
                            />
                        </div>
                        <div>
                            <label className="block text-[10px] font-bold text-gray-500 dark:text-gray-400 uppercase mb-1">Placa</label>
                            <input 
                                type="text" required
                                value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value.toUpperCase()})}
                                className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm font-bold dark:text-white focus:ring-2 focus:ring-blue-500 outline-none" placeholder="ABC1D23"
                            />
                        </div>
                    </div>

                    <div className="grid grid-cols-2 gap-4">
                        <div>
                            <label className="block text-[10px] font-bold text-gray-500 dark:text-gray-400 uppercase mb-1">H. Início (Previsto)</label>
                            <input 
                                type="text" required
                                value={formData.inicio} onChange={e => setFormData({...formData, inicio: e.target.value})}
                                className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm font-mono dark:text-white focus:ring-2 focus:ring-blue-500 outline-none" placeholder="00:00:00"
                            />
                        </div>
                        <div>
                            <label className="block text-[10px] font-bold text-gray-500 dark:text-gray-400 uppercase mb-1">Tempo / Gap</label>
                            <input 
                                type="text"
                                value={formData.tempo} onChange={e => setFormData({...formData, tempo: e.target.value.toUpperCase()})}
                                className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm font-mono dark:text-white focus:ring-2 focus:ring-blue-500 outline-none" placeholder="Ex: OK ou 00:30"
                            />
                        </div>
                    </div>
                </div>

                <div className="flex gap-4 mt-8 pt-6 border-t dark:border-slate-800">
                    <button 
                        type="button" onClick={closeModal}
                        className="flex-1 py-3 bg-gray-200 dark:bg-slate-700 text-gray-700 dark:text-gray-200 font-bold rounded-xl hover:bg-gray-300 dark:hover:bg-slate-600 transition-colors"
                    >
                        Cancelar
                    </button>
                    <button 
                        type="submit"
                        className="flex-2 py-3 bg-blue-600 text-white font-bold rounded-xl hover:bg-blue-700 transition-all shadow-lg flex items-center justify-center gap-2 active:scale-95"
                    >
                        <Save size={20} />
                        Confirmar
                    </button>
                </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default RouteDepartureView;
