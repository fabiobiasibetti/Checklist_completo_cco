
import { Task, OperationStatus, INITIAL_LOCATIONS, LOGISTICA_2_LOCATIONS, Customer, HistoryRecord, VALID_USERS, RouteDeparture } from '../types';

// Base Keys
const BASE_TASKS_KEY = 'crm_logistics_tasks_v3';
const BASE_LOCATIONS_KEY = 'crm_logistics_locations_v1';
const BASE_CUSTOMERS_KEY = 'crm_logistics_customers_v1';
const BASE_HISTORY_KEY = 'crm_logistics_history_v1';
const BASE_DEPARTURES_KEY = 'crm_logistics_departures_v1';
const SESSION_USER_KEY = 'crm_active_user_session';

// Helper to get current user from localStorage
export const getCurrentUser = () => {
  return localStorage.getItem(SESSION_USER_KEY);
};

export const setCurrentUser = (email: string | null) => {
  if (email) {
    localStorage.setItem(SESSION_USER_KEY, email);
  } else {
    localStorage.removeItem(SESSION_USER_KEY);
  }
};

// Helper to determine the key based on the user
const getKey = (baseKey: string) => {
  const user = getCurrentUser();
  if (!user) return `guest_${baseKey}`;
  
  // Sanitize email to be safe for keys
  const safeEmail = user.replace(/[^a-zA-Z0-9]/g, '_');
  return `${safeEmail}_${baseKey}`;
};

export const getLocations = (): string[] => {
  const stored = localStorage.getItem(getKey(BASE_LOCATIONS_KEY));
  if (!stored) {
    const user = getCurrentUser();
    let initial = INITIAL_LOCATIONS;

    // Specific locations for Logistica 2
    if (user === 'cco.logistica2@viagroup.com.br' || user === 'cco.logistica2viagroup.com.br') {
        initial = LOGISTICA_2_LOCATIONS;
    }

    localStorage.setItem(getKey(BASE_LOCATIONS_KEY), JSON.stringify(initial));
    return initial;
  }
  return JSON.parse(stored);
};

export const saveLocations = (locations: string[]) => {
  localStorage.setItem(getKey(BASE_LOCATIONS_KEY), JSON.stringify(locations));
};

const initOps = (locations: string[], status: OperationStatus = 'PR'): Record<string, OperationStatus> => {
  const ops: Record<string, OperationStatus> = {};
  locations.forEach(loc => ops[loc] = status);
  return ops;
};

const getInitialTasks = (): Task[] => {
    const locs = getLocations();
    return [
      {
        id: '1',
        timeRange: '22:00h - 00:00h',
        title: 'Acompanhamento de rotas que ainda não saíram',
        description: 'Preenchimento da planilha com início das rotas, motivos de atraso e data inicio manual',
        category: 'ORGANIZAÇÃO PARA O PRÓXIMO DIA',
        operations: initOps(locs),
        createdAt: new Date().toISOString(),
        isDaily: true
      },
      {
        id: '2',
        timeRange: '22:00h - 00:00h',
        title: 'Atualização do resumo de saídas',
        description: 'Assim que todas as rotas do dia saírem, enviar atualização do resumo de saídas',
        category: 'ORGANIZAÇÃO PARA O PRÓXIMO DIA',
        operations: initOps(locs),
        createdAt: new Date().toISOString(),
        isDaily: true
      },
      {
        id: '3',
        timeRange: '22:00h - 00:00h',
        title: 'Enviar saídas do dia anterior para o histórico',
        description: 'Verificar todas as informações da planilha, enviar para o histórico de saída',
        category: 'ORGANIZAÇÃO PARA O PRÓXIMO DIA',
        operations: initOps(locs),
        createdAt: new Date().toISOString(),
        isDaily: true
      },
      {
        id: '4',
        timeRange: '22:00h - 00:00h',
        title: 'Organizar escala na planilha de saídas',
        description: 'Preencher o checklist com as informações das escalas compartilhadas',
        category: 'ORGANIZAÇÃO PARA O PRÓXIMO DIA',
        operations: initOps(locs),
        createdAt: new Date().toISOString(),
        isDaily: true
      },
      {
        id: '5',
        timeRange: '00:00h - 01:00h',
        title: 'Comparativo de rotas',
        description: 'Após montar a planilha do checklist, realizar uma comparação da quantidade e quais rotas estão rodando pelo KMM, escala da filial e aplicativos de terceiros',
        category: 'ORGANIZAÇÃO PARA O PRÓXIMO DIA',
        operations: initOps(locs),
        createdAt: new Date().toISOString(),
        isDaily: true
      },
      {
        id: '6',
        timeRange: '01:00h - 03:00h',
        title: 'Análise de não coletas do dia anterior (APÓS VALIDAÇÃO SMART)',
        description: 'Verificação de não coletas do dia anterior pelo sistema e apps de terceiros.',
        category: 'ANÁLISE E VERIFICAÇÕES DE ROTAS DO DIA ANTERIOR',
        operations: initOps(locs),
        createdAt: new Date().toISOString(),
        isDaily: true
      },
      {
        id: '7',
        timeRange: '01:00h - 06:00h',
        title: 'Realizar envio de não coletas no grupo Report',
        description: 'Enviar informações no Report e aguardar retorno na parte da manhã',
        category: 'ANÁLISE E VERIFICAÇÕES DE ROTAS DO DIA ANTERIOR',
        operations: initOps(locs),
        createdAt: new Date().toISOString(),
        isDaily: true
      },
      {
        id: '8',
        timeRange: '01:00h - 12:00h',
        title: 'Acompanhamento das saídas',
        description: 'Realizar acompanhamento das saídas, preenchimento das planilhas e cobrança quando houver atrasos ou adiantamentos',
        category: 'ACOMPANHAMENTO DE SAÍDAS E NÃO COLETAS DO DIA ANTERIOR',
        operations: initOps(locs),
        createdAt: new Date().toISOString(),
        isDaily: true
      },
      {
        id: '14',
        timeRange: '10:00h',
        title: 'Troca de turno',
        description: 'Troca de informações entre os dois turnos',
        category: 'ACOMPANHAMENTO DE SAÍDAS E NÃO COLETAS DO DIA ANTERIOR',
        operations: initOps(locs),
        createdAt: new Date().toISOString(),
        isDaily: true
      },
      {
        id: '17',
        timeRange: '12:00h - 18:30h',
        title: 'Análise de não coletas',
        description: 'Verificação de não coletas pelo sistema e apps de terceiros',
        category: 'ACOMPANHAMENTO DIÁRIO',
        operations: initOps(locs),
        createdAt: new Date().toISOString(),
        isDaily: true
      }
    ];
};

export const getTasks = (): Task[] => {
  const stored = localStorage.getItem(getKey(BASE_TASKS_KEY));
  if (!stored) {
    const initial = getInitialTasks();
    localStorage.setItem(getKey(BASE_TASKS_KEY), JSON.stringify(initial));
    return initial;
  }
  return JSON.parse(stored);
};

export const saveTasks = (tasks: Task[]) => {
  localStorage.setItem(getKey(BASE_TASKS_KEY), JSON.stringify(tasks));
};

export const getCustomers = (): Customer[] => {
  const stored = localStorage.getItem(getKey(BASE_CUSTOMERS_KEY));
  if (!stored) return [];
  return JSON.parse(stored);
};

export const saveCustomers = (customers: Customer[]) => {
  localStorage.setItem(getKey(BASE_CUSTOMERS_KEY), JSON.stringify(customers));
};

export const getHistory = (): HistoryRecord[] => {
  const stored = localStorage.getItem(getKey(BASE_HISTORY_KEY));
  if (!stored) return [];
  return JSON.parse(stored);
};

export const saveHistory = (history: HistoryRecord[]) => {
  localStorage.setItem(getKey(BASE_HISTORY_KEY), JSON.stringify(history));
};

export const addToHistory = (tasks: Task[], resetBy: string) => {
  const history = getHistory();
  const newRecord: HistoryRecord = {
    id: Date.now().toString(),
    timestamp: new Date().toISOString(),
    tasks: tasks,
    resetBy: resetBy || 'Desconhecido'
  };
  const updatedHistory = [newRecord, ...history];
  saveHistory(updatedHistory);
};

export const getDepartures = (): RouteDeparture[] => {
  const stored = localStorage.getItem(getKey(BASE_DEPARTURES_KEY));
  if (!stored) return [];
  return JSON.parse(stored);
};

export const saveDepartures = (departures: RouteDeparture[]) => {
  localStorage.setItem(getKey(BASE_DEPARTURES_KEY), JSON.stringify(departures));
};
