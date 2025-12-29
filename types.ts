
export type OperationStatus = 'PR' | 'OK' | 'EA' | 'AR' | 'ATT' | 'AT';

export interface OperationStates {
  [location: string]: OperationStatus;
}

export interface Task {
  id: string;
  timeRange: string;
  title: string;
  description: string;
  category: string;
  operations: OperationStates;
  createdAt: string;
  isDaily: boolean;
  active?: boolean;
}

export enum TaskPriority {
  BAIXA = 'Baixa',
  MEDIA = 'Média',
  ALTA = 'Alta'
}

export enum TaskStatus {
  TODO = 'TODO',
  DONE = 'DONE'
}

export interface Customer {
  id: string;
  name: string;
  company?: string;
  email?: string;
  phone?: string;
  status: 'Lead' | 'Ativo' | 'Inativo';
}

export interface RouteDeparture {
  id: string;
  semana: string;
  rota: string;
  data: string;
  inicio: string;
  motorista: string;
  placa: string;
  saida: string;
  motivo: string;
  observacao: string;
  statusGeral: string;
  aviso: string;
  operacao: string;
  statusOp: string;
  tempo: string;
  createdAt: string;
}

export interface SPTask {
  id: string;
  Title: string;
  Descricao: string;
  Categoria: string;
  Horario: string;
  Ativa: boolean;
  Ordem: number;
}

export interface SPOperation {
  id: string;
  Title: string;
  Ordem: number;
  Email: string;
}

export interface SPStatus {
  id?: string;
  DataReferencia: string;
  TarefaID: string;
  OperacaoSigla: string;
  Status: OperationStatus;
  Usuario: string;
  Title: string;
}

export interface User {
  email: string;
  name: string;
  accessToken?: string;
}

export interface HistoryRecord {
  id: string;
  timestamp: string;
  tasks: Task[];
  resetBy?: string;
  email?: string; // Novo campo para rastreabilidade e filtro
}

export const VALID_USERS = [
  { email: 'cco.logistica@viagroup.com.br', password: '1234', name: 'Logística 1' },
  { email: 'cco.logistica2@viagroup.com.br', password: '1234', name: 'Logística 2' }
];

export const INITIAL_LOCATIONS: string[] = [];
export const LOCATIONS: string[] = [];
export const LOGISTICA_2_LOCATIONS: string[] = ['LAT-CWB', 'LAT-SJP', 'LAT-LDB', 'LAT-MGA'];
