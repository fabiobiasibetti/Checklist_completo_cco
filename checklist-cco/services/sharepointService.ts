
import { SPTask, SPOperation, SPStatus, Task, OperationStatus, HistoryRecord } from '../types';

const SITE_PATH = "vialacteoscombr.sharepoint.com:/sites/CCO";
let cachedSiteId: string | null = null;
const columnMappingCache: Record<string, { mapping: Record<string, string>, readOnly: Set<string>, internalNames: Set<string> }> = {};

async function graphFetch(endpoint: string, token: string, options: RequestInit = {}) {
  const url = endpoint.startsWith('https://') ? endpoint : `https://graph.microsoft.com/v1.0${endpoint}`;
  
  const headers: Record<string, string> = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json',
    'Prefer': 'HonorNonIndexedQueriesWarningMayFailOverLargeLists'
  };

  const res = await fetch(url, {
    ...options,
    headers: { ...headers, ...options.headers }
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(errText);
  }
  return res.status === 204 ? null : res.json();
}

async function getResolvedSiteId(token: string): Promise<string> {
  if (cachedSiteId) return cachedSiteId;
  const siteData = await graphFetch(`/sites/${SITE_PATH}`, token);
  cachedSiteId = siteData.id;
  return siteData.id;
}

async function findListByIdOrName(siteId: string, listName: string, token: string): Promise<any> {
  const data = await graphFetch(`/sites/${siteId}/lists`, token);
  const found = data.value.find((l: any) => 
    l.name?.toLowerCase() === listName.toLowerCase() || 
    l.displayName?.toLowerCase() === listName.toLowerCase()
  );
  if (found) return found;
  throw new Error(`Lista '${listName}' não encontrada.`);
}

function normalizeString(str: string): string {
  if (!str) return "";
  return str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "").trim();
}

async function getListColumnMapping(siteId: string, listId: string, token: string) {
  const cacheKey = `${siteId}_${listId}`;
  if (columnMappingCache[cacheKey]) return columnMappingCache[cacheKey];

  const columns = await graphFetch(`/sites/${siteId}/lists/${listId}/columns`, token);
  const mapping: Record<string, string> = {};
  const readOnly = new Set<string>();
  const internalNames = new Set<string>();
  
  columns.value.forEach((col: any) => {
    const internalName = col.name;
    mapping[normalizeString(col.name)] = internalName;
    mapping[normalizeString(col.displayName)] = internalName;
    internalNames.add(internalName);
    if (col.readOnly || ['ID', 'Author', 'Editor', 'Created', 'Modified'].includes(internalName)) {
      readOnly.add(internalName);
    }
  });

  columnMappingCache[cacheKey] = { mapping, readOnly, internalNames };
  return columnMappingCache[cacheKey];
}

function resolveFieldName(mapping: Record<string, string>, target: string): string {
  return mapping[normalizeString(target)] || target;
}

export const SharePointService = {
  async getTasks(token: string): Promise<SPTask[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Tarefas_Checklist', token);
    const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
    return (data.value || []).map((item: any) => ({
      id: String(item.fields.id || item.id),
      Title: item.fields.Title || "Sem Título",
      Descricao: item.fields.Descricao || "",
      Categoria: item.fields.Categoria || "Geral",
      Horario: item.fields.Horario || "--:--",
      Ativa: item.fields.Ativa !== false,
      Ordem: Number(item.fields.Ordem) || 999
    })).sort((a: any, b: any) => a.Ordem - b.Ordem);
  },

  async getOperations(token: string, userEmail: string): Promise<SPOperation[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Operacoes_Checklist', token);
    const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
    return (data.value || [])
      .filter((item: any) => (item.fields.Email || "").toLowerCase().trim() === userEmail.toLowerCase().trim())
      .map((item: any) => ({
        id: String(item.fields.id || item.id),
        Title: item.fields.Title || "OP",
        Ordem: Number(item.fields.Ordem) || 0,
        Email: item.fields.Email || ""
      })).sort((a: any, b: any) => a.Ordem - b.Ordem);
  },

  // Busca o estado atual da matriz para o usuário (sem filtro de data para ser persistente)
  async getCurrentStatusMatrix(token: string, opSiglas: string[]): Promise<SPStatus[]> {
    if (opSiglas.length === 0) return [];
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    
    // Filtra pelas operações do usuário para não baixar a lista inteira
    const filterParts = opSiglas.map(sigla => `fields/OperacaoSigla eq '${sigla}'`);
    const filter = filterParts.join(' or ');
    
    const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    return (data.value || []).map((item: any) => ({
      id: item.id,
      DataReferencia: item.fields.DataReferencia,
      TarefaID: String(item.fields.TarefaID),
      OperacaoSigla: item.fields.OperacaoSigla,
      Status: item.fields.Status,
      Usuario: item.fields.Usuario,
      Title: item.fields.Title
    }));
  },

  // Função CRÍTICA: Garante que existam registros para cada célula da matriz
  async ensureStatusMatrix(token: string, tasks: SPTask[], ops: SPOperation[]): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const opSiglas = ops.map(o => o.Title);
    const existing = await this.getCurrentStatusMatrix(token, opSiglas);
    
    const existingKeys = new Set(existing.map(s => `${s.TarefaID}_${s.OperacaoSigla}`));
    const missingItems: any[] = [];

    tasks.forEach(task => {
      opSiglas.forEach(sigla => {
        const key = `${task.id}_${sigla}`;
        if (!existingKeys.has(key)) {
          missingItems.push({
            fields: {
              Title: `MATRIZ_${key}`,
              TarefaID: task.id,
              OperacaoSigla: sigla,
              Status: 'PR',
              Usuario: 'Sistema'
            }
          });
        }
      });
    });

    // Cria os registros faltantes em lote (sequencial para evitar 429)
    for (const item of missingItems) {
      await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
        method: 'POST',
        body: JSON.stringify(item)
      });
    }
  },

  async updateStatus(token: string, status: SPStatus): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const matrixKey = `MATRIZ_${status.TarefaID}_${status.OperacaoSigla}`;

    const fields = {
      Status: status.Status,
      Usuario: status.Usuario,
      DataReferencia: new Date().toISOString()
    };

    // Tenta encontrar o registro da matriz pela chave única Title
    const filter = `fields/Title eq '${matrixKey}'`;
    const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    if (existing?.value?.length > 0) {
      const itemId = existing.value[0].id;
      await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
        method: 'PATCH',
        body: JSON.stringify(fields)
      });
    } else {
      // Fallback: se não existir, cria (não deveria ocorrer se o ensureStatusMatrix rodar)
      await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
        method: 'POST',
        body: JSON.stringify({ fields: { ...fields, Title: matrixKey, TarefaID: status.TarefaID, OperacaoSigla: status.OperacaoSigla } })
      });
    }
  },

  async saveHistory(token: string, record: HistoryRecord): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
    const fields = {
      Title: record.resetBy || 'Reset', 
      Data: new Date(record.timestamp).toISOString().split('.')[0] + 'Z',
      DadosJSON: JSON.stringify(record.tasks),
      Celula: record.email
    };

    await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
      method: 'POST',
      body: JSON.stringify({ fields })
    });
  },

  async getHistory(token: string, userEmail: string): Promise<HistoryRecord[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
    const filter = `fields/Celula eq '${userEmail}'`;
    const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    return (data.value || []).map((item: any) => ({
      id: item.id,
      timestamp: item.fields.Data,
      resetBy: item.fields.Title, 
      email: item.fields.Celula,
      tasks: JSON.parse(item.fields.DadosJSON || '[]')
    })).sort((a: any, b: any) => new Date(b.timestamp || 0).getTime() - new Date(a.timestamp || 0).getTime());
  },

  async getRegisteredUsers(token: string, email: string): Promise<string[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Usuarios_cco', token);
    const filter = `fields/Email eq '${email}'`;
    const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    return (data.value || []).map((item: any) => item.fields.Nome || "").filter(Boolean);
  },

  async getAllListsMetadata(token: string) {
    const listNames = ['Tarefas_Checklist', 'Operacoes_Checklist', 'Status_Checklist', 'Historico_checklist_web', 'Usuarios_cco'];
    return Promise.all(listNames.map(async name => {
      try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, name, token);
        const columns = await graphFetch(`/sites/${siteId}/lists/${list.id}/columns`, token);
        return { list, columns: columns.value || [] };
      } catch (e) {
        return { list: { displayName: name, id: 'error' }, columns: [], error: true };
      }
    }));
  }
};
