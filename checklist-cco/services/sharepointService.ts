
import { SPTask, SPOperation, SPStatus, Task, OperationStatus, HistoryRecord } from '../types';

const SITE_PATH = "vialacteoscombr.sharepoint.com:/sites/CCO";
let cachedSiteId: string | null = null;
const columnMappingCache: Record<string, { mapping: Record<string, string>, readOnly: Set<string>, internalNames: Set<string> }> = {};

async function graphFetch(endpoint: string, token: string, options: RequestInit = {}) {
  const url = endpoint.startsWith('https://') ? endpoint : `https://graph.microsoft.com/v1.0${endpoint}`;
  
  const headers: Record<string, string> = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json',
    // Adicionado header para ignorar avisos de colunas não indexadas em listas grandes
    'Prefer': 'HonorNonIndexedQueriesWarningMayFailOverLargeLists'
  };

  const res = await fetch(url, {
    ...options,
    headers: {
      ...headers,
      ...options.headers
    }
  });

  if (!res.ok) {
    let errDetail = "";
    try {
      const err = await res.json();
      errDetail = err.error?.message || JSON.stringify(err);
    } catch(e) {
      errDetail = await res.text();
    }
    console.warn(`Graph API Response [${res.status}]: ${errDetail}`);
    throw new Error(errDetail);
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
  try {
    return await graphFetch(`/sites/${siteId}/lists/${listName}`, token);
  } catch (e) {
    const data = await graphFetch(`/sites/${siteId}/lists`, token);
    const found = data.value.find((l: any) => 
      l.name?.toLowerCase() === listName.toLowerCase() || 
      l.displayName?.toLowerCase() === listName.toLowerCase()
    );
    if (found) return found;
  }
  throw new Error(`Lista '${listName}' não encontrada.`);
}

function normalizeString(str: string): string {
  if (!str) return "";
  return str.toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") 
    .replace(/[^a-z0-9]/g, "")       
    .trim();
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
    const normalizedName = normalizeString(col.name);
    const normalizedDisplay = normalizeString(col.displayName);
    
    mapping[normalizedName] = internalName;
    mapping[normalizedDisplay] = internalName;
    internalNames.add(internalName);

    if (col.readOnly || internalName.startsWith('_') || ['LinkTitle', 'ID', 'Author', 'Editor', 'Created', 'Modified'].includes(internalName)) {
      readOnly.add(internalName);
    }
  });

  columnMappingCache[cacheKey] = { mapping, readOnly, internalNames };
  return columnMappingCache[cacheKey];
}

function resolveFieldName(mapping: Record<string, string>, target: string): string {
  const normalizedTarget = normalizeString(target);
  return mapping[normalizedTarget] || target;
}

export const SharePointService = {
  async getTasks(token: string): Promise<SPTask[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Tarefas_Checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || []).map((item: any) => ({
          id: String(item.fields.id || item.id),
          Title: item.fields.Title || "Sem Título",
          Descricao: item.fields[resolveFieldName(mapping, 'Descricao')] || "",
          Categoria: item.fields[resolveFieldName(mapping, 'Categoria')] || "Geral",
          Horario: item.fields[resolveFieldName(mapping, 'Horario')] || "--:--",
          Ativa: item.fields[resolveFieldName(mapping, 'Ativa')] !== false,
          Ordem: Number(item.fields[resolveFieldName(mapping, 'Ordem')]) || 999
        })).sort((a: any, b: any) => a.Ordem - b.Ordem);
    } catch (e) { return []; }
  },

  async getOperations(token: string, userEmail: string): Promise<SPOperation[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Operacoes_Checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const colEmail = resolveFieldName(mapping, 'Email');
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || [])
          .filter((item: any) => (item.fields[colEmail] || "").toLowerCase().trim() === userEmail.toLowerCase().trim())
          .map((item: any) => ({
            id: String(item.fields.id || item.id),
            Title: item.fields.Title || "OP",
            Ordem: Number(item.fields[resolveFieldName(mapping, 'Ordem')]) || 0,
            Email: item.fields[colEmail] || ""
          })).sort((a: any, b: any) => a.Ordem - b.Ordem);
    } catch (e) { return []; }
  },

  async ensureMatrix(token: string, tasks: SPTask[], ops: SPOperation[]): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);
    
    const today = new Date().toISOString().split('T')[0];
    const todayKey = today.replace(/-/g, '');
    
    const colData = resolveFieldName(mapping, 'DataReferencia');
    // Ajuste de filtro para garantir compatibilidade com colunas de data do SharePoint
    const filter = `fields/${colData} ge '${today}T00:00:00Z' and fields/${colData} le '${today}T23:59:59Z'`;
    const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    const existingKeys = new Set((existing.value || []).map((i: any) => i.fields.Title));

    const createPromises: Promise<any>[] = [];

    for (const task of tasks) {
      if (!task.Ativa) continue;
      for (const op of ops) {
        const uniqueKey = `${todayKey}_${task.id}_${op.Title}`;
        if (!existingKeys.has(uniqueKey)) {
          const rawFields: any = {
            Title: uniqueKey,
            ChaveUnica: uniqueKey, 
            DataReferencia: today + 'T12:00:00Z',
            TarefaID: task.id,
            OperacaoSigla: op.Title,
            Status: 'PR',
            Usuario: 'Sistema'
          };

          const fields: any = {};
          Object.keys(rawFields).forEach(key => {
            const internal = resolveFieldName(mapping, key);
            if (internalNames.has(internal) && (!readOnly.has(internal) || internal === 'Title')) {
              fields[internal] = rawFields[key];
            }
          });

          createPromises.push(graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
            method: 'POST',
            body: JSON.stringify({ fields })
          }));
        }
      }
    }
    
    if (createPromises.length > 0) {
      await Promise.all(createPromises);
    }
  },

  async getStatusByDate(token: string, date: string): Promise<SPStatus[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        
        const colData = resolveFieldName(mapping, 'DataReferencia');
        const filter = `fields/${colData} ge '${date}T00:00:00Z' and fields/${colData} le '${date}T23:59:59Z'`;
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
        
        return (data.value || []).map((item: any) => ({
          id: item.id,
          DataReferencia: item.fields[colData],
          TarefaID: String(item.fields[resolveFieldName(mapping, 'TarefaID')] || ""),
          OperacaoSigla: item.fields[resolveFieldName(mapping, 'OperacaoSigla')],
          Status: item.fields[resolveFieldName(mapping, 'Status')],
          Usuario: item.fields[resolveFieldName(mapping, 'Usuario')],
          Title: item.fields.Title
        }));
    } catch (e) { return []; }
  },

  async updateStatus(token: string, status: SPStatus): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const { mapping, readOnly, internalNames } = await getListColumnMapping(siteId, list.id, token);

    const filter = `fields/Title eq '${status.Title}'`;
    const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    if (!existing.value || existing.value.length === 0) {
        const rawFields: any = {
          Title: status.Title,
          ChaveUnica: status.Title,
          DataReferencia: new Date(status.DataReferencia).toISOString(),
          TarefaID: status.TarefaID,
          OperacaoSigla: status.OperacaoSigla,
          Status: status.Status,
          Usuario: status.Usuario
        };
        const fields: any = {};
        Object.keys(rawFields).forEach(key => {
            const int = resolveFieldName(mapping, key);
            if (internalNames.has(int) && (!readOnly.has(int) || int === 'Title')) fields[int] = rawFields[key];
        });
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
        return;
    }

    const itemId = existing.value[0].id;
    const updateFields: any = {
      Status: status.Status,
      Usuario: status.Usuario
    };

    const fields: any = {};
    Object.keys(updateFields).forEach(key => {
        const int = resolveFieldName(mapping, key);
        if (internalNames.has(int) && !readOnly.has(int)) fields[int] = updateFields[key];
    });

    await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
      method: 'PATCH',
      body: JSON.stringify(fields)
    });
  },

  async saveHistory(token: string, record: HistoryRecord): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
    const { mapping, readOnly, internalNames } = await getListColumnMapping(siteId, list.id, token);

    const rawFields: any = {
      Title: record.resetBy || 'Reset', 
      Data: new Date(record.timestamp).toISOString(),
      DadosJSON: JSON.stringify(record.tasks),
      Celula: record.email
    };

    const fields: any = {};
    Object.keys(rawFields).forEach(key => {
        const internalName = resolveFieldName(mapping, key);
        if (internalNames.has(internalName) && (!readOnly.has(internalName) || internalName === 'Title')) {
            fields[internalName] = rawFields[key];
        }
    });

    await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
  },

  async getHistory(token: string, userEmail: string): Promise<HistoryRecord[]> {
    try {
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
      })).sort((a: any, b: any) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
    } catch (e) { return []; }
  },

  async getRegisteredUsers(token: string, email: string): Promise<string[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Usuarios_cco', token);
      const { mapping } = await getListColumnMapping(siteId, list.id, token);
      const colEmail = resolveFieldName(mapping, 'Email');
      const colNome = resolveFieldName(mapping, 'Nome');
      const filter = `fields/${colEmail} eq '${email}'`;
      const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
      return (data.value || []).map((item: any) => item.fields[colNome] || "").filter(Boolean);
    } catch (e) { return []; }
  },

  async getAllListsMetadata(token: string) {
    const listNames = ['Tarefas_Checklist', 'Operacoes_Checklist', 'Status_Checklist', 'Historico_checklist_web', 'Usuarios_cco'];
    return Promise.all(listNames.map(async name => {
      try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, name, token);
        const columns = await graphFetch(`/sites/${siteId}/lists/${list.id}/columns`, token);
        return { list, columns: columns.value || [] };
      } catch (e) { return { list: { displayName: name, id: 'error' }, columns: [], error: true }; }
    }));
  }
};
