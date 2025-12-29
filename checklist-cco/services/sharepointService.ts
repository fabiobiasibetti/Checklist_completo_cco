
import { SPTask, SPOperation, SPStatus, Task, OperationStatus, HistoryRecord } from '../types';

const SITE_PATH = "vialacteoscombr.sharepoint.com:/sites/CCO";
let cachedSiteId: string | null = null;
const columnMappingCache: Record<string, { mapping: Record<string, string>, readOnly: Set<string> }> = {};

async function graphFetch(endpoint: string, token: string, options: RequestInit = {}) {
  const url = endpoint.startsWith('https://') ? endpoint : `https://graph.microsoft.com/v1.0${endpoint}`;
  
  const headers: Record<string, string> = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json',
    // Formato consolidado de Prefer conforme documentação oficial para evitar erros de indexação
    'Prefer': 'HonorNonIndexedQueriesWarningMayFailOverLargeLists,HonorNonIndexedQueriesWarningMayFailRandomly'
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
    
    // Log detalhado para depuração no console do navegador
    console.warn(`Graph API Response [${res.status}]: ${errDetail}`);
    
    if (res.status === 403) {
      throw new Error("Acesso Negado (403): O SharePoint bloqueou a gravação. Verifique as permissões de EDIÇÃO no site.");
    }
    
    // Se for erro de indexação (400 Bad Request com mensagem de index), repassamos para o catch específico
    const isIndexingError = errDetail.toLowerCase().includes("indexed") || errDetail.toLowerCase().includes("filter");
    if (res.status === 400 && isIndexingError) {
        const error = new Error(errDetail);
        (error as any).isIndexingError = true;
        throw error;
    }

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
  
  columns.value.forEach((col: any) => {
    const internalName = col.name;
    const normalizedName = normalizeString(col.name);
    const normalizedDisplay = normalizeString(col.displayName);
    
    mapping[normalizedName] = internalName;
    mapping[normalizedDisplay] = internalName;

    if (col.readOnly || internalName.startsWith('_') || ['LinkTitle', 'LinkTitleNoMenu', 'ID', 'Author', 'Editor', 'Created', 'Modified', 'Attachments'].includes(internalName)) {
      readOnly.add(internalName);
    }
  });

  columnMappingCache[cacheKey] = { mapping, readOnly };
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

  async getStatusByDate(token: string, date: string): Promise<SPStatus[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
        const filter = `fields/DataReferencia eq '${date}'`;
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
        return (data.value || []).map((item: any) => ({
          id: item.id,
          DataReferencia: item.fields.DataReferencia,
          TarefaID: String(item.fields.TarefaID || item.fields.Title),
          OperacaoSigla: item.fields.OperacaoSigla,
          Status: item.fields.Status,
          Usuario: item.fields.Usuario,
          Title: item.fields.Title
        }));
    } catch (e) { return []; }
  },

  async updateStatus(token: string, status: SPStatus): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const { mapping, readOnly } = await getListColumnMapping(siteId, list.id, token);

    const rawFields: Record<string, any> = {
      Title: status.Title,
      DataReferencia: status.DataReferencia,
      TarefaID: status.TarefaID,
      OperacaoSigla: status.OperacaoSigla,
      Status: status.Status,
      Usuario: status.Usuario,
      ChaveUnica: status.Title // Espelhamos o ID único também na coluna ChaveUnica se existir
    };

    const fields: Record<string, any> = {};
    Object.keys(rawFields).forEach(key => {
      const internalName = resolveFieldName(mapping, key);
      if (!readOnly.has(internalName) || internalName === 'Title') {
        fields[internalName] = rawFields[key];
      }
    });

    try {
        // Tenta buscar item existente para atualizar (PATCH)
        // Se a lista estiver vazia ou a coluna Title não for indexada, isso PODE falhar.
        const filter = `fields/Title eq '${status.Title}'`;
        const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
        
        if (existing?.value?.length > 0) {
          const itemId = existing.value[0].id;
          await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
            method: 'PATCH',
            body: JSON.stringify(fields)
          });
        } else {
          // Se não encontrar, cria novo (POST)
          await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
            method: 'POST',
            body: JSON.stringify({ fields })
          });
        }
    } catch (e: any) {
        // Se cair aqui, pode ser erro de indexação (lista vazia/nova)
        // Tentamos o POST direto como "última esperança"
        console.warn("Busca falhou ou erro de indexação. Tentando POST direto...");
        try {
            await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
              method: 'POST',
              body: JSON.stringify({ fields })
            });
        } catch (postErr: any) {
            console.error("Falha final no POST:", postErr.message);
            throw postErr;
        }
    }
  },

  async saveHistory(token: string, record: HistoryRecord): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
    const { mapping, readOnly } = await getListColumnMapping(siteId, list.id, token);

    const rawFields: any = {
      Title: record.resetBy, 
      Data: record.timestamp,
      DadosJSON: JSON.stringify(record.tasks),
      Celula: record.email
    };

    const fields: any = {};
    Object.keys(rawFields).forEach(key => {
        const internalName = resolveFieldName(mapping, key);
        if (!readOnly.has(internalName) || internalName === 'Title') {
            fields[internalName] = rawFields[key];
        }
    });

    try {
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
          method: 'POST',
          body: JSON.stringify({ fields })
        });
    } catch (error: any) {
        throw new Error(`Erro ao gravar no Histórico: ${error.message}`);
    }
  },

  async getHistory(token: string, userEmail: string): Promise<HistoryRecord[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
      const filter = `fields/Celula eq '${userEmail}'`;
      const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
      
      return (data.value || []).map((item: any) => ({
        id: item.fields.id || item.id,
        timestamp: item.fields.Data,
        resetBy: item.fields.Title, 
        email: item.fields.Celula,
        tasks: JSON.parse(item.fields.DadosJSON || '[]')
      })).sort((a: any, b: any) => new Date(b.timestamp || 0).getTime() - new Date(a.timestamp || 0).getTime());
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
    } catch (e) {
      return [];
    }
  },

  async getAllListsMetadata(token: string) {
    const listNames = [
      'Tarefas_Checklist',
      'Operacoes_Checklist',
      'Status_Checklist',
      'Historico_checklist_web',
      'Usuarios_cco'
    ];
    
    const results = await Promise.all(listNames.map(async name => {
      try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, name, token);
        const columns = await graphFetch(`/sites/${siteId}/lists/${list.id}/columns`, token);
        return {
          list,
          columns: columns.value || []
        };
      } catch (e) {
        return { list: { displayName: name, id: 'error' }, columns: [], error: true };
      }
    }));
    return results;
  }
};
