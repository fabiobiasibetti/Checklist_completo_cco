
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
    // Retornamos o objeto de erro para tratamento específico em vez de travar tudo
    return { error: true, status: res.status, message: errText };
  }
  return res.status === 204 ? null : res.json();
}

async function getResolvedSiteId(token: string): Promise<string> {
  if (cachedSiteId) return cachedSiteId;
  const siteData = await graphFetch(`/sites/${SITE_PATH}`, token);
  if (siteData.error) throw new Error(siteData.message);
  cachedSiteId = siteData.id;
  return siteData.id;
}

async function findListByIdOrName(siteId: string, listName: string, token: string): Promise<any> {
  const data = await graphFetch(`/sites/${siteId}/lists`, token);
  if (data.error) throw new Error(data.message);
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
  if (columns.error) throw new Error(columns.message);
  
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
  const normalized = normalizeString(target);
  return mapping[normalized] || target;
}

export const SharePointService = {
  async getTasks(token: string): Promise<SPTask[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Tarefas_Checklist', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    
    const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
    if (data.error) throw new Error(data.message);
    
    const fId = resolveFieldName(mapping.mapping, 'ID');
    const fTitle = resolveFieldName(mapping.mapping, 'Title');
    const fDesc = resolveFieldName(mapping.mapping, 'Descricao');
    const fCat = resolveFieldName(mapping.mapping, 'Categoria');
    const fHor = resolveFieldName(mapping.mapping, 'Horario');
    const fAtiva = resolveFieldName(mapping.mapping, 'Ativa');
    const fOrd = resolveFieldName(mapping.mapping, 'Ordem');

    return (data.value || []).map((item: any) => ({
      id: String(item.fields[fId] || item.id),
      Title: item.fields[fTitle] || "Sem Título",
      Descricao: item.fields[fDesc] || "",
      Categoria: item.fields[fCat] || "Geral",
      Horario: item.fields[fHor] || "--:--",
      Ativa: item.fields[fAtiva] !== false,
      Ordem: Number(item.fields[fOrd]) || 999
    })).sort((a: any, b: any) => a.Ordem - b.Ordem);
  },

  async getOperations(token: string, userEmail: string): Promise<SPOperation[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Operacoes_Checklist', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    
    const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
    if (data.error) throw new Error(data.message);
    
    const fEmail = resolveFieldName(mapping.mapping, 'Email');
    const fTitle = resolveFieldName(mapping.mapping, 'Title');
    const fOrd = resolveFieldName(mapping.mapping, 'Ordem');
    const fId = resolveFieldName(mapping.mapping, 'ID');

    const filtered = (data.value || [])
      .filter((item: any) => {
        const emailValue = String(item.fields[fEmail] || "").toLowerCase().trim();
        return emailValue === userEmail.toLowerCase().trim();
      });

    return filtered.map((item: any) => ({
        id: String(item.fields[fId] || item.id),
        Title: item.fields[fTitle] || "OP",
        Ordem: Number(item.fields[fOrd]) || 0,
        Email: item.fields[fEmail] || ""
      })).sort((a: any, b: any) => a.Ordem - b.Ordem);
  },

  async getCurrentStatusMatrix(token: string, opSiglas: string[]): Promise<SPStatus[]> {
    if (opSiglas.length === 0) return [];
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    
    const fOpSigla = resolveFieldName(mapping.mapping, 'OperacaoSigla');
    const fTarefaId = resolveFieldName(mapping.mapping, 'TarefaID');
    const fStatus = resolveFieldName(mapping.mapping, 'Status');
    const fUsuario = resolveFieldName(mapping.mapping, 'Usuario');
    const fTitle = resolveFieldName(mapping.mapping, 'Title');
    const fData = resolveFieldName(mapping.mapping, 'DataReferencia');

    const filterParts = opSiglas.map(sigla => `fields/${fOpSigla} eq '${sigla}'`);
    const filter = filterParts.join(' or ');
    
    // Tenta com Filtro (Performance)
    let data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    // Se falhar por erro 400 (não indexado), tenta buscar tudo e filtrar no cliente
    if (data.error && data.status === 400) {
      console.warn("Filtro SharePoint falhou (coluna não indexada). Buscando lista completa para processamento local...");
      data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=5000`, token);
    }

    if (data.error) throw new Error(data.message);

    const items = data.value || [];
    
    // Filtro manual caso o Graph não tenha conseguido filtrar na origem
    const filteredItems = items.filter((item: any) => {
        const sigla = item.fields[fOpSigla];
        return opSiglas.includes(sigla);
    });

    return filteredItems.map((item: any) => ({
      id: item.id,
      DataReferencia: item.fields[fData],
      TarefaID: String(item.fields[fTarefaId]),
      OperacaoSigla: item.fields[fOpSigla],
      Status: item.fields[fStatus],
      Usuario: item.fields[fUsuario],
      Title: item.fields[fTitle]
    }));
  },

  async ensureStatusMatrix(token: string, tasks: SPTask[], ops: SPOperation[]): Promise<void> {
    if (tasks.length === 0 || ops.length === 0) {
      console.warn("Nada a garantir: Tarefas ou Operações vazias.");
      return;
    }

    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    
    const opSiglas = ops.map(o => o.Title);
    const existing = await this.getCurrentStatusMatrix(token, opSiglas);
    
    const existingKeys = new Set(existing.map(s => `${s.TarefaID}_${s.OperacaoSigla}`));
    const missingItems: any[] = [];

    const fTitle = resolveFieldName(mapping.mapping, 'Title');
    const fTarefaId = resolveFieldName(mapping.mapping, 'TarefaID');
    const fOpSigla = resolveFieldName(mapping.mapping, 'OperacaoSigla');
    const fStatus = resolveFieldName(mapping.mapping, 'Status');
    const fUsuario = resolveFieldName(mapping.mapping, 'Usuario');

    tasks.forEach(task => {
      opSiglas.forEach(sigla => {
        const key = `${task.id}_${sigla}`;
        if (!existingKeys.has(key)) {
          const payload: any = { fields: {} };
          payload.fields[fTitle] = `MATRIZ_${key}`;
          payload.fields[fTarefaId] = task.id;
          payload.fields[fOpSigla] = sigla;
          payload.fields[fStatus] = 'PR';
          payload.fields[fUsuario] = 'Sistema';
          missingItems.push(payload);
        }
      });
    });

    if (missingItems.length === 0) return;

    console.log(`Criando ${missingItems.length} novas células na matriz...`);
    // Criamos de 1 em 1 para garantir estabilidade, mas em paralelo limitado se necessário
    for (const item of missingItems) {
      try {
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
          method: 'POST',
          body: JSON.stringify(item)
        });
      } catch (e) {
        console.error("Erro ao criar célula da matriz:", e);
      }
    }
  },

  async updateStatus(token: string, status: SPStatus): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    
    const matrixKey = `MATRIZ_${status.TarefaID}_${status.OperacaoSigla}`;
    const fTitle = resolveFieldName(mapping.mapping, 'Title');
    const fStatus = resolveFieldName(mapping.mapping, 'Status');
    const fUsuario = resolveFieldName(mapping.mapping, 'Usuario');
    const fData = resolveFieldName(mapping.mapping, 'DataReferencia');

    const fields: any = {};
    fields[fStatus] = status.Status;
    fields[fUsuario] = status.Usuario;
    fields[fData] = new Date().toISOString();

    const filter = `fields/${fTitle} eq '${matrixKey}'`;
    let existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    // Fallback se o Title não estiver indexado
    if (existing.error && existing.status === 400) {
        existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        if (!existing.error) {
            existing.value = existing.value.filter((i: any) => i.fields[fTitle] === matrixKey);
        }
    }

    if (!existing.error && existing.value?.length > 0) {
      const itemId = existing.value[0].id;
      await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
        method: 'PATCH',
        body: JSON.stringify(fields)
      });
    } else {
        const fTarefaId = resolveFieldName(mapping.mapping, 'TarefaID');
        const fOpSigla = resolveFieldName(mapping.mapping, 'OperacaoSigla');
        const fullFields: any = { ...fields };
        fullFields[fTitle] = matrixKey;
        fullFields[fTarefaId] = status.TarefaID;
        fullFields[fOpSigla] = status.OperacaoSigla;

        await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
            method: 'POST',
            body: JSON.stringify({ fields: fullFields })
        });
    }
  },

  async saveHistory(token: string, record: HistoryRecord): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    
    const fTitle = resolveFieldName(mapping.mapping, 'Title');
    const fData = resolveFieldName(mapping.mapping, 'Data');
    const fJSON = resolveFieldName(mapping.mapping, 'DadosJSON');
    const fCelula = resolveFieldName(mapping.mapping, 'Celula');

    const fields: any = {};
    fields[fTitle] = record.resetBy || 'Reset';
    fields[fData] = new Date(record.timestamp).toISOString().split('.')[0] + 'Z';
    fields[fJSON] = JSON.stringify(record.tasks);
    fields[fEmailMapping(fCelula)] = record.email;

    await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
      method: 'POST',
      body: JSON.stringify({ fields })
    });
  },

  async getHistory(token: string, userEmail: string): Promise<HistoryRecord[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    
    const fCelula = resolveFieldName(mapping.mapping, 'Celula');
    const fTitle = resolveFieldName(mapping.mapping, 'Title');
    const fData = resolveFieldName(mapping.mapping, 'Data');
    const fJSON = resolveFieldName(mapping.mapping, 'DadosJSON');

    const filter = `fields/${fCelula} eq '${userEmail}'`;
    let data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    if (data.error && data.status === 400) {
        data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        if (!data.error) {
            data.value = data.value.filter((i: any) => String(i.fields[fCelula] || "").toLowerCase() === userEmail.toLowerCase());
        }
    }

    if (data.error) throw new Error(data.message);

    return (data.value || []).map((item: any) => ({
      id: item.id,
      timestamp: item.fields[fData],
      resetBy: item.fields[fTitle], 
      email: item.fields[fCelula],
      tasks: JSON.parse(item.fields[fJSON] || '[]')
    })).sort((a: any, b: any) => new Date(b.timestamp || 0).getTime() - new Date(a.timestamp || 0).getTime());
  },

  async getRegisteredUsers(token: string, email: string): Promise<string[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Usuarios_cco', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    
    const fEmail = resolveFieldName(mapping.mapping, 'Email');
    const fNome = resolveFieldName(mapping.mapping, 'Nome');

    const filter = `fields/${fEmail} eq '${email}'`;
    let data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    if (data.error && data.status === 400) {
        data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        if (!data.error) {
            data.value = data.value.filter((i: any) => String(i.fields[fEmail] || "").toLowerCase() === email.toLowerCase());
        }
    }

    if (data.error) throw new Error(data.message);

    return (data.value || []).map((item: any) => item.fields[fNome] || "").filter(Boolean);
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

// Helper simples para mapeamento de e-mail se necessário
function fEmailMapping(f: string) { return f; }
