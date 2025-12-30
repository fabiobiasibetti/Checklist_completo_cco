
import { GoogleGenAI, Type } from "@google/genai";
import { Task, TaskPriority, TaskStatus, RouteDeparture, SPTask } from "../types";

/**
 * Service to handle AI interactions using Gemini API.
 * Following strict guidelines for initialization and model selection.
 */

export const parseExcelContentToTasks = async (rawText: string): Promise<Partial<SPTask>[]> => {
  // Always initialize AI instance with apiKey inside the function for text tasks
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  const prompt = `
    Analyze the following raw text which was copied from a spreadsheet (Excel/CSV).
    It represents a list of operational tasks for a logistics checklist (Checklist CCO).
    Extract the tasks into a structured JSON array suitable for a SharePoint list.
    
    Mapping:
    - Title: The main action or name of the task.
    - Descricao: Detailed instructions or description.
    - Categoria: The grouping category (e.g., "ORGANIZAÇÃO PARA O PRÓXIMO DIA", "ACOMPANHAMENTO DIÁRIO").
    - Horario: Expected time or range (e.g., "22:00h - 00:00h", "10:00h").
    - Ativa: Boolean, default true.
    - Ordem: Integer representing the sort sequence.
    
    Raw Text:
    """
    ${rawText}
    """
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              Title: { type: Type.STRING },
              Descricao: { type: Type.STRING },
              Categoria: { type: Type.STRING },
              Horario: { type: Type.STRING },
              Ativa: { type: Type.BOOLEAN },
              Ordem: { type: Type.INTEGER },
            },
            required: ["Title", "Categoria"]
          }
        }
      }
    });

    if (response.text) {
      return JSON.parse(response.text.trim());
    }
    return [];
  } catch (error) {
    console.error("Error parsing Checklist tasks with Gemini:", error);
    throw error;
  }
};

export const parseRouteDepartures = async (rawText: string): Promise<Partial<RouteDeparture>[]> => {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  const prompt = `
    Analyze the raw text below copied from a logistics spreadsheet.
    Extract the data into a structured JSON array of RouteDeparture objects.
    
    Map the columns carefully based on these headers:
    1. SEMANA: Week (e.g. DEZ S2)
    2. ROTA: Route number (e.g. 24139D)
    3. DATA: Date (e.g. 13/12/2025). Format as YYYY-MM-DD in output.
    4. INÍCIO: Scheduled start time (e.g. 01:00:00)
    5. MOTORISTA: Driver name
    6. PLACA: Vehicle plate
    7. SAÍDA: Actual departure time
    8. MOTIVO: Reason for delay
    9. OBSERVAÇÃO: Comments
    10. STATUS GERAL: OK or NOK
    11. AVISO: SIM or NÃO
    12. OPERAÇÃO: Client or Location (e.g. LAT-UNA)
    13. STATUS OP: OK or Atrasado
    14. TEMPO: Gap/Time info (e.g. OK or 00:31:00)

    If a value is missing, use empty string or "OK" for status as appropriate.
    
    Raw Text:
    """
    ${rawText}
    """
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              semana: { type: Type.STRING },
              rota: { type: Type.STRING },
              data: { type: Type.STRING, description: "YYYY-MM-DD format" },
              inicio: { type: Type.STRING },
              motorista: { type: Type.STRING },
              placa: { type: Type.STRING },
              saida: { type: Type.STRING },
              motivo: { type: Type.STRING },
              observacao: { type: Type.STRING },
              statusGeral: { type: Type.STRING },
              aviso: { type: Type.STRING },
              operacao: { type: Type.STRING },
              statusOp: { type: Type.STRING },
              tempo: { type: Type.STRING }
            },
            required: ["rota", "data", "operacao"]
          }
        }
      }
    });

    if (response.text) {
      return JSON.parse(response.text.trim());
    }
    return [];
  } catch (error) {
    console.error("Error parsing departures with Gemini:", error);
    throw error;
  }
};

export const suggestTasksFromGoal = async (goal: string): Promise<Partial<Task>[]> => {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

  const prompt = `
    I have a high-level goal for my CRM/Business: "${goal}".
    Break this down into 3 to 5 actionable specific tasks for a checklist.
    Return JSON.
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              title: { type: Type.STRING },
              description: { type: Type.STRING },
              priority: { type: Type.STRING, enum: ["Baixa", "Média", "Alta"] },
              category: { type: Type.STRING },
            },
            required: ["title", "description", "priority", "category"]
          }
        }
      }
    });

    if (response.text) {
        return JSON.parse(response.text.trim());
    }
    return [];
  } catch (error) {
    console.error("Error generating tasks from goal:", error);
    return [];
  }
};
