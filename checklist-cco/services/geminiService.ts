
import { GoogleGenAI, Type } from "@google/genai";
import { Task, TaskPriority, TaskStatus, RouteDeparture } from "../types";

/**
 * Service to handle AI interactions using Gemini API.
 * Following strict guidelines for initialization and model selection.
 */

export const parseExcelContentToTasks = async (rawText: string): Promise<Partial<Task>[]> => {
  // Always initialize AI instance with apiKey inside the function for text tasks
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  const prompt = `
    Analyze the following raw text which was copied from a spreadsheet (Excel/CSV).
    It represents a list of tasks for a CRM or Checklist.
    Extract the tasks into a structured JSON array.
    
    If columns are not clear, infer the meaning based on content.
    - Infer 'title' from the main activity description.
    - Infer 'description' if there are extra details.
    - Infer 'priority' (High, Medium, Low) based on urgency words. Defaults to 'Medium'.
    - Infer 'category' (e.g., Sales, Admin, Meeting) based on context.
    - Infer 'dueDate' if a date is present, format as YYYY-MM-DD. If not, leave empty.
    
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
              title: { type: Type.STRING },
              description: { type: Type.STRING },
              priority: { type: Type.STRING, enum: ["Baixa", "Média", "Alta"] },
              category: { type: Type.STRING },
              dueDate: { type: Type.STRING },
            },
            required: ["title", "priority", "category"]
          }
        }
      }
    });

    if (response.text) {
      const parsed = JSON.parse(response.text.trim());
      return parsed.map((item: any) => ({
        ...item,
        status: TaskStatus.TODO,
        createdAt: new Date().toISOString()
      }));
    }
    return [];
  } catch (error) {
    console.error("Error parsing Excel content with Gemini:", error);
    throw error;
  }
};

export const parseRouteDepartures = async (rawText: string): Promise<Partial<RouteDeparture>[]> => {
  // Always initialize AI instance with apiKey inside the function for text tasks
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
  // Always initialize AI instance with apiKey inside the function for text tasks
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
