
import React, { useState, useEffect } from 'react';
import { User } from '../types';
import { LogIn, Loader2, AlertCircle, ShieldCheck } from 'lucide-react';
import { setCurrentUser } from '../services/storageService';
import { PublicClientApplication, Configuration, LogLevel, InteractionRequiredAuthError } from "@azure/msal-browser";

const msalConfig: Configuration = {
    auth: {
        clientId: "c176306d-f849-4cf4-bfca-22ff214cdaad",
        authority: "https://login.microsoftonline.com/7d9754b3-dcdb-4efe-8bb7-c0e5587b86ed",
        redirectUri: window.location.origin,
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
    }
};

const msalInstance = new PublicClientApplication(msalConfig);

const MicrosoftIcon = () => (
    <svg width="20" height="20" viewBox="0 0 23 23" xmlns="http://www.w3.org/2000/svg">
        <path fill="#f35325" d="M1 1h10v10H1z"/><path fill="#81bc06" d="M12 1h10v10H12z"/><path fill="#05a6f0" d="M1 12h10v10H1z"/><path fill="#ffba08" d="M12 12h10v10H12z"/>
    </svg>
);

const Login: React.FC<{ onLogin: (user: User) => void }> = ({ onLogin }) => {
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [checkingSession, setCheckingSession] = useState(true);

  useEffect(() => {
    const checkSession = async () => {
        try {
            await msalInstance.initialize();
            
            // 1. Tenta processar retorno de redirecionamento
            const response = await msalInstance.handleRedirectPromise();
            if (response && response.account) {
                onLogin({
                    email: response.account.username,
                    name: response.account.name || response.account.username,
                    accessToken: response.accessToken
                });
                return;
            }

            // 2. Verifica se há contas já logadas no browser para login automático
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                setIsLoggingIn(true);
                try {
                    const silentRequest = {
                        scopes: ["User.Read", "Sites.ReadWrite.All"],
                        account: accounts[0]
                    };
                    const silentResponse = await msalInstance.acquireTokenSilent(silentRequest);
                    onLogin({
                        email: silentResponse.account.username,
                        name: silentResponse.account.name || silentResponse.account.username,
                        accessToken: silentResponse.accessToken
                    });
                    return; // Sucesso, sai da função
                } catch (e) {
                    if (e instanceof InteractionRequiredAuthError) {
                        console.log("Interação necessária para renovar sessão.");
                    }
                } finally {
                    setIsLoggingIn(false);
                }
            }
        } catch (e) {
            console.error("Erro ao verificar sessão MSAL:", e);
        } finally {
            setCheckingSession(false);
        }
    };
    checkSession();
  }, [onLogin]);

  const handleMicrosoftLogin = async () => {
    setIsLoggingIn(true);
    setError(null);
    try {
        const loginRequest = {
            scopes: ["User.Read", "Sites.ReadWrite.All"],
            prompt: "select_account"
        };
        const response = await msalInstance.loginPopup(loginRequest);
        if (response && response.account) {
            setCurrentUser(response.account.username);
            onLogin({
                email: response.account.username,
                name: response.account.name || response.account.username,
                accessToken: response.accessToken
            });
        }
    } catch (err: any) {
        console.error(err);
        setError("Falha na autenticação corporativa. Verifique se as janelas popup estão permitidas.");
    } finally {
        setIsLoggingIn(false);
    }
  };

  // Splash Screen no tema branco
  if (checkingSession) {
    return (
        <div className="min-h-screen bg-white flex flex-col items-center justify-center">
            <Loader2 className="text-blue-600 animate-spin mb-4" size={42} />
            <p className="text-slate-500 font-bold uppercase tracking-widest text-[10px] animate-pulse">Sincronizando Sessão...</p>
        </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6 transition-colors duration-500">
      <div className="bg-white rounded-[2.5rem] shadow-2xl w-full max-w-[440px] border border-slate-100 overflow-hidden animate-fade-in">
        <div className="h-2 w-full bg-blue-600"></div>
        <div className="p-10 flex flex-col items-center text-center">
            <div className="mb-8"><img src="https://viagroup.com.br/assets/via_group-22fac685.png" alt="VIA Group" className="max-w-[180px]"/></div>
            <h1 className="text-2xl font-black text-slate-800 mb-2 uppercase tracking-tight">Checklist CCO</h1>
            <p className="text-slate-500 text-sm mb-8">Portal de Gestão de Operações Logísticas</p>
            
            {error && (
                <div className="w-full mb-6 p-4 bg-red-50 text-red-600 text-xs rounded-2xl flex items-center gap-3 border border-red-100">
                    <AlertCircle size={20} className="shrink-0" />
                    <span className="font-bold">{error}</span>
                </div>
            )}

            <button 
                onClick={handleMicrosoftLogin}
                disabled={isLoggingIn}
                className="group relative w-full bg-slate-900 text-white py-4 rounded-2xl font-bold flex items-center justify-center gap-3 transition-all hover:bg-slate-800 hover:scale-[1.02] active:scale-95 disabled:opacity-70 shadow-lg shadow-blue-500/10"
            >
                {isLoggingIn ? <Loader2 className="animate-spin" /> : <><MicrosoftIcon /><span>Acessar via Microsoft 365</span></>}
            </button>
            
            <div className="mt-8 text-[10px] text-slate-400 font-bold uppercase tracking-widest flex items-center gap-2">
                <ShieldCheck size={12} className="text-blue-500" /> Ambiente Corporativo Seguro
            </div>
        </div>
      </div>
    </div>
  );
};

export default Login;
