
import React, { useState, useEffect } from 'react';
import { User } from '../types';
import { LogIn, Loader2, AlertCircle, ShieldCheck } from 'lucide-react';
import { setCurrentUser } from '../services/storageService';
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { msalInstance } from '../services/authService';

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
    const initAuth = async () => {
        try {
            await msalInstance.initialize();
            
            // Verifica se o usuário saiu manualmente anteriormente
            const isManualLogout = localStorage.getItem('msal_manual_logout') === 'true';

            const response = await msalInstance.handleRedirectPromise();
            if (response && response.account) {
                // Se voltamos de um redirect de login, limpamos o flag de logout
                localStorage.removeItem('msal_manual_logout');
                onLogin({
                    email: response.account.username,
                    name: response.account.name || response.account.username,
                    accessToken: response.accessToken
                });
                return;
            }

            // Só tenta o login silencioso se o usuário não tiver clicado em "Sair" recentemente
            if (!isManualLogout) {
                const accounts = msalInstance.getAllAccounts();
                if (accounts.length > 0) {
                    try {
                        const silentRequest = {
                            scopes: ["User.Read", "Sites.ReadWrite.All"],
                            account: accounts[0]
                        };
                        const silentResponse = await msalInstance.acquireTokenSilent(silentRequest);
                        if (silentResponse) {
                            onLogin({
                                email: silentResponse.account.username,
                                name: silentResponse.account.name || silentResponse.account.username,
                                accessToken: silentResponse.accessToken
                            });
                            return;
                        }
                    } catch (silentError) {
                        if (silentError instanceof InteractionRequiredAuthError) {
                            console.warn("Sessão expirada ou requer interação.");
                        }
                    }
                }
            } else {
                console.log("Logout manual ativo: ignorando login silencioso.");
            }
        } catch (e) {
            console.error("Erro na inicialização do MSAL:", e);
        } finally {
            setCheckingSession(false);
        }
    };

    initAuth();
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
            // Sucesso no login: removemos o flag de logout manual
            localStorage.removeItem('msal_manual_logout');
            setCurrentUser(response.account.username);
            onLogin({
                email: response.account.username,
                name: response.account.name || response.account.username,
                accessToken: response.accessToken
            });
        }
    } catch (err: any) {
        console.error(err);
        setError("Falha na autenticação corporativa. Verifique sua conexão.");
    } finally {
        setIsLoggingIn(false);
    }
  };

  if (checkingSession) {
    return (
      <div className="min-h-screen bg-white flex flex-col items-center justify-center p-6">
        <Loader2 className="text-blue-600 animate-spin mb-4" size={40} />
        <p className="text-slate-400 font-bold uppercase tracking-widest text-[10px] animate-pulse">Sincronizando Sessão...</p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6 transition-colors">
      <div className="bg-white rounded-[2.5rem] shadow-2xl w-full max-w-[440px] border border-slate-100 overflow-hidden animate-fade-in">
        <div className="h-2 w-full bg-blue-600"></div>
        <div className="p-10 flex flex-col items-center">
            <div className="mb-8"><img src="https://viagroup.com.br/assets/via_group-22fac685.png" alt="VIA Group" className="max-w-[180px]"/></div>
            <h1 className="text-2xl font-black text-slate-800 mb-2 uppercase tracking-tight">Checklist CCO</h1>
            <p className="text-slate-500 text-sm mb-8">Gestão de Operações em Tempo Real</p>
            
            {error && (
              <div className="w-full mb-6 p-4 bg-red-50 text-red-600 text-xs rounded-2xl flex items-center gap-3 border border-red-100">
                <AlertCircle size={20} className="shrink-0" />
                <span className="font-bold">{error}</span>
              </div>
            )}

            <button 
                onClick={handleMicrosoftLogin}
                disabled={isLoggingIn}
                className="w-full bg-slate-900 text-white py-4 rounded-2xl font-bold flex items-center justify-center gap-3 transition-all hover:bg-slate-800 hover:scale-[1.02] active:scale-95 disabled:opacity-70 shadow-lg"
            >
                {isLoggingIn ? <Loader2 className="animate-spin" /> : <><MicrosoftIcon /><span>Entrar com Microsoft</span></>}
            </button>
            
            <div className="mt-8 text-[10px] text-slate-400 font-bold uppercase tracking-widest flex items-center gap-2">
                <ShieldCheck size={12} className="text-blue-500" /> Acesso Corporativo Seguro
            </div>
        </div>
      </div>
    </div>
  );
};

export default Login;
