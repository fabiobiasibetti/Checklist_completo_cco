
import { PublicClientApplication, Configuration } from "@azure/msal-browser";

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

export const msalInstance = new PublicClientApplication(msalConfig);

export const logout = async () => {
    try {
        const accounts = msalInstance.getAllAccounts();
        
        // Limpa o estado local para evitar que o app use sessões antigas
        localStorage.removeItem('crm_active_user_session');
        // Define um flag indicando que o usuário saiu deliberadamente
        localStorage.setItem('msal_manual_logout', 'true');

        if (accounts.length > 0) {
            // logoutRedirect não abre janelas extras, ele usa a própria aba para o processo.
            // Passar o 'account' faz com que o Microsoft saiba exatamente quem deslogar.
            await msalInstance.logoutRedirect({
                account: accounts[0],
                postLogoutRedirectUri: window.location.origin,
            });
        } else {
            localStorage.clear();
            window.location.reload();
        }
    } catch (e) {
        console.error("Erro durante o logout:", e);
        localStorage.clear();
        window.location.reload();
    }
};
