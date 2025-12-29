
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
        if (accounts.length > 0) {
            await msalInstance.logoutPopup({
                account: accounts[0],
                postLogoutRedirectUri: window.location.origin,
            });
        }
    } catch (e) {
        console.error("Erro durante o logout:", e);
    } finally {
        // Limpa tudo localmente por garantia
        localStorage.clear();
        window.location.reload();
    }
};
