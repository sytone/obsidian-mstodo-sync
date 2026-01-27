import type * as msalCommon from '@azure/msal-common';
import * as msal from '@azure/msal-node';
import { type AuthenticationProvider, Client, type ClientOptions } from '@microsoft/microsoft-graph-client';
import { type App, type DataAdapter, Notice } from 'obsidian';
import { MicrosoftAuthModal } from 'src/gui/microsoftAuthModal';
import { t } from 'src/lib/lang';

class MsalNodeAuthenticationProvider implements AuthenticationProvider {
    /**
     *
     */
    constructor(private readonly clientProvider: MicrosoftClientProvider) {}

    /**
     * This method will get called before every request to the msgraph server
     * This should return a Promise that resolves to an accessToken (in case of success) or rejects with error (in case of failure)
     * Basically this method will contain the implementation for getting and refreshing accessTokens
     */
    public async getAccessToken(): Promise<string> {
        return this.clientProvider.getAccessToken();
    }
}

export class MicrosoftClientProvider {
    private _clientId: string;
    private _authority: string;

    public get clientId() {
        return this._clientId;
    }

    public set clientId(value: string) {
        this._clientId = value;
    }

    public get authority() {
        return this._authority;
    }

    public set authority(value: string) {
        this._authority = value;
    }

    private readonly scopes: string[] = ['Tasks.ReadWrite', 'openid', 'profile'];
    private pca!: msal.PublicClientApplication;
    private readonly adapter: DataAdapter;
    private readonly app: App;
    private readonly cachePath: string;
    private accounts: msal.AccountInfo[] = [];

    constructor(app: App, pluginDir: string) {
        this.adapter = app.vault.adapter;
        this.cachePath = `${pluginDir}/Microsoft_cache.json`;
        this.app = app;

        this._clientId = 'a1172059-5f55-45cd-9665-8dccc98c2587';
        this._authority = 'https://login.microsoftonline.com/consumers';
    }

    public async getClient() {
        const authProvider = async (callback: (argument0: string, argument1: string) => void) => {
            const accessToken = await this.getAccessToken();
            const error = ' ';
            callback(error, accessToken);
        };

        return Client.init({
            authProvider,
        });
    }

    public async getClientWithMiddleware() {
        const clientOptions: ClientOptions = {
            authProvider: new MsalNodeAuthenticationProvider(this),
        };

        const client = Client.initWithMiddleware(clientOptions);

        return client;
    }

    public createPublicClientApplication() {
        const beforeCacheAccess = async (cacheContext: msalCommon.TokenCacheContext) => {
            if (await this.adapter.exists(this.cachePath)) {
                cacheContext.tokenCache.deserialize(await this.adapter.read(this.cachePath));
            }
        };

        const afterCacheAccess = async (cacheContext: msalCommon.TokenCacheContext) => {
            if (cacheContext.cacheHasChanged) {
                await this.adapter.write(this.cachePath, cacheContext.tokenCache.serialize());
            }
        };

        const cachePlugin = {
            beforeCacheAccess,
            afterCacheAccess,
        };
        const config = {
            auth: {
                clientId: this.clientId,
                authority: this.authority,
            },
            cache: {
                cachePlugin,
            },
        };
        this.pca = new msal.PublicClientApplication(config);
    }

    public async getAccessToken() {
        const msalCacheManager = this.pca.getTokenCache();
        if (await this.adapter.exists(this.cachePath)) {
            msalCacheManager.deserialize(await this.adapter.read(this.cachePath));
        }

        this.accounts = await msalCacheManager.getAllAccounts();
        if (this.accounts.length === 0) {
            return this.authByDevice();
        }

        return this.authByCache(this.accounts[0]);
    }

    private async authByDevice(): Promise<string> {
        const app = this.app;
        const deviceCodeRequest = {
            async deviceCodeCallback(response: msalCommon.DeviceCodeResponse) {
                const _notice = new Notice(t('Notice_DeviceCodeOnClipboard'));
                await navigator.clipboard.writeText(response.userCode);
                new MicrosoftAuthModal(app, response.userCode, response.verificationUri).open();
                console.log(t('Notice_DeviceCodeCopiedToClipboard'), response.userCode);
            },
            scopes: this.scopes,
        };
        return this.pca
            .acquireTokenByDeviceCode(deviceCodeRequest)
            .then((request) => (request === null ? 'error' : request.accessToken));
    }

    private async authByCache(account: msal.AccountInfo): Promise<string> {
        const silentRequest = {
            account,
            scopes: this.scopes,
        };
        return this.pca
            .acquireTokenSilent(silentRequest)
            .then((request) => (request === null ? 'error' : request.accessToken))
            .catch(async () => this.authByDevice());
    }
}
