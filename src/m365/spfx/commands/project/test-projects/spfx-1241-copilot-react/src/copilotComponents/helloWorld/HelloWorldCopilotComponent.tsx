import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';

import { BaseCopilotComponent } from '@microsoft/sp-copilot-component';
import type { SPCopilotDisplayMode } from '@microsoft/sp-copilot-component';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';

import HelloWorld from './components/HelloWorld';
import type { IHelloWorldProps } from './components/IHelloWorldProps';
import type { IHelloWorldCopilotComponentProperties } from './HelloWorldCopilotComponentProperties';

import * as strings from 'HelloWorldCopilotComponentStrings';

/**
 * SPFx Copilot Component that renders a React-based UI demonstrating the
 * platform's headline capabilities:
 *
 * - **Brokered SSO data calls** — Microsoft Graph (`/me`) and SharePoint REST
 *   (`/_api/web`) with zero token code. The SPFx runtime's Pairwise Broker
 *   automatically provisions tokens for `SPHttpClient` and `MSGraphClientV3`.
 *
 * - **Host context & theming** — reads `hostContext.theme` and
 *   `hostContext.displayMode` to adapt to the Copilot host environment.
 *
 * - **Bridge actions** — demonstrates `requestDisplayModeAsync`,
 *   `openLinkAsync`, `sendFollowUpMessageAsync`, and `requestSizeChangeAsync`.
 *
 * Lifecycle:
 *  1. `onInit()` — fetches user and site data (runs once before first render).
 *  2. `render()` — mounts the React tree into `this.context.domElement`.
 *     Re-invoked by the framework on host-context changes.
 *  3. `onTeardown()` — unmounts React before the host tears down the iframe.
 */
export default class HelloWorldCopilotComponent extends BaseCopilotComponent<IHelloWorldCopilotComponentProperties> {
  private _userDisplayName: string = '';
  private _siteTitle: string = '';
  private _siteUrl: string = '';
  private _root: Root | undefined;

  protected async onInit(): Promise<void> {
    this._siteUrl = this.context.pageContext.web.absoluteUrl;

    // Fetch user info from Microsoft Graph (brokered SSO — no token code needed).
    // Wrapped in try/catch so the component still renders in the workbench where
    // real services may not be available.
    try {
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const me: { displayName?: string } = await graphClient.api('/me').select('displayName').get();
      this._userDisplayName = me.displayName || 'User';
    } catch {
      this._userDisplayName = this.context.pageContext.user?.displayName || 'User';
    }

    // Fetch site info from SharePoint REST API (brokered SSO).
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this._siteUrl}/_api/web?$select=Title`,
        SPHttpClient.configurations.v1
      );
      const webInfo: { Title?: string } = await response.json();
      this._siteTitle = webInfo.Title || 'SharePoint Site';
    } catch {
      this._siteTitle = 'SharePoint Site';
    }
  }

  protected render(): void {
    const props: IHelloWorldProps = {
      message: this.properties.message,
      userDisplayName: this._userDisplayName,
      siteTitle: this._siteTitle,
      siteUrl: this._siteUrl,
      hostContext: this.hostContext,
      bridge: this.context.copilotBridge,
      onRequestDisplayMode: async (mode: SPCopilotDisplayMode) => {
        await this.requestDisplayModeAsync(mode);
      },
      onRequestSizeChange: async (width: number, height: number) => {
        await this.requestSizeChangeAsync(width, height);
      },
      targetDocument: this.context.domElement.ownerDocument,
      strings
    };

    if (!this._root) {
      this._root = createRoot(this.context.domElement);
    }

    this._root.render(React.createElement(HelloWorld, props));
  }

  protected async onTeardown(): Promise<void> {
    this._root?.unmount();
    this._root = undefined;
  }
}
