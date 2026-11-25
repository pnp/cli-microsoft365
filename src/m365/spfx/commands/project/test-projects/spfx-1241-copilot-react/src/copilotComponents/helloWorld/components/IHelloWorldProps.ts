import type {
  ICopilotComponentHostContext,
  ISPCopilotBridge,
  SPCopilotDisplayMode
} from '@microsoft/sp-copilot-component';

export interface IHelloWorldStrings {
  ExpandButtonLabel: string;
  OpenSiteButtonLabel: string;
  FollowUpButtonLabel: string;
  ResizeButtonLabel: string;
  CompactButtonLabel: string;
  SiteBadgePrefix: string;
  ThemeBadgePrefix: string;
  ModeBadgePrefix: string;
  GreetingPrefix: string;
  UnknownTheme: string;
  DefaultDisplayMode: string;
  FollowUpMessage: string;
}

export interface IHelloWorldProps {
  /** The message passed as a tool argument from the Copilot host. */
  message: string;
  /** User display name fetched from Microsoft Graph /me. */
  userDisplayName: string;
  /** Site title fetched from SharePoint REST /_api/web. */
  siteTitle: string;
  /** Absolute URL of the current SharePoint site. */
  siteUrl: string;
  /** Host context (theme, display mode) from the Copilot host. */
  hostContext: ICopilotComponentHostContext;
  /** Bridge to communicate with the Copilot host (public API surface). */
  bridge: ISPCopilotBridge;
  /** Request the host to change display mode (e.g. 'fullscreen'). */
  onRequestDisplayMode: (mode: SPCopilotDisplayMode) => Promise<void>;
  /** Request the host to resize the component iframe. */
  onRequestSizeChange: (width: number, height: number) => Promise<void>;
  /**
   * Document the FluentProvider should inject its theme styles into. Pass
   * `domElement.ownerDocument` so Griffel writes CSS into the correct iframe
   * document rather than the top-level page.
   */
  targetDocument: Document | undefined;
  /** Localized strings for UI labels. */
  strings: IHelloWorldStrings;
}
