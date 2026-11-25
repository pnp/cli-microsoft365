import { BaseCopilotComponent } from '@microsoft/sp-copilot-component';
import { escape } from '@microsoft/sp-lodash-subset';

import type { IHelloWorldCopilotComponentProperties } from './HelloWorldCopilotComponentProperties';

import styles from './HelloWorldCopilotComponent.module.scss';
import welcomeDark from './assets/welcome-dark.png';
import welcomeLight from './assets/welcome-light.png';

import * as strings from 'HelloWorldCopilotComponentStrings';

const EXPAND_ICON: string = '⛶';

export default class HelloWorldCopilotComponent extends BaseCopilotComponent<IHelloWorldCopilotComponentProperties> {
  protected render(): void {
    const userName: string = this.context.pageContext.user.displayName || 'there';
    const theme: string = this.hostContext.theme || strings.UnknownTheme;
    const displayMode: string = this.hostContext.displayMode || strings.UnknownTheme;
    const isDarkTheme: boolean = this.hostContext.theme === 'dark';
    const isFullscreen: boolean = this.hostContext.displayMode === 'fullscreen';

    const rootClass: string = `${styles.helloWorld} ${isDarkTheme ? styles.dark : ''}`;

    const expandHtml: string = isFullscreen
      ? ''
      : `<span id="hc-expand" class="${styles.expand}" role="button" tabindex="0" title="${strings.ExpandToFullscreenTitle}" aria-label="${strings.ExpandToFullscreenTitle}">${EXPAND_ICON}</span>`;

    this.context.domElement.innerHTML = `
      <section class="${rootClass}">
        <div class="${styles.header}">
          <p class="${styles.greeting}">${strings.WelcomeGreeting} ${escape(userName)}</p>
          ${expandHtml}
        </div>

        <img class="${styles.welcomeImage}" alt="" src="${isDarkTheme ? welcomeDark : welcomeLight}" />

        <div class="${styles.details}">
          <div class="${styles.row}">
            <span class="${styles.label}">${strings.DisplayModeLabel}</span>
            <span class="${styles.value}">${escape(displayMode)}</span>
          </div>
          <div class="${styles.row}">
            <span class="${styles.label}">${strings.ThemeLabel}</span>
            <span class="${styles.value}">${escape(theme)}</span>
          </div>
          <div class="${styles.row}">
            <span class="${styles.label}">${strings.MessageLabel}</span>
            <span class="${styles.value}">${escape(this.properties.message)}</span>
          </div>
        </div>
      </section>`;

    const expandButton: HTMLElement | null = this.context.domElement.querySelector('#hc-expand');
    if (expandButton) {
      expandButton.addEventListener('click', this._handleRequestFullscreen);
      expandButton.addEventListener('keydown', (event: KeyboardEvent) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          this._handleRequestFullscreen().catch(() => undefined);
        }
      });
    }
  }

  private _handleRequestFullscreen = async (): Promise<void> => {
    await this.requestDisplayModeAsync('fullscreen');
  };
}
