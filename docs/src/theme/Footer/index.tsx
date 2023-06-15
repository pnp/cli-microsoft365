import React from 'react';
import clsx from 'clsx';
import styles from './styles.module.scss';
import GitHubSVG from '@site/static/img/github-icon.svg';
import DiscordSVG from '@site/static/img/discord-icon.svg';
import TwitterSVG from '@site/static/img/twitter-icon.svg';
import YouTubeSVG from '@site/static/img/youtube-icon.svg';
import LinkSVG from '@site/static/img/link-icon.svg';
import { MendableFloatingButton } from '@mendable/search';
import { useColorMode } from '@docusaurus/theme-common';
import useDocusaurusContext from '@docusaurus/useDocusaurusContext';

function Footer(): JSX.Element | null {
  const {
    siteConfig: { customFields }
  } = useDocusaurusContext();

  return (
    <div>
      <footer className={clsx('footer', 'footer--dark')}>
        <div className="container container-fluid">
          <div className={styles.footerContainer}>
            <p className={styles.footerPromo}>
              Built with <a href='https://docusaurus.io/' target='_blank' rel="noopener" title='Docusaurus'>Docusaurus</a>
            </p>

            <div className={styles.footerIcons}>    
              <a href="https://github.com/pnp/cli-microsoft365" target="_blank" rel="noopener" title="GitHub">
                <GitHubSVG className={styles.svg} />
              </a>

              <a href="https://aka.ms/cli-m365/discord" target="_blank" rel="noopener" title="Discord">
                <DiscordSVG className={styles.svg} />
              </a>

              <a href="https://twitter.com/climicrosoft365" target="_blank" rel="noopener" title="Twitter">
                <TwitterSVG className={styles.svg} />
              </a>

              <a href="http://aka.ms/sppnp-videos" target="_blank" rel="noopener" title="YouTube">
                <YouTubeSVG className={styles.svg} />
              </a>

              <a href="https://aka.ms/sppnp" target="_blank" rel="noopener" title="Microsoft 365 & Power Platform Community Website">
                <LinkSVG className={styles.svg} />
              </a>
            </div>
          </div>
        </div>
      </footer>

      <MendableFloatingButton
        anon_key={customFields.mendableAnonKey as string}
        style={{
          darkMode: useColorMode().colorMode === "dark",
          accentColor: '#ef5552'
        }} 
        floatingButtonStyle={{
          color: '#fff',
          backgroundColor: '#ef5552'
        }}
        cmdShortcutKey='m'
      />
    </div>
  );
}

export default React.memo(Footer);
