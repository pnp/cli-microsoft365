// @ts-check
// Note: type annotations allow type checking and IDEs autocompletion

// Todo: this file needs to be removed before going live

const lightCodeTheme = require('prism-react-renderer/themes/nightOwlLight');
const darkCodeTheme = require('prism-react-renderer/themes/oceanicNext');

/** @type {import('@docusaurus/types').Config} */
const config = {
  title: 'CLI for Microsoft 365',
  titleDelimiter: '-',
  tagline: 'Docs',
  url: 'https://jwaegebaert.github.io',
  baseUrl: '/cli-microsoft365-docusaurus/',
  onBrokenLinks: 'throw',
  onBrokenMarkdownLinks: 'throw',
  favicon: 'img/favicon.ico',
  organizationName: 'jwaegebaert', 
  projectName: 'cli-microsoft365-docusaurus', 

  i18n: {
    defaultLocale: 'en',
    locales: ['en'],
  },

  customFields: {
    mendableAnonKey: 'd3313d54-6f8e-40e0-90d3-4095019d4be7',
  },

  plugins: [
    'docusaurus-plugin-sass',
    [
      'docusaurus-node-polyfills', { excludeAliases: ['console']}
    ]
  ],

  presets: [
    [
      'classic',
      /** @type {import('@docusaurus/preset-classic').Options} */
      ({
        docs: {
          routeBasePath: '/',
          sidebarPath: require.resolve('./sidebars.js'),
          editUrl: 'https://github.com/pnp/cli-microsoft365/blob/main/docs',
          showLastUpdateTime: true
        },
        blog: false,
        theme: {
          customCss: require.resolve('./src/scss/Global.module.scss')
        }
      })
    ]
  ],

  themeConfig:
    /** @type {import('@docusaurus/preset-classic').ThemeConfig} */
    ({
      navbar: {
        title: '',
        style: 'primary',
        logo: {
          alt: 'CLI for Microsoft 365 Logo',
          src: 'img/pnp-cli-microsoft365-white.svg'
        },
        items: [
          {            
            type: 'docSidebar',
            label: 'Home',            
            sidebarId: 'home',
            position: 'left'
          },
          {            
            type: 'docSidebar',
            label: 'User Guide',            
            sidebarId: 'userGuide',
            position: 'left'
          },
          {            
            type: 'docSidebar',
            label: 'Commands',            
            sidebarId: 'commands',
            position: 'left'
          },
          {            
            type: 'docSidebar',
            label: 'Concepts',            
            sidebarId: 'concepts',
            position: 'left'
          },
          {            
            type: 'docSidebar',
            label: 'Sample Scripts',            
            sidebarId: 'sampleScripts',
            position: 'left'
          },
          {            
            type: 'docSidebar',
            label: 'About',            
            sidebarId: 'about',
            position: 'left'
          },
          {
            href: 'https://github.com/pnp/cli-microsoft365',
            label: 'GitHub',
            position: 'right'
          }
        ]
      },
      prism: {
        theme: require('./src/config/lightCodeTheme.js'),
        darkTheme: require('./src/config/darkCodeTheme.js'),
        additionalLanguages: ['powershell', 'csv']
      },
      announcementBar: {
        id: 'join_discord',
        content:
          `<a href="https://aka.ms/cli-m365/discord" target="_blank" class="cli-announcement">
            Join our <strong>community Discord server</strong>
            <span>
              <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 640 512"><!--! Font Awesome Free 6.2.0 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license/free (Icons: CC BY 4.0, Fonts: SIL OFL 1.1, Code: MIT License) Copyright 2022 Fonticons, Inc.--><path d="M524.531 69.836a1.5 1.5 0 0 0-.764-.7A485.065 485.065 0 0 0 404.081 32.03a1.816 1.816 0 0 0-1.923.91 337.461 337.461 0 0 0-14.9 30.6 447.848 447.848 0 0 0-134.426 0 309.541 309.541 0 0 0-15.135-30.6 1.89 1.89 0 0 0-1.924-.91 483.689 483.689 0 0 0-119.688 37.107 1.712 1.712 0 0 0-.788.676C39.068 183.651 18.186 294.69 28.43 404.354a2.016 2.016 0 0 0 .765 1.375 487.666 487.666 0 0 0 146.825 74.189 1.9 1.9 0 0 0 2.063-.676A348.2 348.2 0 0 0 208.12 430.4a1.86 1.86 0 0 0-1.019-2.588 321.173 321.173 0 0 1-45.868-21.853 1.885 1.885 0 0 1-.185-3.126 251.047 251.047 0 0 0 9.109-7.137 1.819 1.819 0 0 1 1.9-.256c96.229 43.917 200.41 43.917 295.5 0a1.812 1.812 0 0 1 1.924.233 234.533 234.533 0 0 0 9.132 7.16 1.884 1.884 0 0 1-.162 3.126 301.407 301.407 0 0 1-45.89 21.83 1.875 1.875 0 0 0-1 2.611 391.055 391.055 0 0 0 30.014 48.815 1.864 1.864 0 0 0 2.063.7A486.048 486.048 0 0 0 610.7 405.729a1.882 1.882 0 0 0 .765-1.352c12.264-126.783-20.532-236.912-86.934-334.541ZM222.491 337.58c-28.972 0-52.844-26.587-52.844-59.239s23.409-59.241 52.844-59.241c29.665 0 53.306 26.82 52.843 59.239 0 32.654-23.41 59.241-52.843 59.241Zm195.38 0c-28.971 0-52.843-26.587-52.843-59.239s23.409-59.241 52.843-59.241c29.667 0 53.307 26.82 52.844 59.239 0 32.654-23.177 59.241-52.844 59.241Z"></path></svg>
            </span>
          </a>`,
        backgroundColor: '#1b1c23',
        textColor: '#f5f5f5',
        isCloseable: true,
      }
    })
};

module.exports = config;
