// @ts-check
// Note: type annotations allow type checking and IDEs autocompletion

const lightCodeTheme = require('prism-react-renderer/themes/nightOwlLight');
const darkCodeTheme = require('prism-react-renderer/themes/oceanicNext');

/** @type {import('@docusaurus/types').Config} */
const config = {
  title: 'CLI for Microsoft 365',
  titleDelimiter: '-',
  tagline: 'Docs',
  url: 'https://pnp.github.io',
  baseUrl: '/cli-microsoft365/',
  onBrokenLinks: 'throw',
  onBrokenMarkdownLinks: 'throw',
  favicon: 'img/favicon.ico',
  organizationName: 'pnp',
  projectName: 'cli-microsoft365',

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
      'docusaurus-node-polyfills',
      {
        excludeAliases: ['console']
      }
    ],
    [
      '@docusaurus/plugin-google-gtag',
      {
        trackingID: 'G-DH3T88LK5K',
        anonymizeIP: true,
      }
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
          customCss: require.resolve('./src/scss/Global.module.scss'),
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
        theme: lightCodeTheme,
        darkTheme: darkCodeTheme,
      }
    })
};

module.exports = config;
