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

  plugins: ['docusaurus-plugin-sass'],

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
          customCss: require.resolve('./src/css/custom.css'),
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
            href: 'https://github.com/facebook/docusaurus',
            label: 'GitHub',
            position: 'right'
          }
        ]
      },
      footer: {
        style: 'dark',
        links: [
          {
            title: 'Community',
            items: [
              {
                label: 'Discord',
                href: 'https://discordapp.com/invite/7rfW4kg6B5',
              },
              {
                label: 'Twitter',
                href: 'https://twitter.com/climicrosoft365',
              }
            ]
          },
          {
            title: 'Resources',
            items: [
              {
                label: 'GitHub',
                href: 'https://github.com/pnp/cli-microsoft365',
              },
              {
                label: 'YouTube',
                href: 'http://aka.ms/sppnp-videos',
              },
              {
                label: 'Microsoft 365 & Power Platform community',
                href: 'https://aka.ms/sppnp',
              }
            ]
          }
        ],
        copyright: `Copyright © ${new Date().getFullYear()} CLI for Microsoft 365. Built with Docusaurus.`
      },
      prism: {
        theme: lightCodeTheme,
        darkTheme: darkCodeTheme,
      }
    })
};

module.exports = config;
