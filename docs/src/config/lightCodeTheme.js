/** @type {import('prism-react-renderer').PrismTheme} */
const lightCodeTheme = {
  plain: {
    color: '#1c1e21',
    backgroundColor: '#F6F7F8',
  },
  styles: [
    {
      types: ['comment'],
      style: {
        color: '#008000',
      },
    },
    {
      types: ['builtin'],
      style: {
        color: '#0070C1',
      },
    },
    {
      types: ['number', 'variable', 'inserted'],
      style: {
        color: '#098658',
      },
    },
    {
      types: ['operator'],
      style: {
        color: '#000000',
      },
    },
    {
      types: ['constant', 'char'],
      style: {
        color: '#811F3F',
      },
    },
    {
      types: ['tag'],
      style: {
        color: '#800000',
      },
    },
    {
      types: ['attr-name'],
      style: {
        color: '#FF0000',
      },
    },
    {
      types: ['deleted', 'string'],
      style: {
        color: '#A31515',
      },
    },
    {
      types: ['changed', 'punctuation'],
      style: {
        color: '#0451A5',
      },
    },
    {
      types: ['function', 'keyword'],
      style: {
        color: '#0000FF',
      },
    },
    {
      types: ['class-name'],
      style: {
        color: '#267F99',
      },
    },
    {
      types: ['title', 'table-header', 'table-row'],
      languages: ['md', 'markdown'],
      style: {
        color: '#0451A5'
      }
    },
    {
      types: ['property'],
      languages: ['json'],
      style: {
        color: '#124994'
      }
    },
    {
      types: ['punctuation'],
      languages: ['json'],
      style: {
        color: '#000000'
      }
    },
    {
      types: ['function'],
      languages: ['powershell'],
      style: {
        color: '#000000'
      }
    },
    {
      types: ['class-name'],
      languages: ['bash'],
      style: {
        color: '#000000'
      }
    },
    {
      types: ['shebang'],
      languages: ['bash'],
      style: {
        color: '#A0A0A0'
      }
    }
  ],
};

module.exports = lightCodeTheme;