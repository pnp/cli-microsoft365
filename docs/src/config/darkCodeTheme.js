/** @type {import('prism-react-renderer').PrismTheme} */
const darkCodeTheme = {
  plain: {
    color: '#E3E3E3',
    backgroundColor: '#323234',
  },
  styles: [
    {
      types: ['prolog'],
      style: {
        color: '#000080'
      }
    },
    {
      types: ['comment'],
      style: {
        color: '#6A9955'
      }
    },
    {
      types: ['builtin', 'changed', 'keyword', 'interpolation-punctuation'],
      style: {
        color: '#569CD6'
      }
    },
    {
      types: ['number', 'inserted'],
      style: {
        color: '#B5CEA8'
      }
    },
    {
      types: ['constant'],
      style: {
        color: '#646695'
      }
    },
    {
      types: ['attr-name', 'variable'],
      style: {
        color: '#9CDCFE'
      }
    },
    {
      types: ['deleted', 'string', 'attr-value', 'template-punctuation'],
      style: {
        color: '#CE9178'
      }
    },
    {
      types: ['selector'],
      style: {
        color: '#D7BA7D'
      }
    },
    {
      // Fix tag color
      types: ['tag'],
      style: {
        color: '#4EC9B0'
      }
    },
    {
      // Fix tag color for HTML
      types: ['tag'],
      languages: ['markup'],
      style: {
        color: '#569CD6'
      }
    },
    {
      types: ['punctuation', 'operator'],
      style: {
        color: '#D4D4D4'
      }
    },
    {
      // Fix punctuation color for HTML
      types: ['punctuation'],
      languages: ['markup'],
      style: {
        color: '#808080'
      }
    },
    {
      types: ['function'],
      style: {
        color: '#DCDCAA'
      }
    },
    {
      types: ['class-name'],
      style: {
        color: '#4EC9B0'
      }
    },
    {
      types: ['char'],
      style: {
        color: '#D16969'
      }
    },
    {
      types: ['title', 'punctuation', 'table-header', 'table-row'],
      languages: ['md', 'markdown', 'csv'],
      style: {
        color: '#84BDDA'
      }
    },
    {
      types: ['function'],
      languages: ['powershell'],
      style: {
        color: '#E3E3E3'
      }
    },
    {
      types: ['class-name'],
      languages: ['bash'],
      style: {
        color: '#E3E3E3'
      }
    }
  ]
};

module.exports = darkCodeTheme;