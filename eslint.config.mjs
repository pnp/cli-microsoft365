import globals from 'globals';
import js from '@eslint/js';
import parser from '@typescript-eslint/parser';
import tsPlugin from '@typescript-eslint/eslint-plugin';
import mocha from 'eslint-plugin-mocha';
import stylistic from '@stylistic/eslint-plugin';
import cliM365 from './eslint-rules/lib/index.js';

// List of words used in command names used for word breaking
// Sorted alphabetically for easy maintenance
const dictionary = [
  'access',
  'activation',
  'activations',
  'adaptive',
  'administrative',
  'ai',
  'app',
  'application',
  'apply',
  'approve',
  'assessment',
  'assets',
  'assignment',
  'audit',
  'autofill',
  'azure',
  'bin',
  'builder',
  'call',
  'card',
  'catalog',
  'checklist',
  'client',
  'comm',
  'command',
  'community',
  'container',
  'content',
  'conversation',
  'custom',
  'customizer',
  'dataverse',
  'default',
  'definition',
  'dev',
  'details',
  'directory',
  'eligibility',
  'enterprise',
  'entra',
  'event',
  'eventreceiver',
  'external',
  'externalize',
  'folder',
  'fun',
  'group',
  'groupify',
  'groupmembership',
  'guest',
  'health',
  'hide',
  'historical',
  'home',
  'hub',
  'in',
  'info',
  'inheritance',
  'init',
  'install',
  'installed',
  'is',
  'issue',
  'item',
  'label',
  'list',
  'link',
  'log',
  'login',
  'logout',
  'mailbox',
  'management',
  'member',
  'membership',
  'messaging',
  'model',
  'multitenant',
  'm365',
  'news',
  'oauth2',
  'office365',
  'one',
  'open',
  'ops',
  'org',
  'organization',
  'owner',
  'permission',
  'pim',
  'place',
  'policy',
  'profile',
  'pronouns',
  'property',
  'record',
  'records',
  'recycle',
  'registration',
  'request',
  'resolver',
  'retention',
  'revoke',
  'role',
  'room',
  'schema',
  'search',
  'sensitivity',
  'service',
  'session',
  'set',
  'setting',
  'settings',
  'setup',
  'sharing',
  'side',
  'site',
  'status',
  'storage',
  'table',
  'teams',
  'threat',
  'to',
  'todo',
  'token',
  'type',
  'unit',
  'url',
  'user',
  'value',
  'web',
  'webhook'
];

// List of words that should be capitalized in a specific way
const capitalized = [
  'OAuth2'
];

// Sort dictionary to put the longest words first
const sortedDictionary = dictionary.sort((a, b) => b.length - a.length);

export default [
  // Ignored files
  {
    ignores: [
      "**/package-generate/assets/**",
      "**/test-projects/**",
      "clientsidepages.ts",
      "**/*.d.ts",
      "**/*.js",
      "**/*.cjs"
    ]
  },
  // JS recommendations
  js.configs.recommended,
  {
    plugins: { '@typescript-eslint': tsPlugin },
    rules: tsPlugin.configs.recommended.rules
  },
  {
    languageOptions: {
      ecmaVersion: 2015,
      sourceType: 'module',
      parser: parser,
      parserOptions: {
        ecmaVersion: 2015,
        sourceType: 'module',
        project: './tsconfig.json'
      },
      globals: {
        ...globals.node,
        ...globals.commonjs,
        ...globals.es2021,
        ...globals.mocha,
        NodeJS: 'readonly'
      }
    },
    plugins: {
      '@typescript-eslint': tsPlugin,
      'cli-microsoft365': cliM365,
      '@stylistic': stylistic,
      mocha
    },
    rules: {
      'cli-microsoft365/correct-command-class-name': ['error', sortedDictionary, capitalized],
      'cli-microsoft365/correct-command-name': 'error',
      'cli-microsoft365/no-by-server-relative-url-usage': 'error',
      '@stylistic/indent': ['error', 2],
      '@stylistic/semi': ['error'],
      '@stylistic/comma-dangle': ['error', 'never'],
      '@stylistic/brace-style': [
        'error',
        'stroustrup',
        { allowSingleLine: true }
      ],
      '@typescript-eslint/no-explicit-any': 'off',
      '@typescript-eslint/no-var-requires': 'off',
      '@typescript-eslint/no-inferrable-types': 'off',
      '@typescript-eslint/no-non-null-assertion': 'off',
      '@typescript-eslint/explicit-module-boundary-types': [
        'error',
        { allowArgumentsExplicitlyTypedAsAny: true }
      ],
      '@typescript-eslint/no-unused-vars': [
        'error',
        { argsIgnorePattern: '^_' }
      ],
      camelcase: ['error', {
        allow: [
          'child_process',
          'error_description',
          '_Child_Items_',
          '_Object_Type_',
          'FN\\d+',
          'OData__.*',
          'vti_.*',
          'Query.*',
          'app_displayname',
          'access_token',
          'expires_on',
          'extension_*'
        ]
      }],
      curly: ['error', 'all'],
      eqeqeq: ['error', 'always'],
      '@typescript-eslint/naming-convention': [
        'error',
        {
          selector: ['method'],
          format: ['camelCase']
        }
      ],
      '@typescript-eslint/explicit-function-return-type': ['error', { allowExpressions: true }],
      'mocha/no-identical-title': 'error',
      '@typescript-eslint/no-floating-promises': 'error',
      '@typescript-eslint/no-empty-function': 'error'
    }
  },
  {
    files: ['**/*.spec.ts'],
    rules: {
      'no-console': 'error',
      '@typescript-eslint/no-empty-function': 'off',
      'cli-microsoft365/correct-command-class-name': 'off',
      'cli-microsoft365/correct-command-name': 'off'
    }
  },
  {
    files: ['**/viva/commands/engage/**'],
    rules: {
      camelcase: 'off'
    }
  },
  {
    files: ['**/*.mjs'],
    rules: {
      '@typescript-eslint/explicit-function-return-type': 'off'
    }
  }
];