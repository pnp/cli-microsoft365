// list of words used in command names used for word breaking
// sorted alphabetically for easy maintenance
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
  'audit',
  'bin',
  'builder',
  'card',
  'catalog',
  'checklist',
  'client',
  'comm',
  'command',
  'content',
  'conversation',
  'custom',
  'customizer',
  'dataverse',
  'default',
  'event',
  'eventreceiver',
  'external',
  'externalize',
  'fun',
  'group',
  'groupify',
  'guest',
  'health',
  'hide',
  'historical',
  'home',
  'hub',
  'identity',
  'in',
  'info',
  'inheritance',
  'init',
  'install',
  'installed',
  'is',
  'issue',
  'label',
  'list',
  'link',
  'log',
  'login',
  'logout',
  'management',
  'member',
  'messaging',
  'model',
  'm365',
  'news',
  'oauth2',
  'office365',
  'one',
  'org',
  'owner',
  'permission',
  'place',
  'policy',
  'profile',
  'property',
  'records',
  'recycle',
  'retention',
  'role',
  'room',
  'schema',
  'sensitivity',
  'service',
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
  'token',
  'type',
  'unit',
  'user',
  'web',
  'webhook'
];

// list of words that should be capitalized in a specific way
const capitalized = [
  'OAuth2'
];

// sort dictionary to put the longest words first
const sortedDictionary = dictionary.sort((a, b) => b.length - a.length);

module.exports = {
  "root": true,
  "env": {
    "node": true,
    "es2021": true,
    "commonjs": true,
    "mocha": true
  },
  "globals": {
    "NodeJS": true
  },
  "extends": [
    "plugin:@typescript-eslint/recommended"
  ],
  "parser": "@typescript-eslint/parser",
  "parserOptions": {
    "ecmaVersion": 2015,
    "sourceType": "module"
  },
  "plugins": [
    "@typescript-eslint",
    "cli-microsoft365",
    "mocha"
  ],
  "ignorePatterns": [
    "**/package-generate/assets/**",
    "**/test-projects/**",
    "clientsidepages.ts",
    "*.js"
  ],
  "rules": {
    "cli-microsoft365/correct-command-class-name": ["error", sortedDictionary, capitalized],
    "cli-microsoft365/correct-command-name": "error",
    "cli-microsoft365/no-by-server-relative-url-usage": "error",
    "indent": "off",
    "@typescript-eslint/indent": [
      "error",
      2
    ],
    "semi": "off",
    "@typescript-eslint/semi": [
      "error"
    ],
    "@typescript-eslint/no-explicit-any": "off",
    "@typescript-eslint/no-var-requires": "off",
    "@typescript-eslint/no-inferrable-types": "off",
    "@typescript-eslint/no-non-null-assertion": "off",
    "@typescript-eslint/explicit-module-boundary-types": [
      "error",
      {
        "allowArgumentsExplicitlyTypedAsAny": true
      }
    ],
    "@typescript-eslint/no-unused-vars": [
      "error",
      {
        "argsIgnorePattern": "^_"
      }
    ],
    "brace-style": [
      "error",
      "stroustrup",
      {
        "allowSingleLine": true
      }
    ],
    "camelcase": [
      "error",
      {
        "allow": [
          "child_process",
          "error_description",
          "_Child_Items_",
          "_Object_Type_",
          "FN\\d+",
          "OData__.*",
          "vti_.*",
          "Query.*",
          "app_displayname",
          "access_token",
          "expires_on"
        ]
      }
    ],
    "comma-dangle": [
      "error",
      "never"
    ],
    "curly": [
      "error",
      "all"
    ],
    "eqeqeq": [
      "error",
      "always"
    ],
    "@typescript-eslint/naming-convention": [
      "error",
      {
        "selector": [
          "method"
        ],
        "format": [
          "camelCase"
        ]
      }
    ],
    "@typescript-eslint/explicit-function-return-type": ["error", { "allowExpressions": true }],
    "mocha/no-identical-title": "error"
  },
  "overrides": [
    {
      "files": [
        "*.spec.ts"
      ],
      "rules": {
        "no-console": "error",
        "@typescript-eslint/no-empty-function": "off",
        "cli-microsoft365/correct-command-class-name": "off",
        "cli-microsoft365/correct-command-name": "off",
        "@typescript-eslint/explicit-function-return-type": "off"
      }
    },
    {
      "files": [
        "**/yammer/**"
      ],
      "rules": {
        "camelcase": "off"
      }
    }
  ]
};
