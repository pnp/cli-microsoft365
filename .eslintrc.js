// list of words used in command names used for word breaking
// sorted alphabetically for easy maintenance
const dictionary = [
  'access',
  'activation',
  'activations',
  'adaptive',
  'app',
  'apply',
  'approve',
  'assets',
  'bin',
  'catalog',
  'client',
  'comm',
  'content',
  'conversation',
  'custom',
  'default',
  'external',
  'externalize',
  'fun',
  'group',
  'groupify',
  'guest',
  'hide',
  'historical',
  'home',
  'hub',
  'in',
  'init',
  'install',
  'installed',
  'is',
  'issue',
  'list',
  'member',
  'messaging',
  'news',
  'oauth2',
  'office365',
  'one',
  'org',
  'o365',
  'permission',
  'place',
  'property',
  'records',
  'recycle',
  'role',
  'schema',
  'service',
  'setting',
  'settings',
  'side',
  'site',
  'status',
  'storage',
  'token',
  'type',
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
    "cli-microsoft365"
  ],
  "ignorePatterns": [
    "**/pcf-init/assets/**",
    "**/solution-init/assets/**",
    "**/package-generate/assets/**",
    "**/test-projects/**",
    "clientsidepages.ts",
    "*.js"
  ],
  "rules": {
    "cli-microsoft365/correct-command-class-name": ["error", sortedDictionary, capitalized],
    "cli-microsoft365/correct-command-name": "error",
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
    ]
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
        "cli-microsoft365/correct-command-name": "off"
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
}
