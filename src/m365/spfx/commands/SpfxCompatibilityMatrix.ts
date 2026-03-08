export interface VersionCheck {
  /**
   * Required version range in semver
   */
  range: string;
  /**
   * What to do to fix it if the required range isn't met
   */
  fix: string;
}

/**
 * Versions of SharePoint that support SharePoint Framework
 */
export enum SharePointVersion {
  SP2016 = 1 << 0,
  SP2019 = 1 << 1,
  SPO = 1 << 2,
  All = ~(~0 << 3)
}

export interface SpfxVersionPrerequisites {
  gulpCli?: VersionCheck;
  heft?: VersionCheck;
  node: VersionCheck;
  sp: SharePointVersion;
  yo: VersionCheck;
}

export const versions: { [version: string]: SpfxVersionPrerequisites } = {
  '1.0.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^6',
      fix: 'Install Node.js v6'
    },
    sp: SharePointVersion.All,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.1.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^6',
      fix: 'Install Node.js v6'
    },
    sp: SharePointVersion.All,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.2.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^6',
      fix: 'Install Node.js v6'
    },
    sp: SharePointVersion.SP2019 | SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.4.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^6',
      fix: 'Install Node.js v6'
    },
    sp: SharePointVersion.SP2019 | SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.4.1': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^6 || ^8',
      fix: 'Install Node.js v8'
    },
    sp: SharePointVersion.SP2019 | SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.5.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^6 || ^8',
      fix: 'Install Node.js v8'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.5.1': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^6 || ^8',
      fix: 'Install Node.js v8'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.6.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^6 || ^8',
      fix: 'Install Node.js v8'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.7.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^8',
      fix: 'Install Node.js v8'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.7.1': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^8',
      fix: 'Install Node.js v8'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.8.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^8',
      fix: 'Install Node.js v8'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.8.1': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^8',
      fix: 'Install Node.js v8'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.8.2': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^8 || ^10',
      fix: 'Install Node.js v10'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.9.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^8 || ^10',
      fix: 'Install Node.js v10'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.9.1': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^10',
      fix: 'Install Node.js v10'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.10.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^10',
      fix: 'Install Node.js v10'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.11.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^10',
      fix: 'Install Node.js v10'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.12.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^12',
      fix: 'Install Node.js v12'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.12.1': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^12 || ^14',
      fix: 'Install Node.js v12 or v14'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^3',
      fix: 'npm i -g yo@3'
    }
  },
  '1.13.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^12 || ^14',
      fix: 'Install Node.js v12 or v14'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.13.1': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^12 || ^14',
      fix: 'Install Node.js v12 or v14'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.14.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^12 || ^14',
      fix: 'Install Node.js v12 or v14'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.15.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^12.13 || ^14.15 || ^16.13',
      fix: 'Install Node.js v12.13, v14.15, v16.13 or higher'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.15.2': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '^12.13 || ^14.15 || ^16.13',
      fix: 'Install Node.js v12.13, v14.15, v16.13 or higher'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.16.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '>=16.13.0 <17.0.0',
      fix: 'Install Node.js >=16.13.0 <17.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.16.1': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '>=16.13.0 <17.0.0',
      fix: 'Install Node.js >=16.13.0 <17.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.17.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '>=16.13.0 <17.0.0',
      fix: 'Install Node.js >=16.13.0 <17.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.17.1': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '>=16.13.0 <17.0.0',
      fix: 'Install Node.js >=16.13.0 <17.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.17.2': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '>=16.13.0 <17.0.0',
      fix: 'Install Node.js >=16.13.0 <17.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.17.3': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '>=16.13.0 <17.0.0',
      fix: 'Install Node.js >=16.13.0 <17.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.17.4': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '>=16.13.0 <17.0.0',
      fix: 'Install Node.js >=16.13.0 <17.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.18.0': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '>=16.13.0 <17.0.0 || >=18.17.1 <19.0.0',
      fix: 'Install Node.js >=16.13.0 <17.0.0 || >=18.17.1 <19.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.18.1': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '>=16.13.0 <17.0.0 || >=18.17.1 <19.0.0',
      fix: 'Install Node.js >=16.13.0 <17.0.0 || >=18.17.1 <19.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4',
      fix: 'npm i -g yo@4'
    }
  },
  '1.18.2': {
    gulpCli: {
      range: '^1 || ^2',
      fix: 'npm i -g gulp-cli@2'
    },
    node: {
      range: '>=16.13.0 <17.0.0 || >=18.17.1 <19.0.0',
      fix: 'Install Node.js >=16.13.0 <17.0.0 || >=18.17.1 <19.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4 || ^5',
      fix: 'npm i -g yo@5'
    }
  },
  '1.19.0': {
    gulpCli: {
      range: '^1 || ^2 || ^3',
      fix: 'npm i -g gulp-cli@3'
    },
    node: {
      range: '>=18.17.1 <19.0.0',
      fix: 'Install Node.js >=18.17.1 <19.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4 || ^5',
      fix: 'npm i -g yo@5'
    }
  },
  '1.20.0': {
    gulpCli: {
      range: '^1 || ^2 || ^3',
      fix: 'npm i -g gulp-cli@3'
    },
    node: {
      range: '>=18.17.1 <19.0.0',
      fix: 'Install Node.js >=18.17.1 <19.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4 || ^5',
      fix: 'npm i -g yo@5'
    }
  },
  '1.21.0': {
    gulpCli: {
      range: '^1 || ^2 || ^3',
      fix: 'npm i -g gulp-cli@3'
    },
    node: {
      range: '>=22.14.0 <23.0.0',
      fix: 'Install Node.js >=22.14.0 <23.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4 || ^5',
      fix: 'npm i -g yo@5'
    }
  },
  '1.21.1': {
    gulpCli: {
      range: '^1 || ^2 || ^3',
      fix: 'npm i -g gulp-cli@3'
    },
    node: {
      range: '>=22.14.0 <23.0.0',
      fix: 'Install Node.js >=22.14.0 <23.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4 || ^5',
      fix: 'npm i -g yo@5'
    }
  },
  '1.22.0': {
    heft: {
      range: '^1',
      fix: 'npm i -g @rushstack/heft@1'
    },
    node: {
      range: '>=22.14.0 <23.0.0',
      fix: 'Install Node.js >=22.14.0 <23.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4 || ^5 || ^6',
      fix: 'npm i -g yo@6'
    }
  },
  '1.22.1': {
    heft: {
      range: '^1',
      fix: 'npm i -g @rushstack/heft@1'
    },
    node: {
      range: '>=22.14.0 <23.0.0',
      fix: 'Install Node.js >=22.14.0 <23.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4 || ^5 || ^6',
      fix: 'npm i -g yo@6'
    }
  },
  '1.22.2': {
    heft: {
      range: '^1',
      fix: 'npm i -g @rushstack/heft@1'
    },
    node: {
      range: '>=22.14.0 <23.0.0',
      fix: 'Install Node.js >=22.14.0 <23.0.0'
    },
    sp: SharePointVersion.SPO,
    yo: {
      range: '^4 || ^5 || ^6',
      fix: 'npm i -g yo@6'
    }
  }
};