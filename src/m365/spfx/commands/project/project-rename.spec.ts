import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./project-rename');

describe(commands.PROJECT_RENAME, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
  let writeFileSyncSpy: sinon.SinonStub;
  const projectPath: string = 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react';

  before(() => {
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    telemetry = null;
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    writeFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
  });

  afterEach(() => {
    sinonUtil.restore([
      (command as any).generateNewId,
      (command as any).getProjectRoot,
      (command as any).getProject,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PROJECT_RENAME), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('calls telemetry', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));

    command.action(logger, { options: { newName: 'spfx-react' } }, () => {
      assert(trackEvent.called);
    });
  });

  it('logs correct telemetry event', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));

    command.action(logger, { options: { newName: 'spfx-react' } }, () => {
      assert.strictEqual(telemetry.name, commands.PROJECT_RENAME);
    });
  });

  it('shows error if the project path couldn\'t be determined', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => null);

    command.action(logger, { options: { newName: 'spfx-react' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Couldn't find project root folder`, 1)));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('updates only the files found and skips other files', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));
    sinon.stub(command as any, 'getProject').callsFake(_ => {
      return {
        path: projectPath,
        packageJson: {
          dependencies: {},
          name: 'spfx'
        }
      };
    });
    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    command.action(logger, { options: { newName: 'spfx-react' } } as any, () => {
      try {
        assert(writeFileSyncSpy.notCalled);
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('handles error while updating the files', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));
    sinon.stub(command as any, 'getProject').callsFake(_ => {
      return {
        path: projectPath,
        packageJson: {
          dependencies: {},
          name: 'spfx'
        }
      };
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => { throw 'error'; });
    command.action(logger, { options: { newName: 'spfx-react' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('error')));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('replaces project name in package.json', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));

    const replacedContent = `{
  "name": "spfx-react",
  "version": "0.0.1",
  "private": true,
  "engines": {
    "node": ">=0.10.0"
  },
  "scripts": {
    "build": "gulp bundle",
    "clean": "gulp clean",
    "test": "gulp test"
  },
  "dependencies": {
    "react": "16.7.0",
    "react-dom": "16.7.0",
    "@types/react": "16.7.22",
    "@types/react-dom": "16.8.0",
    "office-ui-fabric-react": "6.143.0",
    "@microsoft/sp-core-library": "1.8.2",
    "@microsoft/sp-property-pane": "1.8.2",
    "@microsoft/sp-webpart-base": "1.8.2",
    "@microsoft/sp-lodash-subset": "1.8.2",
    "@microsoft/sp-office-ui-fabric-core": "1.8.2",
    "@types/webpack-env": "1.13.1",
    "@types/es6-promise": "0.0.33"
  },
  "resolutions": {
    "@types/react": "16.7.22"
  },
  "devDependencies": {
    "@microsoft/sp-build-web": "1.8.2",
    "@microsoft/sp-tslint-rules": "1.8.2",
    "@microsoft/sp-module-interfaces": "1.8.2",
    "@microsoft/sp-webpart-workbench": "1.8.2",
    "@microsoft/rush-stack-compiler-2.9": "0.7.7",
    "gulp": "~3.9.1",
    "@types/chai": "3.4.34",
    "@types/mocha": "2.2.38",
    "ajv": "~5.2.2"
  }
}`;

    command.action(logger, { options: { newName: 'spfx-react', generateNewId: true } } as any, () => {
      try {
        assert(writeFileSyncSpy.calledWith(sinon.match.string, replacedContent, 'utf-8'));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('replaces only project name in .yo-rc.json when --generateNewId is not passed', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));

    const replacedContent = `{
  "@microsoft/generator-sharepoint": {
    "version": "1.8.2",
    "libraryName": "spfx-react",
    "libraryId": "da1c365f-1532-4e10-aca2-7a0d29c3245b",
    "environment": "spo",
    "packageManager": "npm",
    "solutionName": "spfx-react",
    "skipFeatureDeployment": false,
    "componentType": "webpart",
    "framework": "react",
    "componentName": "HelloWorld",
    "componentDescription": "HelloWorld",
    "isCreatingSolution": true,
    "isDomainIsolated": false
  }
}`;

    command.action(logger, { options: { newName: 'spfx-react' } } as any, () => {
      try {
        assert(writeFileSyncSpy.calledWith(sinon.match.string, replacedContent, 'utf-8'));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('replaces project name and id in .yo-rc.json when --generateNewId is passed', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));

    sinon.stub((command as any), 'generateNewId').callsFake(() => {
      return '69cb6882-acc1-4148-b059-31ae149ba077';
    });

    const replacedContent = `{
  "@microsoft/generator-sharepoint": {
    "version": "1.8.2",
    "libraryName": "spfx-react",
    "libraryId": "69cb6882-acc1-4148-b059-31ae149ba077",
    "environment": "spo",
    "packageManager": "npm",
    "solutionName": "spfx-react",
    "skipFeatureDeployment": false,
    "componentType": "webpart",
    "framework": "react",
    "componentName": "HelloWorld",
    "componentDescription": "HelloWorld",
    "isCreatingSolution": true,
    "isDomainIsolated": false
  }
}`;

    command.action(logger, { options: { newName: 'spfx-react', generateNewId: true, debug: true } } as any, () => {
      try {
        assert(writeFileSyncSpy.calledWith(sinon.match.string, replacedContent, 'utf-8'));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('replaces only project name in package-solution.json when --generateNewId is not passed', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));

    const replacedContent = `{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
  "solution": {
    "name": "spfx-react-client-side-solution",
    "id": "da1c365f-1532-4e10-aca2-7a0d29c3245b",
    "version": "1.0.0.0",
    "includeClientSideAssets": true,
    "isDomainIsolated": false
  },
  "paths": {
    "zippedPackage": "solution/spfx-react.sppkg"
  }
}`;

    command.action(logger, { options: { newName: 'spfx-react' } } as any, () => {
      try {
        assert(writeFileSyncSpy.calledWith(sinon.match.string, replacedContent, 'utf-8'));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('replaces project name and id in package-solution.json when --generateNewId is passed', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));

    sinon.stub((command as any), 'generateNewId').callsFake(() => {
      return '69cb6882-acc1-4148-b059-31ae149ba077';
    });

    const replacedContent = `{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
  "solution": {
    "name": "spfx-react-client-side-solution",
    "id": "69cb6882-acc1-4148-b059-31ae149ba077",
    "version": "1.0.0.0",
    "includeClientSideAssets": true,
    "isDomainIsolated": false
  },
  "paths": {
    "zippedPackage": "solution/spfx-react.sppkg"
  }
}`;

    command.action(logger, { options: { newName: 'spfx-react', generateNewId: true } } as any, () => {
      try {
        assert(writeFileSyncSpy.calledWith(sinon.match.string, replacedContent, 'utf-8'));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('replaces project name in deploy-azure-storage.json', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));

    const replacedContent = `{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/deploy-azure-storage.schema.json",
  "workingDir": "./temp/deploy/",
  "account": "<!-- STORAGE ACCOUNT NAME -->",
  "container": "spfx-react",
  "accessKey": "<!-- ACCESS KEY -->"
}`;

    command.action(logger, { options: { newName: 'spfx-react' } } as any, () => {
      try {
        assert(writeFileSyncSpy.calledWith(sinon.match.string, replacedContent, 'utf-8'));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('replaces project name in README.md', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));

    let replacedContent = `## spfx-react

This is where you include your WebPart documentation.

### Building the code

\`\`\`bash
git clone the repo
npm i
npm i -g gulp
gulp
\`\`\`

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
`;

    command.action(logger, { options: { newName: 'spfx-react', debug: true } } as any, () => {
      try {
        let fileSyncContent: string = writeFileSyncSpy.lastCall.args[1];
        fileSyncContent = fileSyncContent.replace(/(\r\n|\n|\r)/gm, "");
        replacedContent = replacedContent.replace(/(\r\n|\n|\r)/gm, "");
        assert.strictEqual(fileSyncContent, replacedContent);
        assert.strictEqual(loggerLogToStderrSpy.getCall(5).args[0], `Updated README.md`);
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});