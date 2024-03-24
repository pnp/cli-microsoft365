import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { BaseProjectCommand } from "./base-project-command.js";
import { Project } from "./project-model/index.js";
import { CommandError } from '../../../../Command.js';

class MockCommand extends BaseProjectCommand {
  public get name(): string {
    return 'Mock';
  }
  public get description(): string {
    return 'Mock command';
  }
  public async commandAction(): Promise<void> {
  }

  public getProjectPublic(): Project {
    return this.getProject('./src/m365/spfx/commands/project/test-projects/spfx-141-webpart-nolib');
  }
}

describe('BaseProjectCommand', () => {
  const projectPath: string = 'src/m365/spfx/commands/project/test-projects/spfx-141-webpart-nolib';

  const scenarios = [
    { file: 'config/config.json', invalidJson: '{ "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/config.2.0.schema.json", "version": "2.0", "bundles":' },
    { file: 'config/copy-assets.json', invalidJson: '' },
    { file: 'config/deploy-azure-storage.json', invalidJson: '' },
    { file: 'package.json', invalidJson: '' },
    { file: 'config/package-solution.json', invalidJson: '' },
    { file: 'config/serve.json', invalidJson: '' },
    { file: 'tsconfig.json', invalidJson: '' },
    { file: 'config/tslint.json', invalidJson: '' },
    { file: 'tslint.json', invalidJson: '' },
    { file: 'config/write-manifests.json', invalidJson: '' },
    { file: '.yo-rc.json', invalidJson: '' },
    { file: '.vscode/settings.json', invalidJson: '' },
    { file: '.vscode/extensions.json', invalidJson: '' },
    { file: '.vscode/launch.json', invalidJson: '' }
  ];

  afterEach(() => {
    sinonUtil.restore([
      fs.readFileSync,
      fs.existsSync
    ]);
  });

  it(`doesn't fail if reading .gitignore file contents failed`, () => {
    const readFileSyncOriginal = fs.readFileSync;
    const existsSyncOriginal = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.gitignore') > -1) {
        return true;
      }
      else {
        return existsSyncOriginal(path);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().indexOf('.gitignore') > -1) {
        throw new Error();
      }
      else {
        return readFileSyncOriginal(path, encoding);
      }
    });

    const command = new MockCommand();
    const project = command.getProjectPublic();
    assert.notStrictEqual(typeof project, 'undefined');
  });

  it(`doesn't fail if reading .npmignore file contents failed`, () => {
    const readFileSyncOriginal = fs.readFileSync;
    const existsSyncOriginal = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.npmignore') > -1) {
        return true;
      }
      else {
        return existsSyncOriginal(path);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().indexOf('.npmignore') > -1) {
        throw new Error();
      }
      else {
        return readFileSyncOriginal(path, encoding);
      }
    });

    const command = new MockCommand();
    const project = command.getProjectPublic();
    assert.notStrictEqual(typeof project, 'undefined');
  });

  scenarios.forEach(({ file, invalidJson }) => {
    it(`throws CommandError when '${file}' contains invalid JSON`, () => {
      const readFileSyncOriginal = fs.readFileSync;
      const existsSyncOriginal = fs.existsSync;
      sinon.stub(fs, 'existsSync').callsFake(path => {
        if (path.toString() === projectPath + '/' + file) {
          return true;
        }
        else {
          return existsSyncOriginal(path);
        }
      });

      sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
        if (path.toString() === projectPath + '/' + file) {
          return invalidJson;
        }
        else {
          return readFileSyncOriginal(path, encoding);
        }
      });

      const command = new MockCommand();

      assert.throws(() => {
        command.getProjectPublic();
      }, new CommandError('The file ' + projectPath + '/' + file + ' is not a valid JSON file or is not utf-8 encoded.')); // Adjust this to match how CommandError is detected in your assertions
    });
  });

});
