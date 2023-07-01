import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { BaseProjectCommand } from "./base-project-command.js";
import { Project } from "./project-model/index.js";

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
});
