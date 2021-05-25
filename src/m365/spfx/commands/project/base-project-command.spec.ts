import * as fs from 'fs';
import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from '../../../../Utils';
import { BaseProjectCommand } from "./base-project-command";
import { Project } from "./model";

class MockCommand extends BaseProjectCommand {
  public get name(): string {
    return 'Mock';
  }
  public get description(): string {
    return 'Mock command';
  }
  public commandAction(): void {
  }

  public getProjectPublic(): Project {
    return this.getProject('./src/m365/spfx/commands/project/test-projects/spfx-141-webpart-nolib');
  }
}

describe('BaseProjectCommand', () => {
  afterEach(() => {
    Utils.restore([
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
});