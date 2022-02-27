import * as assert from 'assert';
import * as fs from 'fs';
import { Cli, Logger } from '../../cli';
import Utils from '../../Utils';
import AppCommand from './AppCommand';
import sinon = require('sinon');

class MockCommand extends AppCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(): void {
  }

  public commandHelp(): void {
  }
}

describe('AppCommand', () => {
  let cmd: MockCommand;
  let logger: Logger;
  let log: string[];

  beforeEach(() => {
    cmd = new MockCommand();
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
  });

  afterEach(() => {
    Utils.restore([
      fs.existsSync,
      fs.readFileSync,
      Cli.prompt
    ]);
  });

  it('defines correct resource', () => {
    assert.strictEqual((cmd as any).resource, 'https://graph.microsoft.com');
  });

  it('returns error if .m365rc.json file not found in the current directory', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    cmd.action(logger, { options: {} }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'Could not find file: .m365rc.json');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error if the .m365rc.json file is empty', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '');
    cmd.action(logger, { options: {} }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'File .m365rc.json is empty');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`returns error if the .m365rc.json file contents couldn't be parsed`, (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '{');
    cmd.action(logger, { options: {} }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'Could not parse file: .m365rc.json');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`returns error if the .m365rc.json file is empty doesn't contain any apps`, (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: []
    }));
    cmd.action(logger, { options: {} }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'No Azure AD apps found in .m365rc.json');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`returns error if the specified appId not found in the .m365rc.json file`, (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ]
    }));
    cmd.action(logger, { options: { appId: 'e23d235c-fcdf-45d1-ac5f-24ab2ee06951' } }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'App e23d235c-fcdf-45d1-ac5f-24ab2ee06951 not found in .m365rc.json');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`prompts to choose an app when multiple apps found in .m365rc.json and no appId specified`, (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        },
        {
          "appId": "9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d",
          "name": "CLI app1"
        }
      ]
    }));
    const cliPromptStub = sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { appIdIndex: number }) => void) => {
      cb({ appIdIndex: 0 });
    });
    cmd.action(logger, { options: {} }, () => {
      try {
        assert(cliPromptStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`uses app selected by the user in the prompt`, (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        },
        {
          "appId": "9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d",
          "name": "CLI app1"
        }
      ]
    }));
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { appIdIndex: number }) => void) => {
      cb({ appIdIndex: 1 });
    });
    cmd.action(logger, { options: {} }, () => {
      try {
        assert.strictEqual((cmd as any).appId, '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`uses app specified in the appId command option`, (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        },
        {
          "appId": "9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d",
          "name": "CLI app1"
        }
      ]
    }));
    cmd.action(logger, { options: { appId: '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d' } }, () => {
      try {
        assert.strictEqual((cmd as any).appId, '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`uses app from the .m365rc.json if only one app defined`, (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ]
    }));
    cmd.action(logger, { options: {} }, () => {
      try {
        assert.strictEqual((cmd as any).appId, 'e23d235c-fcdf-45d1-ac5f-24ab2ee0695d');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the specified appId is not a valid GUID', () => {
    const actual = cmd.validate({ options: { appId: 'e23d235c-fcdf-45d1-ac5f-24ab2ee0695' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the specified appId is a valid GUID', () => {
    const actual = cmd.validate({ options: { appId: 'e23d235c-fcdf-45d1-ac5f-24ab2ee0695d' } });
    assert.strictEqual(actual, true);
  });
});