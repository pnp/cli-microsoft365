import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './app-update.js';

describe(commands.APP_UPDATE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const fsStats: fs.Stats = {
    isDirectory: () => false,
    isFile: () => false,
    isBlockDevice: () => false,
    isCharacterDevice: () => false,
    isSymbolicLink: () => false,
    isFIFO: () => false,
    isSocket: () => false,
    dev: 0,
    ino: 0,
    mode: 0,
    nlink: 0,
    uid: 0,
    gid: 0,
    rdev: 0,
    size: 0,
    blksize: 0,
    blocks: 0,
    atimeMs: 0,
    mtimeMs: 0,
    ctimeMs: 0,
    birthtimeMs: 0,
    atime: new Date(),
    mtime: new Date(),
    ctime: new Date(),
    birthtime: new Date()
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.put,
      fs.readFileSync,
      fs.existsSync,
      fs.lstatSync,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_UPDATE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both id and name options are passed', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'e3e29acb-8c79-412b-b746-e6c39ff4cd22',
      name: 'Test app',
      filePath: 'teamsapp.zip'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both id and name options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({
      filePath: 'teamsapp.zip'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if the id is not a valid GUID.', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'invalid',
      filePath: 'teamsapp.zip'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if the filePath does not exist', () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = commandOptionsSchema.safeParse({ id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22", filePath: 'invalid.zip' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if the filePath points to a directory', () => {
    const stats = { ...fsStats, isDirectory: () => true };
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(stats);

    const actual = commandOptionsSchema.safeParse({ id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22", filePath: './' });
    assert.notStrictEqual(actual.success, true);
  });

  it('validates for a correct input.', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    const actual = commandOptionsSchema.safeParse({
      id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
      filePath: 'teamsapp.zip'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails to get Teams app when app does not exists', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/appCatalogs/teamsApps?$filter=displayName eq '`) > -1) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        name: 'Test app',
        filePath: 'teamsapp.zip'
      })
    }), new CommandError('The specified Teams app does not exist'));
  });

  it('handles error when multiple Teams apps with the specified name found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/appCatalogs/teamsApps?$filter=displayName eq '`) > -1) {
        return {
          "value": [
            {
              "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
              "displayName": "Test app"
            },
            {
              "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
              "displayName": "Test app"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        name: 'Test app',
        filePath: 'teamsapp.zip'
      })
    }), new CommandError('Multiple Teams apps with name Test app found. Found: e3e29acb-8c79-412b-b746-e6c39ff4cd22, 5b31c38c-2584-42f0-aa47-657fb3a84230.'));
  });

  it('handles selecting single result when multiple Teams apps found with the specified name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/appCatalogs/teamsApps?$filter=displayName eq '`) > -1) {
        return {
          "value": [
            {
              "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
              "displayName": "Test app"
            },
            {
              "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
              "displayName": "Test app"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: '5b31c38c-2584-42f0-aa47-657fb3a84230' });

    let updateTeamsAppCalled = false;
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/5b31c38c-2584-42f0-aa47-657fb3a84230`) {
        updateTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    await command.action(logger, { options: commandOptionsSchema.parse({ filePath: 'teamsapp.zip', name: 'Test app' }) });
    assert(updateTeamsAppCalled);
  });

  it('update Teams app in the tenant app catalog by id', async () => {
    let updateTeamsAppCalled = false;
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        updateTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    await command.action(logger, { options: commandOptionsSchema.parse({ filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` }) });
    assert(updateTeamsAppCalled);
  });

  it('update Teams app in the tenant app catalog by id (debug)', async () => {
    let updateTeamsAppCalled = false;

    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        updateTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` }) });
    assert(updateTeamsAppCalled);
  });

  it('update Teams app in the tenant app catalog by name (debug)', async () => {
    let updateTeamsAppCalled = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/appCatalogs/teamsApps?$filter=displayName eq '`) > -1) {
        return {
          "value": [
            {
              "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
              "displayName": "Test app"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        updateTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        filePath: 'teamsapp.zip',
        name: 'Test app'
      })
    });
    assert(updateTeamsAppCalled);
  });

  it('correctly handles error when updating an app', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };
    sinon.stub(request, 'put').rejects(error);

    sinon.stub(fs, 'readFileSync').returns('123');
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` }) }), new CommandError('An error has occurred'));
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'e3e29acb-8c79-412b-b746-e6c39ff4cd22',
      filePath: 'teamsapp.zip',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
