import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./user-license-remove');

describe(commands.USER_LICENSE_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validUserId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validUserName = 'John.Doe@contoso.com';
  const validIds = "45715bb8-13f9-4bf6-927f-ef96c102d394,0118A350-71FC-4EC3-8F0C-6A1CB8867561";
  const validIdsSingle = '45715bb8-13f9-4bf6-927f-ef96c102d394';
  //#endregion

  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.USER_LICENSE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if ids is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        ids: 'Invalid GUID', userId: validUserId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        ids: validIds, userId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when userName is not a valid upn', async () => {
    const actual = await command.validate({
      options: {
        ids: validIds, userName: 'Invalid upn'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (userId)', async () => {
    const actual = await command.validate({ options: { ids: validIds, userId: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (userName)', async () => {
    const actual = await command.validate({ options: { ids: validIds, userName: validUserName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified user licenses when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        ids: validIds,
        userId: validUserId
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified user licenses when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, {
      options: {
        ids: validIds,
        userId: validUserId
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes a single user license by userId without confirmation prompt', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/users/${validUserId}/assignLicense`)) {
        return;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { userId: validUserId, ids: validIdsSingle, confirm: true } });
    assert(postSpy.called);
  });

  it('removes the specified user licenses by userName when prompt confirmed', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${validUserName}/assignLicense`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, {
      options: {
        verbose: true, userName: validUserName, ids: validIds
      }
    });
    assert(postSpy.called);
  });

  it('removes the specified user licenses by userId without confirmation prompt', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${validUserId}/assignLicense`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true, userId: validUserId, ids: validIds, confirm: true
      }
    });
    assert(postSpy.called);
  });

  it('fails when removing one license is not a valid company license', async () => {
    const error = {
      error: {
        message: 'License 0118a350-71fc-4ec3-8f0c-6a1cb8867561 does not correspond to a valid company License.'
      }
    };

    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/users/${validUserId}/assignLicense`)) {
        throw error;
      }

      throw `Invalid request ${opts.url}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true, userId: validUserId, ids: validIdsSingle, confirm: true
      }
    }), new CommandError(error.error.message));
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The license cannot be removes.'
      }
    };
    sinon.stub(request, 'post').callsFake(async () => { throw error; });

    await assert.rejects(command.action(logger, {
      options: {
        userName: validUserName, ids: validIds, confirm: true
      }
    }), new CommandError(error.error.message));
  });
});
