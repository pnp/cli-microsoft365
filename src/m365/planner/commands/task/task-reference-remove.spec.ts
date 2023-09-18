import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './task-reference-remove.js';

describe(commands.TASK_REFERENCE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;
  const validTaskId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validUrl = 'https://www.microsoft.com';
  const validAlias = 'Test';

  const referenceResponse = {
    "https%3A//www%2Emicrosoft%2Ecom": {
      "alias": "Test",
      "type": "Word"
    }
  };

  const multiReferencesResponseNoAlias = {
    "https%3A//www%2Emicrosoft%2Ecom": {
      "type": "Word"
    },
    "https%3A//www%2Emicrosoft2%2Ecom": {
      "type": "Word"
    }
  };

  const multiReferencesResponse = {
    "https%3A//www%2Emicrosoft%2Ecom": {
      "alias": "Test",
      "type": "Word"
    },
    "https%3A//www%2Emicrosoft2%2Ecom": {
      "alias": "Test",
      "type": "Word"
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
    sinon.stub(Cli.getInstance().config, 'all').value({});
    commandInfo = Cli.getCommandInfo(command);
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

    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(true);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      Cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_REFERENCE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if url does not contain http or https', async () => {
    const actual = await command.validate({
      options: {
        taskId: validTaskId,
        url: 'www.microsoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid url with http specified', async () => {
    const actual = await command.validate({
      options: {
        taskId: validTaskId,
        url: 'http://www.microsoft.com'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid url with https specified', async () => {
    const actual = await command.validate({
      options: {
        taskId: validTaskId,
        url: 'https://www.microsoft.com'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removal when confirm option not passed', async () => {
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        taskId: validTaskId,
        url: validUrl
      }
    });


    assert(promptIssued);
  });

  it('passes validation when valid options specified', async () => {
    const actual = await command.validate({
      options: {
        taskId: validTaskId,
        url: validUrl
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly removes reference', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return { references: null };
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      taskId: validTaskId,
      url: validUrl,
      force: true
    };

    await command.action(logger, { options: options } as any);
  });

  it('correctly removes reference by alias with prompting', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return { references: null };
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return {
          "@odata.etag": "TestEtag",
          references: referenceResponse
        };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      taskId: validTaskId,
      alias: validAlias
    };

    await command.action(logger, { options: options } as any);
  });

  it('fails validation when no references found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return {
          "@odata.etag": "TestEtag",
          references: {}
        };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        taskId: validTaskId,
        alias: validAlias
      }
    }), new CommandError(`The specified reference with alias ${validAlias} does not exist`));
  });

  it('fails validation when reference does not contain alias', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return {
          "@odata.etag": "TestEtag",
          references: multiReferencesResponseNoAlias
        };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        taskId: validTaskId,
        alias: validAlias
      }
    }), new CommandError(`The specified reference with alias ${validAlias} does not exist`));
  });

  it('fails validation when multiple references found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return {
          "@odata.etag": "TestEtag",
          references: multiReferencesResponse
        };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        taskId: validTaskId,
        alias: validAlias
      }
    }), new CommandError(`Multiple references with alias ${validAlias} found. Pass one of the following urls within the "--url" option : https://www.microsoft.com,https://www.microsoft2.com`));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
