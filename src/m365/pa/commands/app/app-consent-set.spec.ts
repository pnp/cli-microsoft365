import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-consent-set.js';

describe(commands.APP_CONSENT_SET, () => {
  //#region Mocked Responses
  const environmentName = 'Default-4be50206-9576-4237-8b17-38d8aadfaa36';
  const name = 'e0c89645-7f00-4877-a290-cbaf6e060da1';
  //#endregion

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_CONSENT_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the name is not valid GUID', async () => {
    const actual = await command.validate({
      options: {
        environmentName: environmentName,
        name: 'invalid',
        bypass: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name specified', async () => {
    const actual = await command.validate({
      options: {
        environmentName: environmentName,
        name: name,
        bypass: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before bypassing consent for the specified Microsoft Power App when force option not passed', async () => {
    await command.action(logger, {
      options: {
        environmentName: environmentName,
        name: name,
        bypass: true
      }
    });

    assert(promptIssued);
  });

  it('aborts bypassing the consent for the specified Microsoft Power App when force option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        environmentName: environmentName,
        name: name,
        bypass: true
      }
    });
    assert(postSpy.notCalled);
  });

  it('bypasses consent for the specified Microsoft Power App when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${environmentName}/apps/${name}/setPowerAppConnectionDirectConsentBypass?api-version=2021-02-01`) {
        return { statusCode: 204 };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        environmentName: environmentName,
        name: name,
        bypass: true
      }
    }));
  });

  it('bypasses consent for the specified Microsoft Power App without prompting when confirm specified', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${environmentName}/apps/${name}/setPowerAppConnectionDirectConsentBypass?api-version=2021-02-01`) {
        return { statusCode: 204 };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environmentName: environmentName,
        name: name,
        bypass: true,
        force: true
      }
    }));
  });

  it('correctly handles API OData error', async () => {
    const error = {
      error: {
        message: `Something went wrong bypassing the consent for the Microsoft Power App`
      }
    };

    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        environmentName: environmentName,
        name: name,
        bypass: true,
        force: true
      }
    } as any), new CommandError(error.error.message));
  });
});
