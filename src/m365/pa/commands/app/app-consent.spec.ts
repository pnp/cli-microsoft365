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
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { powerPlatform } from '../../../../utils/powerPlatform';
const command: Command = require('./app-consent-set');

describe(commands.APP_CONSENT_SET, () => {
  //#region Mocked Responses
  const environment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const name = 'e0c89645-7f00-4877-a290-cbaf6e060da1';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  //#endregion

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      request.patch,
      Cli.prompt,
      powerPlatform.getDynamicsInstanceApiUrl
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
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
        environment: environment,
        name: 'invalid',
        enabled: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name specified', async () => {
    const actual = await command.validate({
      options: {
        environment: environment,
        name: name,
        enabled: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before bypassing consent for the specified Microsoft Power App when confirm option not passed', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    await command.action(logger, {
      options: {
        environment: environment,
        name: name,
        enabled: true
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts bypassing the consent for the specified Microsoft Power App when confirm option not passed and prompt not confirmed', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    const postSpy = sinon.spy(request, 'patch');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, {
      options: {
        environment: environment,
        name: name,
        enabled: true
      }
    });
    assert(postSpy.notCalled);
  });

  it('bypasses consent for the specified Microsoft Power App when prompt confirmed (debug)', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/canvasapps(${name})`) {
        return { statusCode: 200 };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        environment: environment,
        name: name,
        enabled: true
      }
    }));
  });

  it('bypasses consent for the specified Microsoft Power App without prompting when confirm specified (debug)', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/canvasapps(${name})`) {
        return { statusCode: 204 };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        environment: environment,
        name: name,
        enabled: true,
        confirm: true
      }
    }));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    const error = {
      error: {
        message: `Something went wrong bypassing the consent for the Microsoft Power App`
      }
    };

    sinon.stub(request, 'patch').callsFake(async () => {
      throw error;
    });

    await assert.rejects(command.action(logger, {
      options: {
        environment: environment,
        name: name,
        enabled: true,
        confirm: true
      }
    } as any), new CommandError(error.error.message));
  });
});
