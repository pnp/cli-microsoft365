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
  const environment = 'Default-4be50206-9576-4237-8b17-38d8aadfaa36';
  const name = 'e0c89645-7f00-4877-a290-cbaf6e060da1';
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
      request.post,
      Cli.prompt,
      powerPlatform.getDynamicsInstanceApiUrl
    ]);
  });

  after(() => {
    sinon.restore();
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
        bypass: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name specified', async () => {
    const actual = await command.validate({
      options: {
        environment: environment,
        name: name,
        bypass: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before bypassing consent for the specified Microsoft Power App when confirm option not passed', async () => {
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
    const postSpy = sinon.spy(request, 'post');
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
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${environment}/apps/${name}/setPowerAppConnectionDirectConsentBypass?api-version=2021-02-01`) {
        return { statusCode: 204 };
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
        bypass: true
      }
    }));
  });

  it('bypasses consent for the specified Microsoft Power App without prompting when confirm specified', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${environment}/apps/${name}/setPowerAppConnectionDirectConsentBypass?api-version=2021-02-01`) {
        return { statusCode: 204 };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environment: environment,
        name: name,
        bypass: true,
        confirm: true
      }
    }));
  });

  it('correctly handles API OData error', async () => {
    const error = {
      error: {
        message: `Something went wrong bypassing the consent for the Microsoft Power App`
      }
    };

    sinon.stub(request, 'post').callsFake(async () => {
      throw error;
    });

    await assert.rejects(command.action(logger, {
      options: {
        environment: environment,
        name: name,
        bypass: true,
        confirm: true
      }
    } as any), new CommandError(error.error.message));
  });
});
