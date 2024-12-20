import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './website-webrole-remove.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { CommandError } from '../../../../Command.js';

describe(commands.WEBSITE_WEBROLE_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c';
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validWebsiteId = '3bbc8102-8ee7-4dac-afbb-807cc5b6f9c2';
  const validWebsiteName = 'CLI 365 PowerPageSite';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  //#endregion

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      request.get,
      powerPlatform.getDynamicsInstanceApiUrl,
      cli.promptForConfirmation,
      cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.WEBSITE_WEBROLE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        id: 'Invalid GUID',
        websiteId: validWebsiteId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if websiteId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        id: validId,
        websiteId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (id and websiteId)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, id: validId, websiteId: validWebsiteId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (id and websiteName)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, id: validId, websiteName: validWebsiteName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified webrole owned by the currently signed-in user when force option not passed', async () => {
    await command.action(logger, {
      options: {
        environmentName: validEnvironment,
        id: validId,
        websiteId: validWebsiteId
      }
    });

    assert(promptIssued);
  });

  it('aborts removing the specified webrole owned by the currently signed-in user when force option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    await command.action(logger, {
      options: {
        environmentName: validEnvironment,
        id: validId,
        websiteId: validWebsiteId
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified webrole owned by the currently signed-in user when prompt confirmed by id and websiteId', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webroles?$filter=_mspp_websiteid_value eq ${validWebsiteId} and mspp_webroleid eq ${validId}&$select=mspp_webroleid`) {
        return {
          "value": [
            {
              "mspp_webroleid": validId
            }
          ]
        };
      }
      throw `The specified webrole '${validId}' does not exist for the specified website.`;
    });
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webroles(${validId})`) {
        return;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, {
      options: {
        verbose: true,
        environmentName: validEnvironment,
        id: validId,
        websiteId: validWebsiteId
      }
    });
    assert(loggerLogToStderrSpy.called);
  });



  it('removes the specified webrole without confirmation prompt', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webroles?$filter=_mspp_websiteid_value eq ${validWebsiteId} and mspp_webroleid eq ${validId}&$select=mspp_webroleid`) {
        return {
          "value": [
            {
              "mspp_webroleid": validId
            }
          ]
        };
      }
      throw `The specified webrole '${validId}' does not exist for the specified website.`;
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webroles(${validId})`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        id: validId,
        websiteId: validWebsiteId,
        force: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified webrole owned by the currently signed-in user when prompt confirmed by id and websiteName', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);


    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/powerpagesites?$filter=name eq '${validWebsiteName}'&$select=powerpagesiteid`) {
        return {
          "value": [
            {
              "powerpagesiteid": validWebsiteId
            }
          ]
        };
      }

      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webroles?$filter=_mspp_websiteid_value eq ${validWebsiteId} and mspp_webroleid eq ${validId}&$select=mspp_webroleid`) {
        return {
          "value": [
            {
              "mspp_webroleid": validId
            }
          ]
        };
      }
      throw `The specified webrole '${validId}' does not exist for the specified website.`;
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webroles(${validId})`) {
        return;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, {
      options: {
        verbose: true,
        environmentName: validEnvironment,
        id: validId,
        websiteName: validWebsiteName
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('fails validation on unable to find website based on websiteName', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    const EmptyWebsiteResponse = {
      value: [
      ]
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/powerpagesites?$filter=name eq '${validWebsiteName}'&$select=powerpagesiteid`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return EmptyWebsiteResponse;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        environmentName: validEnvironment,
        id: validId,
        websiteName: validWebsiteName
      }
    }), new CommandError(`The specified website '${validWebsiteName}' does not exist.`));
  });

  it('fails validation on unable to find webrole based on website', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    const EmptyWebsiteWebRoleResponse = {
      value: [
      ]
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webroles?$filter=_mspp_websiteid_value eq ${validWebsiteId} and mspp_webroleid eq ${validId}&$select=mspp_webroleid`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return EmptyWebsiteWebRoleResponse;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        environmentName: validEnvironment,
        id: validId,
        websiteId: validWebsiteId
      }
    }), new CommandError(`The specified webrole '${validId}' does not exist for the specified website.`));
  });


  it('correctly handles API OData error', async () => {
    const errorMessage = 'post_request_failed: Post request failed from the network, could be a 4xx/5xx or a network unavailability. Please check the exact error code for details. invalid_grant';

    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'delete').callsFake(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        id: validId,
        websiteId: validWebsiteId,
        force: true
      }
    }), new CommandError(errorMessage));
  });
});