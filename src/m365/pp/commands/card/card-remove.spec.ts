import assert from 'assert';
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
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './card-remove.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.CARD_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validName = 'CLI 365 Card';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const appResponse = {
    solutionid: 'fd140aae-4df4-11dd-bd17-0019b9312238',
    modifiedon: '2022-10-11T08:52:12Z',
    '_owninguser_value': '7d48edd3-69fd-ec11-82e5-000d3ab87733',
    overriddencreatedon: null,
    ismanaged: false,
    schemaversion: null,
    tags: null,
    importsequencenumber: null,
    componentidunique: 'd7c1acb5-37a4-4873-b24e-34b18c15c6a5',
    '_modifiedonbehalfby_value': null,
    componentstate: 0,
    statecode: 0,
    name: validName,
    versionnumber: 3044006,
    utcconversiontimezonecode: null,
    cardid: validId,
    publishdate: null,
    '_createdonbehalfby_value': null,
    '_modifiedby_value': '7d48edd3-69fd-ec11-82e5-000d3ab87733',
    createdon: '2022-10-11T08:52:12Z',
    overwritetime: '1900-01-01T00:00:00Z',
    '_owningbusinessunit_value': '2199f44c-195b-ec11-8f8f-000d3adca49c',
    hiddentags: null,
    description: ' ',
    appdefinition: '{\'screens\':{\'main\':{\'template\':{\'type\':\'AdaptiveCard\',\'body\':[{\'type\':\'TextBlock\',\'size\':\'Medium\',\'weight\':\'bolder\',\'text\':\'Your card title goes here\'},{\'type\':\'TextBlock\',\'text\':\'Add and remove element to customize your new card.\',\'wrap\':true}],\'actions\':[],\'$schema\':\'http://adaptivecards.io/schemas/1.4.0/adaptive-card.json\',\'version\':\'1.4\'},\'verbs\':{\'submit\':\'echo\'}}},\'sampleData\':{\'main\':{}},\'connections\':{},\'variables\':{},\'flows\':{}}',
    statuscode: 1,
    remixsourceid: null,
    sizes: null,
    '_owningteam_value': null,
    coowners: null,
    '_createdby_value': '7d48edd3-69fd-ec11-82e5-000d3ab87733',
    '_ownerid_value': '7d48edd3-69fd-ec11-82e5-000d3ab87733',
    publishsourceid: null,
    timezoneruleversionnumber: null,
    iscustomizable: {
      Value: true,
      CanBeChanged: true,
      ManagedPropertyLogicalName: 'iscustomizableanddeletable'
    },
    owninguser: {
      azureactivedirectoryobjectid: '88e85b64-e687-4e0b-bbf4-f42f5f8e574c',
      fullname: 'Contoso Admin',
      systemuserid: '7d48edd3-69fd-ec11-82e5-000d3ab87733',
      ownerid: '7d48edd3-69fd-ec11-82e5-000d3ab87733'
    }
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      powerPlatform.getDynamicsInstanceApiUrl,
      powerPlatform.getCardByName,
      cli.promptForConfirmation,
      cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CARD_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        id: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (id)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, name: validName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified card owned by the currently signed-in user when force option not passed', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    await command.action(logger, {
      options: {
        environmentName: validEnvironment,
        id: validId
      }
    });

    assert(promptIssued);
  });

  it('aborts removing the specified card owned by the currently signed-in user when force option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, {
      options: {
        environmentName: validEnvironment,
        id: validId
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified card owned by the currently signed-in user when prompt confirmed', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);
    sinon.stub(powerPlatform, 'getCardByName').resolves(appResponse);

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards(${validId})`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, {
      options: {
        debug: true,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: 'CLI 365 Card'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified card without confirmation prompt', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards(${validId})`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        id: validId,
        force: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'delete').callsFake(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        id: validId,
        force: true
      }
    }), new CommandError(errorMessage));
  });
});
