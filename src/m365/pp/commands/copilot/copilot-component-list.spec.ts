import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './copilot-component-list.js';
import ppCopilotGetCommand from './copilot-get.js';

describe(commands.COPILOT_COMPONENT_LIST, () => {
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validCopilotId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validCopilotName = 'CLI 365 Copilot';
  const envUrl = 'https://contoso-dev.api.crm4.dynamics.com';
  const select = 'botcomponentid,name,componenttype,schemaname,ismanaged,statecode';
  const baseUrl = `${envUrl}/api/data/v9.1/botcomponents?$select=${select}&$filter=_parentbotid_value eq ${validCopilotId}`;

  const componentsResponse = {
    value: [
      {
        botcomponentid: '0f1b6f8c-5d3e-4c0a-9d0b-6c1a7a1f2e01',
        name: 'Greeting',
        componenttype: 9,
        schemaname: 'new_bot_23f5f58697fd43d595eb451c9797a53d.topic.Greeting',
        ismanaged: false,
        statecode: 0
      },
      {
        botcomponentid: '5b0c2f55-8a4b-4d83-b0f4-3a8f4e2b7c02',
        name: 'HR policies',
        componenttype: 16,
        schemaname: 'new_bot_23f5f58697fd43d595eb451c9797a53d.knowledge.HRPolicies',
        ismanaged: false,
        statecode: 0
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      powerPlatform.getDynamicsInstanceApiUrl,
      cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.COPILOT_COMPONENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['botcomponentid', 'name', 'componenttype', 'schemaname', 'ismanaged', 'statecode']);
  });

  it('fails validation if copilotId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment, copilotId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither copilotId nor copilotName is specified', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both copilotId and copilotName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment, copilotId: validCopilotId, copilotName: validCopilotName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if type is not a known component type', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment, copilotId: validCopilotId, type: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment, copilotId: validCopilotId, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with copilotId', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment, copilotId: validCopilotId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with copilotName and type', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment, copilotName: validCopilotName, type: 'knowledgeSource' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for every documented component type', () => {
    const types = ['topic', 'skill', 'variable', 'entity', 'dialog', 'trigger', 'languageUnderstanding', 'languageGeneration', 'dialogSchema', 'topicV2', 'translationsV2', 'entityV2', 'variableV2', 'skillV2', 'fileAttachment', 'customGpt', 'knowledgeSource', 'externalTrigger', 'settings', 'testCase', 'customMetricDefinition'];
    for (const type of types) {
      const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment, copilotId: validCopilotId, type });
      assert.strictEqual(actual.success, true, type);
    }
  });

  it('lists the components of a copilot specified by id', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(envUrl);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === baseUrl) {
        return componentsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ environmentName: validEnvironment, copilotId: validCopilotId }) });
    assert(loggerLogSpy.calledOnceWithExactly(componentsResponse.value));
  });

  it('filters the components by type', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(envUrl);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${baseUrl} and componenttype eq 16`) {
        return { value: [componentsResponse.value[1]] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ environmentName: validEnvironment, copilotId: validCopilotId, type: 'knowledgeSource' }) });
    assert(loggerLogSpy.calledOnceWithExactly([componentsResponse.value[1]]));
  });

  it('filters on the value 20 for the customMetricDefinition type', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(envUrl);
    const getStub = sinon.stub(request, 'get').resolves({ value: [] });

    await command.action(logger, { options: commandOptionsSchema.parse({ environmentName: validEnvironment, copilotId: validCopilotId, type: 'customMetricDefinition' }) });
    assert.strictEqual(getStub.lastCall.args[0].url, `${baseUrl} and componenttype eq 20`);
  });

  it('follows the next link to retrieve all components', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(envUrl);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === baseUrl) {
        return { value: [componentsResponse.value[0]], '@odata.nextLink': `${baseUrl}&$skiptoken=page2` };
      }

      if (opts.url === `${baseUrl}&$skiptoken=page2`) {
        return { value: [componentsResponse.value[1]] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ environmentName: validEnvironment, copilotId: validCopilotId }) });
    assert(loggerLogSpy.calledOnceWithExactly(componentsResponse.value));
  });

  it('resolves the copilot by name using pp copilot get and passes asAdmin on', async () => {
    const dynamicsStub = sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(envUrl);
    const executeStub = sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (cmd): Promise<any> => {
      if (cmd === ppCopilotGetCommand) {
        return { stdout: JSON.stringify({ botid: validCopilotId, name: validCopilotName }) };
      }

      throw new CommandError('Unknown case');
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === baseUrl) {
        return componentsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, environmentName: validEnvironment, copilotName: validCopilotName, asAdmin: true }) });

    const getOptions = (executeStub.lastCall.args[1] as any).options;
    assert.strictEqual(getOptions.name, validCopilotName);
    assert.strictEqual(getOptions.asAdmin, true);
    assert(dynamicsStub.calledWith(validEnvironment, true));
    assert(loggerLogSpy.calledWith(componentsResponse.value));
  });

  it('throws when the copilot name cannot be resolved', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(envUrl);
    sinon.stub(cli, 'executeCommandWithOutput').rejects(new CommandError(`The specified copilot '${validCopilotName}' does not exist.`));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ environmentName: validEnvironment, copilotName: validCopilotName }) }),
      new CommandError(`The specified copilot '${validCopilotName}' does not exist.`));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(envUrl);
    sinon.stub(request, 'get').rejects({ error: { message: 'An error has occurred' } });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ environmentName: validEnvironment, copilotId: validCopilotId }) }),
      new CommandError('An error has occurred'));
  });
});
