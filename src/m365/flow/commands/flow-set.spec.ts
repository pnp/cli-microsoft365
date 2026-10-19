import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { CommandError } from '../../../Command.js';
import { cli } from '../../../cli/cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import request from '../../../request.js';
import { telemetry } from '../../../telemetry.js';
import { accessToken } from '../../../utils/accessToken.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import commands from '../commands.js';
import command, { options } from './flow-set.js';

describe(commands.SET, () => {
  const environmentName = 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5';
  const flowName = '3989cb59-ce1a-4a5c-bb78-257c5c39381d';
  const definition = JSON.stringify({
    properties: {
      definition: {
        $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
        actions: {},
        parameters: {},
        triggers: {},
        contentVersion: '1.0.0.0',
        outputs: {}
      },
      connectionReferences: {},
      displayName: 'Test Flow',
      environment: { name: environmentName }
    }
  });
  const baseUrl = `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}`;

  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch,
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the name is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environmentName,
      name: 'invalid',
      definition: definition
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if definition is not valid JSON', async () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environmentName,
      name: flowName,
      definition: 'not-valid-json'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when required options are specified correctly', async () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environmentName,
      name: flowName,
      definition: definition
    });
    assert.strictEqual(actual.success, true);
  });

  it('updates the flow without prompting when force specified and no warnings or errors', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}?api-version=2016-11-01`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        environmentName: environmentName,
        name: flowName,
        definition: definition,
        force: true
      }
    });
    assert(loggerLogToStderrSpy.notCalled);
  });

  it('strips flow get metadata from definition before sending PATCH request', async () => {
    const flowGetOutput = JSON.stringify({
      name: flowName,
      id: `/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}`,
      type: 'Microsoft.ProcessSimple/environments/flows',
      displayName: 'Test Flow',
      description: 'A test flow',
      triggers: 'Recurrence',
      actions: 'Http-HttpRequest',
      properties: {
        displayName: 'Test Flow',
        definition: {
          $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
          actions: {},
          parameters: {},
          triggers: {},
          contentVersion: '1.0.0.0',
          outputs: {}
        },
        connectionReferences: {
          ['shared_sharepointonline']: {
            connectionName: 'shared_sharepointonline',
            id: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
            operationDefinitions: {
              GetItems: { summary: 'Get items', description: 'Gets SharePoint items' }
            },
            apiDefinition: {
              name: 'shared_sharepointonline',
              id: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline'
            }
          }
        },
        environment: { name: environmentName }
      }
    });

    let patchBody: any;
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}?api-version=2016-11-01`) {
        patchBody = opts.data;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        environmentName: environmentName,
        name: flowName,
        definition: flowGetOutput,
        force: true
      }
    });

    assert.strictEqual(patchBody.displayName, undefined);
    assert.strictEqual(patchBody.description, undefined);
    assert.strictEqual(patchBody.triggers, undefined);
    assert.strictEqual(patchBody.actions, undefined);
    assert.strictEqual(patchBody.properties.connectionReferences.shared_sharepointonline.operationDefinitions, undefined);
    assert.strictEqual(patchBody.properties.connectionReferences.shared_sharepointonline.apiDefinition, undefined);
    assert.strictEqual(patchBody.properties.connectionReferences.shared_sharepointonline.connectionName, 'shared_sharepointonline');
  });

  it('updates the flow and publishes it when publish and force specified', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}?api-version=2016-11-01`) {
        return;
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}/publish?api-version=2016-11-01`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        environmentName: environmentName,
        name: flowName,
        definition: definition,
        publish: true,
        force: true
      }
    });

    assert(patchStub.calledOnce);
    assert(postStub.calledOnce);
  });

  it('updates the flow (debug) when no issues found', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}?api-version=2016-11-01`) {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}/checkFlowErrors?api-version=2016-11-01`) {
        return [];
      }
      if (opts.url === `${baseUrl}/checkFlowWarnings?api-version=2016-11-01`) {
        return [];
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: environmentName,
        name: flowName,
        definition: definition
      }
    });

    assert(loggerLogToStderrSpy.called);
  });

  it('throws error when errors found', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}/checkFlowErrors?api-version=2016-11-01`) {
        return [{ error: { code: 'InvalidDefinition', message: 'The flow definition is invalid.' } }];
      }
      if (opts.url === `${baseUrl}/checkFlowWarnings?api-version=2016-11-01`) {
        return [];
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: environmentName,
          name: flowName,
          definition: definition
        }
      }),
      new CommandError(`The flow definition has the following errors:\n  - The flow definition is invalid.`)
    );
  });

  it('throws error with JSON-formatted details when errors found and output is json', async () => {
    const errors = [{ error: { code: 'InvalidDefinition', message: 'The flow definition is invalid.' } }];
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}/checkFlowErrors?api-version=2016-11-01`) {
        return errors;
      }
      if (opts.url === `${baseUrl}/checkFlowWarnings?api-version=2016-11-01`) {
        return [];
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: environmentName,
          name: flowName,
          definition: definition,
          output: 'json'
        }
      }),
      new CommandError(`The flow definition has the following errors:\n${JSON.stringify(errors, null, 2)}`)
    );
  });

  it('falls back to raw JSON in error details when an error entry has no error.message', async () => {
    const errors = [{ code: 'InvalidDefinition' } as any];
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}/checkFlowErrors?api-version=2016-11-01`) {
        return errors;
      }
      if (opts.url === `${baseUrl}/checkFlowWarnings?api-version=2016-11-01`) {
        return [];
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: environmentName,
          name: flowName,
          definition: definition
        }
      }),
      new CommandError(`The flow definition has the following errors:\n  - ${JSON.stringify(errors[0])}`)
    );
  });

  it('prompts when warnings found and aborts when not confirmed', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async () => {
      throw 'Should not be called';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}/checkFlowErrors?api-version=2016-11-01`) {
        return [];
      }
      if (opts.url === `${baseUrl}/checkFlowWarnings?api-version=2016-11-01`) {
        return [{ error: { code: 'Warning', message: 'Connection is deprecated.' } }];
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        environmentName: environmentName,
        name: flowName,
        definition: definition
      }
    });

    assert(patchStub.notCalled);
  });

  it('prompts with JSON-formatted details when warnings found and output is json', async () => {
    const warnings = [{ error: { code: 'Warning', message: 'Connection is deprecated.' } }];
    sinon.stub(request, 'patch').callsFake(async () => {
      throw 'Should not be called';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}/checkFlowErrors?api-version=2016-11-01`) {
        return [];
      }
      if (opts.url === `${baseUrl}/checkFlowWarnings?api-version=2016-11-01`) {
        return warnings;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    const promptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        environmentName: environmentName,
        name: flowName,
        definition: definition,
        output: 'json'
      }
    });

    assert.strictEqual(
      promptStub.firstCall.args[0].message,
      `The flow definition has the following warnings:\n\n${JSON.stringify(warnings, null, 2)}\n\nDo you want to proceed with the update?`
    );
  });

  it('falls back to raw JSON in warning details when a warning entry has no error.message', async () => {
    const warnings = [{ code: 'Warning' } as any];
    sinon.stub(request, 'patch').callsFake(async () => {
      throw 'Should not be called';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}/checkFlowErrors?api-version=2016-11-01`) {
        return [];
      }
      if (opts.url === `${baseUrl}/checkFlowWarnings?api-version=2016-11-01`) {
        return warnings;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    const promptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        environmentName: environmentName,
        name: flowName,
        definition: definition
      }
    });

    assert.strictEqual(
      promptStub.firstCall.args[0].message,
      `The flow definition has the following warnings:\n\n  - ${JSON.stringify(warnings[0])}\n\nDo you want to proceed with the update?`
    );
  });

  it('prompts when warnings found and updates the flow when confirmed', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}?api-version=2016-11-01`) {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}/checkFlowErrors?api-version=2016-11-01`) {
        return [];
      }
      if (opts.url === `${baseUrl}/checkFlowWarnings?api-version=2016-11-01`) {
        return [{ error: { code: 'Warning', message: 'Connection is deprecated.' } }];
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        environmentName: environmentName,
        name: flowName,
        definition: definition
      }
    });

    assert(patchStub.calledOnce);
  });

  it('throws error when both errors and warnings found', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}/checkFlowErrors?api-version=2016-11-01`) {
        return [{ error: { code: 'InvalidDefinition', message: 'The flow definition is invalid.' } }];
      }
      if (opts.url === `${baseUrl}/checkFlowWarnings?api-version=2016-11-01`) {
        return [{ error: { code: 'Warning', message: 'Connection is deprecated.' } }];
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: environmentName,
          name: flowName,
          definition: definition
        }
      }),
      new CommandError(`The flow definition has the following errors:\n  - The flow definition is invalid.`)
    );
  });

  it('throws error when check endpoints fail', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${baseUrl}?api-version=2016-11-01`) {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').rejects(new Error('Check endpoint not available'));

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: environmentName,
          name: flowName,
          definition: definition
        }
      })
    );

    assert(patchStub.notCalled);
  });

  it('correctly handles error when updating the flow', async () => {
    sinon.stub(request, 'patch').rejects({
      error: {
        code: 'FlowNotFound',
        message: `The flow '${flowName}' is not found.`
      }
    });

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: environmentName,
          name: flowName,
          definition: definition,
          force: true
        }
      }),
      new CommandError(`The flow '${flowName}' is not found.`)
    );
  });

  it('correctly handles environment access denied error', async () => {
    sinon.stub(request, 'patch').rejects({
      error: {
        code: 'EnvironmentAccessDenied',
        message: `Access to the environment '${environmentName}' is denied.`
      }
    });

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: environmentName,
          name: flowName,
          definition: definition,
          force: true
        }
      }),
      new CommandError(`Access to the environment '${environmentName}' is denied.`)
    );
  });
});
