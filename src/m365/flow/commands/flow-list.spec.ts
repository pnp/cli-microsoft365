import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { CommandError } from '../../../Command.js';
import { cli } from '../../../cli/cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import request from '../../../request.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import commands from '../commands.js';
import command from './flow-list.js';
import { accessToken } from '../../../utils/accessToken.js';

describe(commands.LIST, () => {
  const environmentName = 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5';
  const flowResponse = { value: [{ "name": "1c6ee23a-a835-44bc-a4f5-462b658efc13", "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/1c6ee23a-a835-44bc-a4f5-462b658efc13", "type": "Microsoft.ProcessSimple/environments/flows", "properties": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows", "displayName": "Send myself a reminder in 10 minutes", "triggerSchema": { "type": "object", "required": [], "properties": {} }, "state": "Started", "createdTime": "2018-03-23T17:58:41.4590149Z", "lastModifiedTime": "2018-03-23T17:58:41.4596534Z", "templateName": "2ec8fd1095d711e69e6b05429ec0d0d7", "environment": { "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "type": "Microsoft.ProcessSimple/environments", "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5" }, "definitionSummary": { "triggers": [{ "type": "Request", "kind": "Button" }], "actions": [{ "type": "Wait" }, { "type": "ApiConnection", "swaggerOperationId": "SendNotification", "metadata": { "flowSystemMetadata": { "swaggerOperationId": "SendNotification" } } }], "description": "Use this template to send yourself a custom delayed reminder which can be triggered with a button tap - for example, when you are close to completing a meeting or when you step into the office." }, "creator": { "tenantId": "d87a7535-dd31-4437-bfe1-95340acd55c5", "objectId": "da8f7aea-cf43-497f-ad62-c2feae89a194", "userId": "da8f7aea-cf43-497f-ad62-c2feae89a194", "userType": "ActiveDirectory" }, "provisioningMethod": "FromDefinition", "flowFailureAlertSubscribed": true } }, { "name": "3989cb59-ce1a-4a5c-bb78-257c5c39381d", "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d", "type": "Microsoft.ProcessSimple/environments/flows", "properties": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows", "displayName": "Get a daily digest of the top CNN news", "state": "Started", "createdTime": "2018-03-23T17:59:35.4407282Z", "lastModifiedTime": "2018-03-23T17:59:37.1164508Z", "templateName": "a04de6ce52984b3db0b907f588994bc8", "environment": { "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "type": "Microsoft.ProcessSimple/environments", "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5" }, "definitionSummary": { "triggers": [{ "type": "Recurrence" }], "actions": [{ "type": "If" }, { "type": "Query" }, { "type": "ApiConnection", "swaggerOperationId": "ListFeedItems", "metadata": { "flowSystemMetadata": { "swaggerOperationId": "ListFeedItems" } } }, { "type": "Foreach" }, { "type": "ApiConnection", "swaggerOperationId": "SendEmailNotification", "metadata": { "flowSystemMetadata": { "swaggerOperationId": "SendEmailNotification" } } }, { "type": "Compose" }], "description": "Each day, get an email with a list of all of the top CNN posts from the last day." }, "creator": { "tenantId": "d87a7535-dd31-4437-bfe1-95340acd55c5", "objectId": "da8f7aea-cf43-497f-ad62-c2feae89a194", "userId": "da8f7aea-cf43-497f-ad62-c2feae89a194", "userType": "ActiveDirectory" }, "provisioningMethod": "FromDefinition", "flowFailureAlertSubscribed": true } }] };
  const regularFlowResponse = {
    value: [
      {
        "name": "1c6ee23a-a835-44bc-a4f5-462b658efc13",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/1c6ee23a-a835-44bc-a4f5-462b658efc13",
        "type": "Microsoft.ProcessSimple/environments/flows",
        "properties": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
          "displayName": "Send myself a reminder in 10 minutes",
          "triggerSchema": {
            "type": "object",
            "required": [],
            "properties": {}
          },
          "state": "Started",
          "createdTime": "2018-03-23T17:58:41.4590149Z",
          "lastModifiedTime": "2018-03-23T17:58:41.4596534Z",
          "templateName": "2ec8fd1095d711e69e6b05429ec0d0d7",
          "environment": {
            "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
            "type": "Microsoft.ProcessSimple/environments",
            "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5"
          },
          "definitionSummary": {
            "triggers": [
              {
                "type": "Request",
                "kind": "Button"
              }
            ],
            "actions": [
              {
                "type": "Wait"
              },
              {
                "type": "ApiConnection",
                "swaggerOperationId": "SendNotification",
                "metadata": {
                  "flowSystemMetadata": {
                    "swaggerOperationId": "SendNotification"
                  }
                }
              }
            ],
            "description": "Use this template to send yourself a custom delayed reminder which can be triggered with a button tap - for example, when you are close to completing a meeting or when you step into the office."
          },
          "creator": {
            "tenantId": "d87a7535-dd31-4437-bfe1-95340acd55c5",
            "objectId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
            "userId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
            "userType": "ActiveDirectory"
          },
          "provisioningMethod": "FromDefinition",
          "flowFailureAlertSubscribed": true
        }
      },
      {
        "name": "3989cb59-ce1a-4a5c-bb78-257c5c39381d",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d",
        "type": "Microsoft.ProcessSimple/environments/flows",
        "properties": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
          "displayName": "Get a daily digest of the top CNN news",
          "state": "Started",
          "createdTime": "2018-03-23T17:59:35.4407282Z",
          "lastModifiedTime": "2018-03-23T17:59:37.1164508Z",
          "templateName": "a04de6ce52984b3db0b907f588994bc8",
          "environment": {
            "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
            "type": "Microsoft.ProcessSimple/environments",
            "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5"
          },
          "definitionSummary": {
            "triggers": [
              {
                "type": "Recurrence"
              }
            ],
            "actions": [
              {
                "type": "If"
              },
              {
                "type": "Query"
              },
              {
                "type": "ApiConnection",
                "swaggerOperationId": "ListFeedItems",
                "metadata": {
                  "flowSystemMetadata": {
                    "swaggerOperationId": "ListFeedItems"
                  }
                }
              },
              {
                "type": "Foreach"
              },
              {
                "type": "ApiConnection",
                "swaggerOperationId": "SendEmailNotification",
                "metadata": {
                  "flowSystemMetadata": {
                    "swaggerOperationId": "SendEmailNotification"
                  }
                }
              },
              {
                "type": "Compose"
              }
            ],
            "description": "Each day, get an email with a list of all of the top CNN posts from the last day."
          },
          "creator": {
            "tenantId": "d87a7535-dd31-4437-bfe1-95340acd55c5",
            "objectId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
            "userId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
            "userType": "ActiveDirectory"
          },
          "provisioningMethod": "FromDefinition",
          "flowFailureAlertSubscribed": true
        }
      }
    ]
  };
  const adminFlowResponse = {
    value: [
      {
        "name": "fc2d4ef5-4151-4e93-9aaa-1f380d7ed95d",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-1e852b49-bf4b-4ba5-bcd4-a8c4706c8ed4/flows/fc2d4ef5-4151-4e93-9aaa-1f380d7ed95d",
        "type": "Microsoft.ProcessSimple/environments/flows",
        "properties": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
          "displayName": "Invoicing flow",
          "state": "Started",
          "createdTime": "2024-07-09T21:32:15Z",
          "lastModifiedTime": "2024-07-09T21:32:15Z",
          "flowSuspensionReason": "None",
          "environment": {
            "name": "Default-1e852b49-bf4b-4ba5-bcd4-a8c4706c8ed4",
            "type": "Microsoft.ProcessSimple/environments",
            "id": "/providers/Microsoft.ProcessSimple/environments/Default-1e852b49-bf4b-4ba5-bcd4-a8c4706c8ed4"
          },
          "definitionSummary": {
            "triggers": [],
            "actions": []
          },
          "creator": {
            "tenantId": "1e852b49-bf4b-4ba5-bcd4-a8c4706c8ed4",
            "objectId": "9723d9b0-b8d6-48af-b381-fb4a9f0403e4",
            "userId": "9723d9b0-b8d6-48af-b381-fb4a9f0403e4",
            "userType": "ActiveDirectory"
          },
          "flowFailureAlertSubscribed": false,
          "isManaged": false,
          "machineDescriptionData": {},
          "flowOpenAiData": {
            "isConsequential": false,
            "isConsequentialFlagOverwritten": false
          }
        }
      },
      {
        "name": "0eabab30-8e55-4100-911f-e7afd9ba8919",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-1e852b49-bf4b-4ba5-bcd4-a8c4706c8ed4/flows/0eabab30-8e55-4100-911f-e7afd9ba8919",
        "type": "Microsoft.ProcessSimple/environments/flows",
        "properties": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
          "displayName": "Invoicing flow 2",
          "state": "Started",
          "createdTime": "2024-07-09T21:32:15Z",
          "lastModifiedTime": "2024-07-09T21:32:15Z",
          "flowSuspensionReason": "None",
          "environment": {
            "name": "Default-1e852b49-bf4b-4ba5-bcd4-a8c4706c8ed4",
            "type": "Microsoft.ProcessSimple/environments",
            "id": "/providers/Microsoft.ProcessSimple/environments/Default-1e852b49-bf4b-4ba5-bcd4-a8c4706c8ed4"
          },
          "definitionSummary": {
            "triggers": [],
            "actions": []
          },
          "creator": {
            "tenantId": "1e852b49-bf4b-4ba5-bcd4-a8c4706c8ed4",
            "objectId": "9723d9b0-b8d6-48af-b381-fb4a9f0403e4",
            "userId": "9723d9b0-b8d6-48af-b381-fb4a9f0403e4",
            "userType": "ActiveDirectory"
          },
          "flowFailureAlertSubscribed": false,
          "isManaged": false,
          "machineDescriptionData": {},
          "flowOpenAiData": {
            "isConsequential": false,
            "isConsequentialFlagOverwritten": false
          }
        }
      }
    ]
  };
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'displayName']);
  });

  it('fails validation if asAdmin is specified in combination with a sharingStatus', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, asAdmin: true, sharingStatus: 'all' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if sharingStatus is not a valid sharingstatus', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, sharingStatus: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if sharingStatus is valid', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, sharingStatus: 'all' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if asAdmin is passed', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, asAdmin: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves available flows', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows?api-version=2016-11-01`) {
        return regularFlowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName } });
    assert(loggerLogSpy.calledOnceWithExactly(regularFlowResponse.value));
  });

  it('retrieves available Flows when specifying sharingStatus ownedByMe', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows?api-version=2016-11-01`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentName, sharingStatus: 'ownedByMe' } });
    assert(loggerLogSpy.calledOnceWithExactly(flowResponse.value));
  });

  it('retrieves available Flows when specifying sharingStatus all', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows?api-version=2016-11-01&$filter=search('personal')`) {
        return { value: [flowResponse.value[0]] };
      }
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows?api-version=2016-11-01&$filter=search('team')`) {
        return { value: [flowResponse.value[1]] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentName, sharingStatus: 'all' } });
    assert(loggerLogSpy.calledOnceWithExactly(flowResponse.value));
  });

  it('retrieves available Flows when specifying sharingStatus personal', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows?api-version=2016-11-01&$filter=search('personal')`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentName, sharingStatus: 'personal' } });
    assert(loggerLogSpy.calledOnceWithExactly(flowResponse.value));
  });

  it('retrieves available Flows when specifying sharingStatus sharedWithMe', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows?api-version=2016-11-01&$filter=search('team')`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentName, sharingStatus: 'sharedWithMe' } });
    assert(loggerLogSpy.calledOnceWithExactly(flowResponse.value));
  });

  it('retrieves available Flows as admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${environmentName}/v2/flows?api-version=2016-11-01`) {
        return adminFlowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentName, asAdmin: true, verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly(adminFlowResponse.value));
  });

  it('correctly handles no environment found', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "EnvironmentAccessDenied",
        "message": `Access to the environment '${environmentName}' is denied.`
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName } } as any),
      new CommandError(`Access to the environment '${environmentName}' is denied.`));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName } } as any),
      new CommandError('An error has occurred'));
  });

  it('retrieves flows including flows from solutions', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows?api-version=2016-11-01&include=includeSolutionCloudFlows`) {
        return {
          value: [
            {
              name: "1c6ee23a-a835-44bc-a4f5-462b658efc13",
              id: "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/1c6ee23a-a835-44bc-a4f5-462b658efc13",
              properties: {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
                "displayName": "Get a daily digest of the top CNN news"
              }
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentName, includeSolutions: true } } as any);

    assert(loggerLogSpy.calledOnceWithExactly([
      {
        name: "1c6ee23a-a835-44bc-a4f5-462b658efc13",
        id: "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/1c6ee23a-a835-44bc-a4f5-462b658efc13",
        properties: {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
          "displayName": "Get a daily digest of the top CNN news"
        }
      }
    ]));
  });

  it('retrieves flows and removes duplicates', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows?api-version=2016-11-01`) {
        return {
          value: [
            {
              name: "1c6ee23a-a835-44bc-a4f5-462b658efc13",
              id: "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/1c6ee23a-a835-44bc-a4f5-462b658efc13",
              properties: {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
                "displayName": "Get a daily digest of the top CNN news"
              }
            },
            {
              name: "1c6ee23a-a835-44bc-a4f5-462b658efc13",
              id: "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/1c6ee23a-a835-44bc-a4f5-462b658efc13",
              properties: {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
                "displayName": "Get a daily digest of the top CNN news"
              }
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } } as any);
    assert(loggerLogSpy.calledWith([
      {
        name: "1c6ee23a-a835-44bc-a4f5-462b658efc13",
        id: "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/1c6ee23a-a835-44bc-a4f5-462b658efc13",
        properties: {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
          "displayName": "Get a daily digest of the top CNN news"
        }
      }
    ]));
  });

  it('correctly transforms output for non-JSON output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows?api-version=2016-11-01`) {
        return regularFlowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, output: 'text' } });
    const expectedOutput = regularFlowResponse.value.map(f => ({ ...f, displayName: f.properties.displayName }));
    assert(loggerLogSpy.calledOnceWithExactly(expectedOutput));
  });

  it('correctly transforms output for non-JSON output as admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${environmentName}/v2/flows?api-version=2016-11-01`) {
        return adminFlowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, asAdmin: true, output: 'text' } });
    const expectedOutput = adminFlowResponse.value.map(f => ({ ...f, displayName: f.properties.displayName }));
    assert(loggerLogSpy.calledOnceWithExactly(expectedOutput));
  });
});
