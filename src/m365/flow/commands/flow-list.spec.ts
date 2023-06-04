import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../telemetry';
import auth from '../../../Auth';
import { Logger } from '../../../cli/Logger';
import Command, { CommandError } from '../../../Command';
import { CommandInfo } from '../../../cli/CommandInfo';
import request from '../../../request';
import { pid } from '../../../utils/pid';
import { session } from '../../../utils/session';
import { sinonUtil } from '../../../utils/sinonUtil';
import commands from '../commands';
import { Cli } from '../../../cli/Cli';
const command: Command = require('./flow-list');

describe(commands.LIST, () => {
  const environmentId = 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5';
  const flowResponse = { value: [{ "name": "1c6ee23a-a835-44bc-a4f5-462b658efc13", "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/1c6ee23a-a835-44bc-a4f5-462b658efc13", "type": "Microsoft.ProcessSimple/environments/flows", "properties": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows", "displayName": "Send myself a reminder in 10 minutes", "triggerSchema": { "type": "object", "required": [], "properties": {} }, "state": "Started", "createdTime": "2018-03-23T17:58:41.4590149Z", "lastModifiedTime": "2018-03-23T17:58:41.4596534Z", "templateName": "2ec8fd1095d711e69e6b05429ec0d0d7", "environment": { "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "type": "Microsoft.ProcessSimple/environments", "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5" }, "definitionSummary": { "triggers": [{ "type": "Request", "kind": "Button" }], "actions": [{ "type": "Wait" }, { "type": "ApiConnection", "swaggerOperationId": "SendNotification", "metadata": { "flowSystemMetadata": { "swaggerOperationId": "SendNotification" } } }], "description": "Use this template to send yourself a custom delayed reminder which can be triggered with a button tap - for example, when you are close to completing a meeting or when you step into the office." }, "creator": { "tenantId": "d87a7535-dd31-4437-bfe1-95340acd55c5", "objectId": "da8f7aea-cf43-497f-ad62-c2feae89a194", "userId": "da8f7aea-cf43-497f-ad62-c2feae89a194", "userType": "ActiveDirectory" }, "provisioningMethod": "FromDefinition", "flowFailureAlertSubscribed": true } }, { "name": "3989cb59-ce1a-4a5c-bb78-257c5c39381d", "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d", "type": "Microsoft.ProcessSimple/environments/flows", "properties": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows", "displayName": "Get a daily digest of the top CNN news", "state": "Started", "createdTime": "2018-03-23T17:59:35.4407282Z", "lastModifiedTime": "2018-03-23T17:59:37.1164508Z", "templateName": "a04de6ce52984b3db0b907f588994bc8", "environment": { "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "type": "Microsoft.ProcessSimple/environments", "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5" }, "definitionSummary": { "triggers": [{ "type": "Recurrence" }], "actions": [{ "type": "If" }, { "type": "Query" }, { "type": "ApiConnection", "swaggerOperationId": "ListFeedItems", "metadata": { "flowSystemMetadata": { "swaggerOperationId": "ListFeedItems" } } }, { "type": "Foreach" }, { "type": "ApiConnection", "swaggerOperationId": "SendEmailNotification", "metadata": { "flowSystemMetadata": { "swaggerOperationId": "SendEmailNotification" } } }, { "type": "Compose" }], "description": "Each day, get an email with a list of all of the top CNN posts from the last day." }, "creator": { "tenantId": "d87a7535-dd31-4437-bfe1-95340acd55c5", "objectId": "da8f7aea-cf43-497f-ad62-c2feae89a194", "userId": "da8f7aea-cf43-497f-ad62-c2feae89a194", "userType": "ActiveDirectory" }, "provisioningMethod": "FromDefinition", "flowFailureAlertSubscribed": true } }] };
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
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
    const actual = await command.validate({ options: { environmentName: environmentId, asAdmin: true, sharingStatus: 'all' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if sharingStatus is not a valid sharingstatus', async () => {
    const actual = await command.validate({ options: { environmentName: environmentId, sharingStatus: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if sharingStatus is valid', async () => {
    const actual = await command.validate({ options: { environmentName: environmentId, sharingStatus: 'all' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if asAdmin is passed', async () => {
    const actual = await command.validate({ options: { environmentName: environmentId, asAdmin: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves available Flows (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            "value": [
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
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } });
    assert(loggerLogSpy.calledWith([
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
        },
        displayName: "Send myself a reminder in 10 minutes"
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
        },
        displayName: "Get a daily digest of the top CNN news"
      }
    ]));
  });

  it('retrieves available Flows', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            "value": [
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
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } });
    assert(loggerLogSpy.calledWith([
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
        },
        displayName: 'Send myself a reminder in 10 minutes'
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
        },
        displayName: 'Get a daily digest of the top CNN news'
      }
    ]));
  });

  it('retrieves available Flows when specifying sharingStatus ownedByMe', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${environmentId}/flows?api-version=2016-11-01`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentId, sharingStatus: 'ownedByMe' } });
    assert(loggerLogSpy.calledWith(flowResponse.value));
  });

  it('retrieves available Flows when specifying sharingStatus all', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${environmentId}/flows?api-version=2016-11-01&$filter=search('team')`) {
        return flowResponse;
      }
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${environmentId}/flows?api-version=2016-11-01&$filter=search('personal')`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentId, sharingStatus: 'all' } });
    assert(getStub.calledTwice);
  });

  it('retrieves available Flows when specifying sharingStatus personal', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${environmentId}/flows?api-version=2016-11-01&$filter=search('personal')`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentId, sharingStatus: 'personal' } });
    assert(loggerLogSpy.calledWith(flowResponse.value));
  });

  it('retrieves available Flows when specifying sharingStatus sharedWithMe', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${environmentId}/flows?api-version=2016-11-01&$filter=search('team')`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentId, sharingStatus: 'sharedWithMe' } });
    assert(loggerLogSpy.calledWith(flowResponse.value));
  });

  it('retrieves available Flows in pages', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('skiptoken') === -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            "nextLink": "https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01&%24skiptoken=eyJuZXh0TWFya2VyIjoiMjAxOTAyMDRUMTg1NDU2Wi02YTA5NGQwMi02NDFhLTQ4OTEtYjRkZi00NDA1OTRmMjZjODUifQ%3d%3d",
            "value": [
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
              }
            ]
          };
        }
      }
      else {
        return {
          "value": [
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
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } });
    assert(loggerLogSpy.calledWith([
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
        },
        displayName: 'Send myself a reminder in 10 minutes'
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
        },
        displayName: 'Get a daily digest of the top CNN news'
      }
    ]));
  });

  it('retrieves available Flows as admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/scopes/admin/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            "value": [
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
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', asAdmin: true } });
    assert(loggerLogSpy.calledWith([
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
        },
        displayName: 'Send myself a reminder in 10 minutes'
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
        },
        displayName: 'Get a daily digest of the top CNN news'
      }
    ]));
  });

  it('retrieves available Flows in pages as admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('skiptoken') === -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            "nextLink": "https://emea.api.flow.microsoft.com:11777/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01&%24skiptoken=eyJuZXh0TWFya2VyIjoiMjAxOTAyMDRUMTg1NDU2Wi02YTA5NGQwMi02NDFhLTQ4OTEtYjRkZi00NDA1OTRmMjZjODUifQ%3d%3d",
            "value": [
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
              }
            ]
          };
        }
      }
      else {
        if ((opts.url as string).indexOf('https://emea.api.flow.microsoft.com:11777') > -1) {
          throw 'Invalid request';
        }

        return {
          "value": [
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
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', asAdmin: true } });
    assert(loggerLogSpy.calledWith([
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
        },
        displayName: 'Send myself a reminder in 10 minutes'
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
        },
        displayName: 'Get a daily digest of the top CNN news'
      }
    ]));
  });

  it('correctly handles no environment found', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "EnvironmentAccessDenied",
        "message": "Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied."
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } } as any),
      new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`));
  });

  it('correctly handles no Flows found', async () => {
    sinon.stub(request, 'get').resolves({ value: [] });

    await command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles no Flows found (debug)', async () => {
    sinon.stub(request, 'get').resolves({ value: [] });

    await command.action(logger, { options: { debug: true, environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } });
    assert(loggerLogToStderrSpy.calledWith('No Flows found'));
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

    await assert.rejects(command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles error when retrieving the second page of data', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('skiptoken') === -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            "nextLink": "https://management.azure.comproviders/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01&%24skiptoken=eyJuZXh0TWFya2VyIjoiMjAxOTAyMDRUMTg1NDU2Wi02YTA5NGQwMi02NDFhLTQ4OTEtYjRkZi00NDA1OTRmMjZjODUifQ%3d%3d",
            "value": [
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
              }
            ]
          };
        }
      }
      else {
        throw {
          error: {
            'odata.error': {
              code: '-1, InvalidOperationException',
              message: {
                value: 'An error has occurred'
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } } as any),
      new CommandError('An error has occurred'));
  });

  it('supports specifying environment name', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--environment') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying option to retrieve Flows as admin', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--asAdmin') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
