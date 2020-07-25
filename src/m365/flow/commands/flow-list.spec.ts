import commands from '../commands';
import Command, { CommandOption, CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
const command: Command = require('./flow-list');
import * as assert from 'assert';
import request from '../../../request';
import Utils from '../../../Utils';

describe(commands.FLOW_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FLOW_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves available Flows (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
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
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: '1c6ee23a-a835-44bc-a4f5-462b658efc13',
            displayName: 'Send myself a reminder in 10 minutes'
          },
          {
            name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d',
            displayName: 'Get a daily digest of the top CNN news'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves available Flows', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
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
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: '1c6ee23a-a835-44bc-a4f5-462b658efc13',
            displayName: 'Send myself a reminder in 10 minutes'
          },
          {
            name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d',
            displayName: 'Get a daily digest of the top CNN news'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves available Flows in pages', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('skiptoken') === -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
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
          });
        }
      }
      else {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: '1c6ee23a-a835-44bc-a4f5-462b658efc13',
            displayName: 'Send myself a reminder in 10 minutes'
          },
          {
            name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d',
            displayName: 'Get a daily digest of the top CNN news'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves available Flows as admin', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/scopes/admin/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
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
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', asAdmin: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: '1c6ee23a-a835-44bc-a4f5-462b658efc13',
            displayName: 'Send myself a reminder in 10 minutes'
          },
          {
            name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d',
            displayName: 'Get a daily digest of the top CNN news'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves available Flows in pages as admin', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('skiptoken') === -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
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
          });
        }
      }
      else {
        if ((opts.url as string).indexOf('https://emea.api.flow.microsoft.com:11777') > -1) {
          return Promise.reject('Invalid request');
        }

        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', asAdmin: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: '1c6ee23a-a835-44bc-a4f5-462b658efc13',
            displayName: 'Send myself a reminder in 10 minutes'
          },
          {
            name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d',
            displayName: 'Get a daily digest of the top CNN news'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs all properties when output is JSON', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
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
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
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
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no environment found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no Flows found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.resolve({ value: [] });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no Flows found (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.resolve({ value: [] });
    });

    cmdInstance.action({ options: { debug: true, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('No Flows found'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving the second page of data', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('skiptoken') === -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
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
          });
        }
      }
      else {
        return Promise.reject({
          error: {
            'odata.error': {
              code: '-1, InvalidOperationException',
              message: {
                value: 'An error has occurred'
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying environment name', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--environment') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying option to retrieve Flows as admin', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--asAdmin') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});