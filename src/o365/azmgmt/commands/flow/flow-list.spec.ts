import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../AzmgmtAuth';
const command: Command = require('./flow-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.FLOW_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.service = new Service('https://management.azure.com/');
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.FLOW_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.FLOW_LIST);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to the Azure Management Service', (done) => {
    auth.service = new Service('https://management.azure.com/');
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to the Azure Management Service first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves available Flows (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
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

    auth.service = new Service('https://management.azure.com/');
    auth.service.connected = true;
    cmdInstance.action = command.action();
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
      if (opts.url.indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
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

    auth.service = new Service('https://management.azure.com/');
    auth.service.connected = true;
    cmdInstance.action = command.action();
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
      if (opts.url.indexOf(`providers/Microsoft.ProcessSimple/scopes/admin/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
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

    auth.service = new Service('https://management.azure.com/');
    auth.service.connected = true;
    cmdInstance.action = command.action();
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
      if (opts.url.indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows?api-version=2016-11-01`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
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

    auth.service = new Service('https://management.azure.com/');
    auth.service.connected = true;
    cmdInstance.action = command.action();
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

    auth.service = new Service('https://management.azure.com/');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`)));
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

    auth.service = new Service('https://management.azure.com/');
    auth.service.connected = true;
    cmdInstance.action = command.action();
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

    auth.service = new Service('https://management.azure.com/');
    auth.service.connected = true;
    cmdInstance.action = command.action();
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

    auth.service = new Service('https://management.azure.com/');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the environment name is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('passes validation when the environment name option specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } });
    assert.equal(actual, true);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.FLOW_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.service = new Service('https://management.azure.com/');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});