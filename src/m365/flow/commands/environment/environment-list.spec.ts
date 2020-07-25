import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./environment-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FLOW_ENVIRONMENT_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    assert.strictEqual(command.name.startsWith(commands.FLOW_ENVIRONMENT_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves Microsoft Flow environments (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            value: [
              {
                "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (default)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": true,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              },
              {
                "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (test)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": false,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5',
            displayName: 'Contoso (default)'
          },
          {
            name: 'Test-d87a7535-dd31-4437-bfe1-95340acd55c5',
            displayName: 'Contoso (test)'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves Microsoft Flow environments', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            value: [
              {
                "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (default)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": true,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              },
              {
                "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (test)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": false,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5',
            displayName: 'Contoso (default)'
          },
          {
            name: 'Test-d87a7535-dd31-4437-bfe1-95340acd55c5',
            displayName: 'Contoso (test)'
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
      if ((opts.url as string).indexOf(`/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            value: [
              {
                "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (default)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": true,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              },
              {
                "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (test)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": false,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
            "location": "europe",
            "type": "Microsoft.ProcessSimple/environments",
            "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
            "properties": {
              "displayName": "Contoso (default)",
              "createdTime": "2018-03-22T20:20:46.08653Z",
              "createdBy": {
                "id": "SYSTEM",
                "displayName": "SYSTEM",
                "type": "NotSpecified"
              },
              "provisioningState": "Succeeded",
              "creationType": "DefaultTenant",
              "environmentSku": "Default",
              "environmentType": "Production",
              "isDefault": true,
              "azureRegionHint": "westeurope",
              "runtimeEndpoints": {
                "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                "microsoft.PowerApps": "https://europe.api.powerapps.com",
                "microsoft.Flow": "https://europe.api.flow.microsoft.com"
              }
            }
          },
          {
            "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
            "location": "europe",
            "type": "Microsoft.ProcessSimple/environments",
            "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
            "properties": {
              "displayName": "Contoso (test)",
              "createdTime": "2018-03-22T20:20:46.08653Z",
              "createdBy": {
                "id": "SYSTEM",
                "displayName": "SYSTEM",
                "type": "NotSpecified"
              },
              "provisioningState": "Succeeded",
              "creationType": "DefaultTenant",
              "environmentSku": "Default",
              "environmentType": "Production",
              "isDefault": false,
              "azureRegionHint": "westeurope",
              "runtimeEndpoints": {
                "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                "microsoft.PowerApps": "https://europe.api.powerapps.com",
                "microsoft.Flow": "https://europe.api.flow.microsoft.com"
              }
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

  it('correctly handles no environments', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            value: []
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
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
              value: `Resource '' does not exist or one of its queried reference-property objects are not present`
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`)));
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
});