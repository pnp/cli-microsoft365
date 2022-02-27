import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./environment-list');

describe(commands.ENVIRONMENT_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ENVIRONMENT_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'displayName']);
  });

  it('retrieves Microsoft Flow environments (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
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
            },
            "displayName": "Contoso (default)"
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
            },
            "displayName": "Contoso (test)"
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
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
            },
            "displayName": "Contoso (default)"
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
            },
            "displayName": "Contoso (test)"
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({
            value: []
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
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

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
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
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});