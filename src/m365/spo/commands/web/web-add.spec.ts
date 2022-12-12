import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./web-add');

describe(commands.WEB_ADD, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.WEB_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['url']);
  });

  it('creates web without inheriting the navigation', async () => {
    let configuredNavigation: boolean = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/webinfos/add') {
        return Promise.resolve({
          Configuration: 0,
          Created: "2018-01-24T18:24:20",
          Description: '',
          Id: "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language: opts.data.parameters.Language,
          LastItemModifiedDate: "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl: `/${opts.data.parameters.Url}`,
          Title: opts.data.parameters.Title,
          WebTemplate: "STS",
          WebTemplateId: 0
        });
      }

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        configuredNavigation = true;
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, {
      options: {
        title: "subsite",
        url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        locale: 1033,
        breakInheritance: true,
        inheritNavigation: false,
        debug: true
      }
    });
    assert(loggerLogSpy.calledWith({
      Configuration: 0,
      Created: "2018-01-24T18:24:20",
      Description: '',
      Id: "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
      Language: 1033,
      LastItemModifiedDate: "2018-01-24T18:24:27Z",
      LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
      ServerRelativeUrl: "/subsite",
      Title: "subsite",
      WebTemplate: "STS",
      WebTemplateId: 0
    }), 'Invalid web info');
    assert.strictEqual(configuredNavigation, false, 'Configured inheriting navigation while not expected');
  });

  it('creates web and does not set the inherit navigation (Noscript enabled)', async () => {
    let configuredNavigation: boolean = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/webinfos/add') {
        return Promise.resolve({
          Configuration: 0,
          Created: "2018-01-24T18:24:20",
          Description: "subsite",
          Id: "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language: 1033,
          LastItemModifiedDate: "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl: "/subsite",
          Title: "subsite",
          WebTemplate: "STS",
          WebTemplateId: 0
        });
      }

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        configuredNavigation = true;
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/effectivebasepermissions') > -1) {
        // PermissionKind.ManageLists, PermissionKind.AddListItems, PermissionKind.DeleteListItems
        return Promise.resolve(
          {
            High: 2058,
            Low: 0
          }
        );
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, {
      options: {
        title: "subsite",
        url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        locale: 1033
      }
    });
    assert(loggerLogSpy.calledWith({
      Configuration: 0,
      Created: "2018-01-24T18:24:20",
      Description: "subsite",
      Id: "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
      Language: 1033,
      LastItemModifiedDate: "2018-01-24T18:24:27Z",
      LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
      ServerRelativeUrl: "/subsite",
      Title: "subsite",
      WebTemplate: "STS",
      WebTemplateId: 0
    }), 'Incorrect web info');
    assert.strictEqual(configuredNavigation, false, 'Configured inheriting navigation while not expected');
  });

  it('creates web and does not set the inherit navigation (Noscript enabled; debug)', async () => {
    let configuredNavigation: boolean = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({
          Configuration: 0,
          Created: "2018-01-24T18:24:20",
          Description: "subsite",
          Id: "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language: 1033,
          LastItemModifiedDate: "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl: "/subsite",
          Title: "subsite",
          WebTemplate: "STS",
          WebTemplateId: 0
        });
      }

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        configuredNavigation = true;
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/effectivebasepermissions') > -1) {
        // PermissionKind.ManageLists, PermissionKind.AddListItems, PermissionKind.DeleteListItems
        return Promise.resolve(
          {
            High: 2058,
            Low: 0
          }
        );
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, {
      options: {
        title: "subsite",
        url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        locale: 1033,
        debug: true
      }
    });
    assert(loggerLogSpy.calledWith({
      Configuration: 0,
      Created: "2018-01-24T18:24:20",
      Description: "subsite",
      Id: "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
      Language: 1033,
      LastItemModifiedDate: "2018-01-24T18:24:27Z",
      LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
      ServerRelativeUrl: "/subsite",
      Title: "subsite",
      WebTemplate: "STS",
      WebTemplateId: 0
    }), 'Incorrect web info');
    assert.strictEqual(configuredNavigation, false, 'Configured inheriting navigation while not expected');
  });

  it('creates web and inherits the navigation (debug)', async () => {
    let configuredNavigation: boolean = false;

    // Create web
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({
          Configuration: 0,
          Created: "2018-01-24T18:24:20",
          Description: "subsite",
          Id: "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language: 1033,
          LastItemModifiedDate: "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl: "/subsite",
          Title: "subsite",
          WebTemplate: "STS",
          WebTemplateId: 0
        });
      }

      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.data.indexOf("UseShared") > -1) {
        configuredNavigation = true;

        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1203", "ErrorInfo": null, "TraceCorrelationId": "4556449e-0067-4000-1529-39a0d88e307d"
          }, 1, {
            "IsNull": false
          }, 3, {
            "IsNull": false
          }, 5, {
            "IsNull": false
          }, 7, {
            "_ObjectType_": "SP.Navigation", "UseShared": true
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    // Full permission.
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/effectivebasepermissions') > -1) {
        return Promise.resolve(
          {
            High: 2147483647,
            Low: 4294967295
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        title: "subsite",
        url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        locale: 1033,
        debug: true
      }
    });
    assert.strictEqual(configuredNavigation, true);
  });

  it('creates web and inherits the navigation', async () => {
    let configuredNavigation: boolean = false;

    // Create web
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({
          Configuration: 0,
          Created: "2018-01-24T18:24:20",
          Description: "subsite",
          Id: "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language: 1033,
          LastItemModifiedDate: "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl: "/subsite",
          Title: "subsite",
          WebTemplate: "STS",
          WebTemplateId: 0
        });
      }

      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.data.indexOf("UseShared") > -1) {
        configuredNavigation = true;

        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1203", "ErrorInfo": null, "TraceCorrelationId": "4556449e-0067-4000-1529-39a0d88e307d"
          }, 1, {
            "IsNull": false
          }, 3, {
            "IsNull": false
          }, 5, {
            "IsNull": false
          }, 7, {
            "_ObjectType_": "SP.Navigation", "UseShared": true
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    // Full permission.
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/effectivebasepermissions') > -1) {
        return Promise.resolve(
          {
            High: 2147483647,
            Low: 4294967295
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        title: "subsite",
        url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        locale: 1033
      }
    });
    assert.strictEqual(configuredNavigation, true);
  });

  it('correctly handles the set inheritNavigation error', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      // Create web
      if ((opts.url as string).indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({
          Configuration: 0,
          Created: "2018-01-24T18:24:20",
          Description: "subsite",
          Id: "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language: 1033,
          LastItemModifiedDate: "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl: "/subsite",
          Title: "subsite",
          WebTemplate: "STS",
          WebTemplateId: 0
        });
      }

      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        // SetInheritNavigation failed.
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": {
              "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799", "ErrorCode": -2146232832, "ErrorTypeName": "Microsoft.SharePoint.SPFieldValidationException"
            }, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    // Full permission.
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/effectivebasepermissions') > -1) {
        return Promise.resolve(
          {
            High: 2147483647,
            Low: 4294967295
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: "subsite",
        url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        local: 1033,
        debug: true
      }
    } as any), new CommandError('An error has occurred.'));
  });

  it('correctly handles the createweb call error', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/webinfos/add') > -1) {
        return Promise.reject({
          error: {
            "odata.error": {
              "code": "-2147024713, Microsoft.SharePoint.SPException",
              "message": {
                "lang": "en-US",
                "value": "The Web site address \"/sites/test/subsite\" is already in use."
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: "subsite",
        url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com/sites/test",
        inheritNavigation: true,
        local: 1033,
        debug: true
      }
    } as any), new CommandError("The Web site address \"/sites/test/subsite\" is already in use."));
  });

  it('creates web and handles the effectivebasepermission call error', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({
          Configuration: 0,
          Created: "2018-01-24T18:24:20",
          Description: "subsite",
          Id: "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language: 1033,
          LastItemModifiedDate: "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl: "/subsite",
          Title: "subsite",
          WebTemplate: "STS",
          WebTemplateId: 0
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/effectivebasepermissions') > -1) {
        return Promise.reject({
          error: {
            "odata.error": {
              "code": "-2147024713, Microsoft.SharePoint.SPException",
              "message": {
                "lang": "en-US",
                "value": "An error has occurred."
              }
            }
          }
        });
      }

      return Promise.resolve('abc');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: "subsite",
        url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        local: 1033,
        debug: true
      }
    } as any), new CommandError('An error has occurred.'));
  });

  it('correctly handles the parentweb contextinfo call error', async () => {
    sinonUtil.restore(spo.getRequestDigest);
    sinon.stub(spo, 'getRequestDigest').callsFake(() => { return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } }); });

    await assert.rejects(command.action(logger, {
      options: {
        title: "subsite",
        url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        local: 1033,
        debug: true
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('correctly handles generic API error', async () => {
    sinonUtil.restore(spo.getRequestDigest);
    sinon.stub(spo, 'getRequestDigest').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: "subsite",
        url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        local: 1033,
        debug: true
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('passes validation if all required options are specified', async () => {
    const actual = await command.validate({
      options: {
        title: "subsite", url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com", webTemplate: "STS#0"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if all required options and valid locale are specified', async () => {
    const actual = await command.validate({
      options: {
        title: "subsite", url: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com", webTemplate: "STS#0", locale: 1033
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the parentWebUrl option not specified', async () => {
    const actual = await command.validate({
      options: {
        title: "subsite",
        url: "subsite", webTemplate: "STS#0", locale: 1033
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the parentWebUrl option is not a valid SharePoint URL', async () => {
    const actual = await command.validate({
      options: {
        title: "subsite",
        url: "subsite", webTemplate: "STS#0", locale: 1033,
        parentWebUrl: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified locale is not a number', async () => {
    const actual = await command.validate({
      options: {
        title: "subsite", url: "subsite", parentWebUrl: "https://contoso.sharepoint.com", webTemplate: 'STS#0', locale: 'abc'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
