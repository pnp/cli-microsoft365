import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./web-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.WEB_ADD, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc' }); });
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
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.WEB_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates web without inheriting the navigation', (done) => {
    let configuredNavigation: boolean = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({
          Configuration: 0,
          Created: "2018-01-24T18:24:20",
          Description: '',
          Id: "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language: opts.body.parameters.Language,
          LastItemModifiedDate: "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl: `/${opts.body.parameters.Url}`,
          Title: opts.body.parameters.Title,
          WebTemplate: "STS",
          WebTemplateId: 0
        });
      }

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        configuredNavigation = true;
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({
      options: {
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        locale: 1033,
        breakInheritance: true,
        inheritNavigation: false,
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates web and does not set the inherit navigation (Noscript enabled)', (done) => {
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
    cmdInstance.action({
      options: {
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        locale: 1033
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates web and does not set the inherit navigation (Noscript enabled; debug)', (done) => {
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
    cmdInstance.action({
      options: {
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        locale: 1033,
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates web and inherits the navigation (debug)', (done) => {
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
        opts.body.indexOf("UseShared") > -1) {
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

    cmdInstance.action({
      options: {
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        locale: 1033,
        debug: true
      }
    }, () => {
      try {
        assert.strictEqual(configuredNavigation, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates web and inherits the navigation', (done) => {
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
        opts.body.indexOf("UseShared") > -1) {
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

    cmdInstance.action({
      options: {
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        locale: 1033
      }
    }, () => {
      try {
        assert.strictEqual(configuredNavigation, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles the set inheritNavigation error', (done) => {
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
        ]))
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

    cmdInstance.action({
      options: {
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        local: 1033,
        debug: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles the createweb call error', (done) => {
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

    cmdInstance.action({
      options: {
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com/sites/test",
        inheritNavigation: true,
        local: 1033,
        debug: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("The Web site address \"/sites/test/subsite\" is already in use.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates web and handles the effectivebasepermission call error', (done) => {
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
    cmdInstance.action({
      options: {
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        local: 1033,
        debug: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles the parentweb contextinfo call error', (done) => {
    Utils.restore((command as any).getRequestDigest);
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } }); });

    cmdInstance.action({
      options: {
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        local: 1033,
        debug: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles generic API error', (done) => {
    Utils.restore((command as any).getRequestDigest);
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action({
      options: {
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com",
        inheritNavigation: true,
        local: 1033,
        debug: true
      }
    }, (err?: any) => {
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
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('passes validation if all required options are specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        title: "subsite", webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com", webTemplate: "STS#0"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if all required options and valid locale are specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        title: "subsite", webUrl: "subsite",
        parentWebUrl: "https://contoso.sharepoint.com", webTemplate: "STS#0", locale: 1033
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the parentWebUrl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        title: "subsite",
        webUrl: "subsite", webTemplate: "STS#0", locale: 1033
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the parentWebUrl option is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        title: "subsite",
        webUrl: "subsite", webTemplate: "STS#0", locale: 1033,
        parentWebUrl: 'foo'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified locale is not a number', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        title: "subsite", webUrl: "subsite", parentWebUrl: "https://contoso.sharepoint.com", webTemplate: 'STS#0', locale: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });
});