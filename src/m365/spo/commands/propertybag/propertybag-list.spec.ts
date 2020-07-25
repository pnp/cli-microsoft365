import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./propertybag-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { ClientSvc, IdentityResponse } from '../../ClientSvc';

describe(commands.PROPERTYBAG_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let stubAllPostRequests: any = (
    requestObjectIdentityResp: any = null,
    getFolderPropertyBagResp: any = null,
    getWebPropertyBagResp: any = null
  ) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      // fake requestObjectIdentity
      if (opts.body.indexOf('3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a') > -1) {
        if (requestObjectIdentityResp) {
          return requestObjectIdentityResp;
        } else {
          return Promise.resolve(JSON.stringify([{
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7331.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
          }, 7, {
            "_ObjectType_": "SP.Web",
            "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
            "ServerRelativeUrl": "\u002fsites\u002fabc"
          }]));
        }
      }

      // fake getFolderPropertyBag
      if (opts.body.indexOf('GetFolderByServerRelativeUrl') > -1) {
        if (getFolderPropertyBagResp) {
          return getFolderPropertyBagResp;
        } else {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1206", "ErrorInfo": null, "TraceCorrelationId": "93e5499e-00f1-5000-1f36-3ab12512a7e9"
            }, 18, {
              "IsNull": false
            }, 19, {
              "_ObjectIdentity_": "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1"
            }, 20, {
              "_ObjectType_": "SP.Folder", "_ObjectIdentity_": "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1", "Properties": {
                "_ObjectType_": "SP.PropertyValues", "vti_folderitemcount$  Int32": 0, "vti_level$  Int32": 1, "vti_parentid": "{1C5271C8-DB93-459E-9C18-68FC33EFD856}", "vti_winfileattribs": "00000012", "vti_candeleteversion": "true", "vti_foldersubfolderitemcount$  Int32": 0, "vti_timelastmodified": "\/Date(2017,10,7,11,29,31,0)\/", "vti_dirlateststamp": "\/Date(2018,1,12,22,34,31,0)\/", "vti_isscriptable": "false", "vti_isexecutable": "false", "vti_metainfoversion$  Int32": 1, "vti_isbrowsable": "true", "vti_timecreated": "\/Date(2017,10,7,11,29,31,0)\/", "vti_etag": "\"{DF4291DE-226F-4C39-BBCC-DF21915F5FC1},256\"", "vti_hassubdirs": "true", "vti_docstoreversion$  Int32": 256, "vti_rtag": "rt:DF4291DE-226F-4C39-BBCC-DF21915F5FC1@00000000256", "vti_docstoretype$  Int32": 1, "vti_replid": "rid:{DF4291DE-226F-4C39-BBCC-DF21915F5FC1}"
              }
            }
          ]));
        }
      }

      // fake getWebPropertyBag
      if (opts.body.indexOf('Property Name="AllProperties" SelectAll="true"') > -1) {
        if (getWebPropertyBagResp) {
          return getWebPropertyBagResp;
        } else {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1206", "ErrorInfo": null, "TraceCorrelationId": "e7e5499e-7031-5000-ccf1-ddcbe51e534c"
            }, 25, {
              "_ObjectType_": "SP.Web", "_ObjectIdentity_": "e7e5499e-7031-5000-ccf1-ddcbe51e534c|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5", "ServerRelativeUrl": "\u002fsites\u002fVisionTestDev1\u002fen", "AllProperties": {
                "_ObjectType_": "SP.PropertyValues", "_PnP_ProvisioningTemplateInfo": "{\"TemplateId\":\"TEMPLATE-B5D1728BA91E48E5B3FCB8CFF5CFCF66\",\"TemplateVersion\":1.0,\"TemplateSitePolicy\":null,\"ProvisioningTime\":\"2017-11-07T11:37:35.6130975+00:00\",\"Result\":true}", "vti_indexedpropertykeys": "XwBQAG4AUABfAFAAcgBvAHYAaQBzAGkAbwBuAGkAbgBnAFQAZQBtAHAAbABhAHQAZQBJAGQA|", "__InheritCurrentNavigation": "False", "_webnavigationsettings": "<?xml version=\"1.0\" encoding=\"utf-16\" standalone=\"yes\"?>\r\n<WebNavigationSettings Version=\"1.1\">\r\n  <SiteMapProviderSettings>\r\n    <SwitchableSiteMapProviderSettings Name=\"CurrentNavigationSwitchableProvider\" TargetProviderName=\"CurrentNavigation\" \u002f>\r\n    <TaxonomySiteMapProviderSettings Name=\"CurrentNavigationTaxonomyProvider\" Disabled=\"True\" \u002f>\r\n    <SwitchableSiteMapProviderSettings Name=\"GlobalNavigationSwitchableProvider\" TargetProviderName=\"GlobalNavigation\" \u002f>\r\n    <TaxonomySiteMapProviderSettings Name=\"GlobalNavigationTaxonomyProvider\" Disabled=\"True\" \u002f>\r\n  <\u002fSiteMapProviderSettings>\r\n  <NewPageSettings AddNewPagesToNavigation=\"True\" CreateFriendlyUrlsForNewPages=\"True\" \u002f>\r\n<\u002fWebNavigationSettings>\r\n", "vti_defaultlanguage": "en-us", "vti_mastercssfilecache": "corev15app.css", "_PnP_ProvisioningTemplateId": "TEMPLATE-B5D1728BA91E48E5B3FCB8CFF5CFCF66", "vti_extenderversion": "16.0.0.7025", "vti_approvallevels": "Approved Rejected Pending\\ Review", "vti_categories": "Travel Expense\\ Report Business Competition Goals\u002fObjectives Ideas Miscellaneous Waiting VIP In\\ Process Planning Schedule", "NoCrawl": "false", "$": "sdf", "__NavigationShowSiblings": "false"
              }
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc'
    }));
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
      request.post,
      (command as any).getWebPropertyBag,
      (command as any).getFolderPropertyBag,
      ClientSvc.prototype.getCurrentWebIdentity,
      (command as any).formatOutput
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
    assert.strictEqual(command.name.startsWith(commands.PROPERTYBAG_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should call getWebPropertyBag when folder is not specified', (done) => {
    stubAllPostRequests();
    const getWebPropertyBagSpy = sinon.spy((command as any), 'getWebPropertyBag');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      debug: true
    }
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002fabc"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(getWebPropertyBagSpy.calledWith(objIdentity, 'https://contoso.sharepoint.com', cmdInstance));
        assert(getWebPropertyBagSpy.calledOnce === true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call getFolderPropertyBag when folder is specified', (done) => {
    stubAllPostRequests();
    const getFolderPropertyBagSpy = sinon.spy((command as any), 'getFolderPropertyBag');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      debug: true
    }
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002fabc"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(getFolderPropertyBagSpy.calledWith(objIdentity, 'https://contoso.sharepoint.com', '/', cmdInstance));
        assert(getFolderPropertyBagSpy.calledOnce === true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle getFolderPropertyBag reject promise', (done) => {
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return reject('abc'); }));
    const getFolderPropertyBagSpy = sinon.spy((command as any), 'getFolderPropertyBag');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/'
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert(getFolderPropertyBagSpy.calledOnce === true);
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('abc')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle getWebPropertyBag reject promise', (done) => {
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return reject('abc1'); }));
    const getWebPropertyBagSpy = sinon.spy((command as any), 'getWebPropertyBag');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      debug: false
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert(getWebPropertyBagSpy.calledOnce === true);
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('abc1')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle getFolderPropertyBag ClientSvc error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "getFolderPropertyBag error" } }]);
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return resolve(error); }));
    const getFolderPropertyBagSpy = sinon.spy((command as any), 'getFolderPropertyBag');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      verbose: true
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert(getFolderPropertyBagSpy.calledOnce === true);
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('getFolderPropertyBag error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle getWebPropertyBag ClientSvc error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "getWebPropertyBag error" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return resolve(error); }));
    const getWebPropertyBagSpy = sinon.spy((command as any), 'getWebPropertyBag');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com'
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert(getWebPropertyBagSpy.calledOnce === true);
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('getWebPropertyBag error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle requestObjectIdentity error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestObjectIdentity error" } }]);

    stubAllPostRequests(new Promise<any>((resolve, reject) => { return resolve(error) }), null, null);
    const requestObjectIdentitySpy = sinon.spy(ClientSvc.prototype, 'getCurrentWebIdentity');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com'
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert(requestObjectIdentitySpy.calledOnce === true);
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('requestObjectIdentity error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle requestObjectIdentity ErrorMessage null response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": undefined } }]);

    stubAllPostRequests(new Promise<any>((resolve, reject) => { return resolve(error) }), null, null);
    const requestObjectIdentitySpy = sinon.spy(ClientSvc.prototype, 'getCurrentWebIdentity');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com'
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert(requestObjectIdentitySpy.calledOnce === true);
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly format response output (text)', (done) => {
    stubAllPostRequests();
    const formatOutputSpy = sinon.spy((command as any), 'formatOutput');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(formatOutputSpy.calledOnce === true);

        const out = cmdInstanceLogSpy.lastCall.args[0];
        const expectedDate = new Date(2017, 10, 7, 11, 29, 31, 0);

        assert.strictEqual(out[0].key, 'vti_folderitemcount');
        assert.strictEqual(out[0].value, 0);
        assert.strictEqual(out[1].key, 'vti_level');
        assert.strictEqual(out[1].value, 1);
        assert.strictEqual(out[2].key, 'vti_parentid');
        assert.strictEqual(out[2].value, '{1C5271C8-DB93-459E-9C18-68FC33EFD856}');
        assert.strictEqual(out[3].key, 'vti_winfileattribs');
        assert.strictEqual(out[3].value, '00000012');
        assert.strictEqual(out[4].key, 'vti_candeleteversion');
        assert.strictEqual(out[4].value, true);
        assert.strictEqual(out[5].key, 'vti_foldersubfolderitemcount');
        assert.strictEqual(out[5].value, 0);
        assert.strictEqual(out[6].key, 'vti_timelastmodified');
        assert.strictEqual(Object.prototype.toString.call(out[6].value), '[object Date]');
        assert.strictEqual((out[6].value as Date).getUTCMonth(), expectedDate.getUTCMonth(), 'getUTCMonth');
        assert.strictEqual((out[6].value as Date).getUTCFullYear(), expectedDate.getUTCFullYear(), 'getUTCFullYear');
        assert.strictEqual((out[6].value as Date).getUTCDate(), expectedDate.getUTCDate(), 'getUTCDate');
        assert.strictEqual((out[6].value as Date).getUTCHours(), expectedDate.getUTCHours(), 'getUTCHours');
        assert.strictEqual((out[6].value as Date).getUTCMinutes(), expectedDate.getUTCMinutes(), 'getUTCMinutes');
        assert.strictEqual((out[6].value as Date).getSeconds(), expectedDate.getSeconds(), 'getSeconds');
        assert.strictEqual(out[8].key, 'vti_isscriptable');
        assert.strictEqual(out[8].value, false);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsVerboseOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsVerboseOption = true;
      }
    });
    assert(containsVerboseOption);
  });

  it('supports specifying folder', () => {
    const options = (command.options() as CommandOption[]);
    let containsScopeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[folder]') > -1) {
        containsScopeOption = true;
      }
    });
    assert(containsScopeOption);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return []; });
    const options = (command.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('should properly format integer property', () => {
    const prop = (command as any).formatProperty('vti_folderitemcount$  Int32', 0);
    assert.strictEqual(prop.key, 'vti_folderitemcount');
    assert.strictEqual(prop.value, 0);
  });

  it('should properly format date property', () => {
    const prop = (command as any).formatProperty('vti_timecreated', '\/Date(2017,10,7,11,29,31,0)\/');
    assert.strictEqual(prop.key, 'vti_timecreated');
    assert.strictEqual(Object.prototype.toString.call(prop.value), '[object Date]');
    assert.strictEqual((prop.value as Date).toISOString(), new Date(2017, 10, 7, 11, 29, 31, 0).toISOString());
  });

  it('should properly format boolean property', () => {
    const prop = (command as any).formatProperty('vti_timecreated', 'false');
    assert.strictEqual(prop.key, 'vti_timecreated');
    assert.strictEqual(prop.value, false);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          webUrl: 'foo'
        }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the url option specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          webUrl: "https://contoso.sharepoint.com"
        }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the url and folder options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          webUrl: "https://contoso.sharepoint.com",
          folder: "/"
        }
    });
    assert.strictEqual(actual, true);
  });

  it('doesn\'t fail validation if the optional folder option not specified', () => {
    const actual = (command.validate() as CommandValidate)(
      {
        options:
          {
            webUrl: "https://contoso.sharepoint.com"
          }
      });
    assert.strictEqual(actual, true);
  });
});