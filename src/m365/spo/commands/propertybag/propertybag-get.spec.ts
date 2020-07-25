import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./propertybag-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import config from '../../../../config';
import { IdentityResponse, ClientSvc } from '../../ClientSvc';

describe(commands.PROPERTYBAG_GET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let stubAllPostRequests: any = (
    requestObjectIdentityResp: any = null,
    getFolderPropertyBagResp: any = null,
    getWebPropertyBagResp: any = null
  ) => {
    return sinon.stub(request, 'post').callsFake((opts) => {
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
            "ServerRelativeUrl": "\u002f"
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
      (command as any).filterByKey,
      (command as any).getFolderPropertyBag,
      ClientSvc.prototype.getCurrentWebIdentity
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
    assert.strictEqual(command.name.startsWith(commands.PROPERTYBAG_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should call getWebPropertyBag when folder is not specified and site is /', (done) => {
    stubAllPostRequests();
    const getWebPropertyBagSpy = sinon.spy((command as any), 'getWebPropertyBag');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      debug: true
    }
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002f"
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

  it('should call getWebPropertyBag when folder is not specified and site is /sites/test', (done) => {
    stubAllPostRequests(new Promise((resolve, reject) => { return resolve(JSON.stringify([{
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7331.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
    }, 7, {
      "_ObjectType_": "SP.Web",
      "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      "ServerRelativeUrl": "\u002fsites\u002ftest"
    }])) }));

    const getWebPropertyBagSpy = sinon.spy((command as any), 'getWebPropertyBag');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      debug: true
    }
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002ftest"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(getWebPropertyBagSpy.calledWith(objIdentity, 'https://contoso.sharepoint.com/sites/test', cmdInstance));
        assert(getWebPropertyBagSpy.calledOnce === true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call getFolderPropertyBag when folder is specified and site is /', (done) => {
    stubAllPostRequests();
    const getFolderPropertyBagSpy = sinon.spy((command as any), 'getFolderPropertyBag');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      debug: true
    }
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002f"
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

  it('should call getFolderPropertyBag when folder is specified and site is /sites/test', (done) => {
    stubAllPostRequests(new Promise((resolve, reject) => { return resolve(JSON.stringify([{
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7331.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
    }, 7, {
      "_ObjectType_": "SP.Web",
      "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      "ServerRelativeUrl": "\u002fsites\u002ftest"
    }])) }));

    const getFolderPropertyBagSpy = sinon.spy((command as any), 'getFolderPropertyBag');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      folder: '/',
      debug: true
    }
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002ftest"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(getFolderPropertyBagSpy.calledWith(objIdentity, 'https://contoso.sharepoint.com/sites/test', '/', cmdInstance));
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

  it('should correctly handle getFolderPropertyBag ErrorMessage null response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": undefined } }]);

    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return resolve(error) }));
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'vti_parentid',
      folder: '/'
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle getFolderPropertyBag ErrorMessage null response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": undefined } }]);

    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return resolve(error) }));
    cmdInstance.action = command.action();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'vti_parentid'
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly return string property', (done) => {
    stubAllPostRequests();
    const filterByKeySpy = sinon.spy((command as any), 'filterByKey');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_parentid'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(filterByKeySpy.calledOnce === true);

        const out = cmdInstanceLogSpy.lastCall.args[0];
        assert.strictEqual(out, '{1C5271C8-DB93-459E-9C18-68FC33EFD856}');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly return date property (text)', (done) => {
    stubAllPostRequests();
    const filterByKeySpy = sinon.spy((command as any), 'filterByKey');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_timelastmodified' //\/Date(2017,10,7,11,29,31,0)\/
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(filterByKeySpy.calledOnce === true);

        const out = cmdInstanceLogSpy.lastCall.args[0];
        const expectedDate = new Date(2017, 10, 7, 11, 29, 31, 0);
        assert.strictEqual(out.getUTCMonth(), expectedDate.getUTCMonth(), 'getUTCMonth');
        assert.strictEqual(out.getUTCFullYear(), expectedDate.getUTCFullYear(), 'getUTCFullYear');
        assert.strictEqual(out.getUTCDate(), expectedDate.getUTCDate(), 'getUTCDate');
        assert.strictEqual(out.getUTCHours(), expectedDate.getUTCHours(), 'getUTCHours');
        assert.strictEqual(out.getUTCMinutes(), expectedDate.getUTCMinutes(), 'getUTCMinutes');
        assert.strictEqual(out.getSeconds(), expectedDate.getSeconds(), 'getSeconds');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly return date property (json)', (done) => {
    stubAllPostRequests();
    const filterByKeySpy = sinon.spy((command as any), 'filterByKey');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_timelastmodified',
      output: 'json'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(filterByKeySpy.calledOnce === true);

        const out = cmdInstanceLogSpy.lastCall.args[0];
        assert.strictEqual(Object.prototype.toString.call(out), '[object Date]');
        const expectedDate = new Date(2017, 10, 7, 11, 29, 31, 0);
        assert.strictEqual(out.getUTCMonth(), expectedDate.getUTCMonth(), 'getUTCMonth');
        assert.strictEqual(out.getUTCFullYear(), expectedDate.getUTCFullYear(), 'getUTCFullYear');
        assert.strictEqual(out.getUTCDate(), expectedDate.getUTCDate(), 'getUTCDate');
        assert.strictEqual(out.getUTCHours(), expectedDate.getUTCHours(), 'getUTCHours');
        assert.strictEqual(out.getUTCMinutes(), expectedDate.getUTCMinutes(), 'getUTCMinutes');
        assert.strictEqual(out.getSeconds(), expectedDate.getSeconds(), 'getSeconds');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly return int property', (done) => {
    stubAllPostRequests();
    const filterByKeySpy = sinon.spy((command as any), 'filterByKey');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_level'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(filterByKeySpy.calledOnce === true);

        const out = cmdInstanceLogSpy.lastCall.args[0];
        assert.strictEqual(out, 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly return int property with value 0', (done) => {
    stubAllPostRequests();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_folderitemcount'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0], 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly return bool property', (done) => {
    stubAllPostRequests();
    const filterByKeySpy = sinon.spy((command as any), 'filterByKey');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_candeleteversion'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(filterByKeySpy.calledOnce === true);
        
        const out = cmdInstanceLogSpy.lastCall.args[0];
        assert.strictEqual(out, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly return property not found (verbose)', (done) => {
    stubAllPostRequests();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'abc',
      verbose: true
    }

    cmdInstance.action({ options: options }, () => {

      try {
        const out = cmdInstanceLogSpy.lastCall.args[0];
        assert.strictEqual(out, 'Property not found.');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly return empty line if not found and not verbose', (done) => {
    stubAllPostRequests();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'abc',
      verbose: false
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should reject promise if _ObjectIdentity_ not found', (done) => {
    stubAllPostRequests(new Promise<any>((resolve, reject) => { return resolve('[{}]') }));
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_parentid'
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Cannot proceed. _ObjectIdentity_ not found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should reject promise if Properties not found', (done) => {
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return resolve('[{}]') }));
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_parentid'
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Cannot proceed. Properties not found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should reject promise if AllProperties not found', (done) => {
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return resolve('[{}]') }));
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'vti_parentid'
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Cannot proceed. AllProperties not found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should return error if requestObjectIdentity reqest failed', (done) => {
    stubAllPostRequests(new Promise<any>((resolve, reject) => { return reject('error1') }));
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'vti_parentid'
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('error1')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly post url, headers and body when calling client.svc when requestObjectIdentity', (done) => {
    const postRequestSpy: sinon.SinonSpy = stubAllPostRequests();
    
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'vti_parentid'
    };

    cmdInstance.action({ options: options }, () => {
      try {
        const secondCall = postRequestSpy.getCalls()[0];
        assert(secondCall.calledWith(sinon.match({ url: 'https://contoso.sharepoint.com/_vti_bin/client.svc/ProcessQuery' })), 'url');
        assert(secondCall.calledWith(sinon.match({ headers: { 'X-RequestDigest': 'abc'}})), 'request digest');
        assert(secondCall.calledWith(sinon.match({ body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
        })), 'body');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly post url, headers and body when calling client.svc when getWebPropertyBag', (done) => {
    const postRequestSpy: sinon.SinonSpy = stubAllPostRequests();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'vti_parentid'
    };

    cmdInstance.action({ options: options }, () => {
      try {
        const lastCall = postRequestSpy.lastCall;
        assert(lastCall.calledWith(sinon.match({ url: 'https://contoso.sharepoint.com/_vti_bin/client.svc/ProcessQuery' })));
        assert(lastCall.calledWith(sinon.match({ headers: { 'X-RequestDigest': 'abc'}})));
        assert(lastCall.calledWith(sinon.match({ body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="97" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /><Property Name="AllProperties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275" /></ObjectPaths></Request>`
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly post payload when calling client.svc when getFolderPropertyBag and site is /', (done) => {
    const postRequestSpy: sinon.SinonSpy = stubAllPostRequests();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'vti_parentid',
      folder: '/'
    };

    cmdInstance.action({ options: options }, () => {

      try {
        const lastCall = postRequestSpy.lastCall;
        assert(lastCall.calledWith(sinon.match({ url: 'https://contoso.sharepoint.com/_vti_bin/client.svc/ProcessQuery' })));
        assert(lastCall.calledWith(sinon.match({ headers: { 'X-RequestDigest': 'abc'}})));
        assert(lastCall.calledWith(sinon.match({ body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="Properties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="9" ParentId="5" Name="GetFolderByServerRelativeUrl"><Parameters><Parameter Type="String">/</Parameter></Parameters></Method><Identity Id="5" Name="38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275" /></ObjectPaths></Request>`
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly post payload when calling client.svc when getFolderPropertyBag and site is /sites/test', (done) => {
    const postRequestSpy: sinon.SinonSpy = stubAllPostRequests(new Promise((resolve, reject) => { return resolve(JSON.stringify([{
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7331.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
    }, 7, {
      "_ObjectType_": "SP.Web",
      "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      "ServerRelativeUrl": "\u002fsites\u002ftest"
    }])) }));

    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      key: 'vti_parentid',
      folder: '/'
    };

    cmdInstance.action({ options: options }, () => {

      try {
        const lastCall = postRequestSpy.lastCall;
        assert(lastCall.calledWith(sinon.match({ url: 'https://contoso.sharepoint.com/sites/test/_vti_bin/client.svc/ProcessQuery' })));
        assert(lastCall.calledWith(sinon.match({ headers: { 'X-RequestDigest': 'abc'}})));
        assert(lastCall.calledWith(sinon.match({ body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="Properties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="9" ParentId="5" Name="GetFolderByServerRelativeUrl"><Parameters><Parameter Type="String">/sites/test/</Parameter></Parameters></Method><Identity Id="5" Name="38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275" /></ObjectPaths></Request>`
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly post payload when calling client.svc when getFolderPropertyBag and site is /', (done) => {
    const postRequestSpy: sinon.SinonSpy = stubAllPostRequests();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'vti_parentid',
      folder: '/'
    };

    cmdInstance.action({ options: options }, () => {

      try {
        const lastCall = postRequestSpy.lastCall;
        assert(lastCall.calledWith(sinon.match({ url: 'https://contoso.sharepoint.com/_vti_bin/client.svc/ProcessQuery' })));
        assert(lastCall.calledWith(sinon.match({ headers: { 'X-RequestDigest': 'abc'}})));
        assert(lastCall.calledWith(sinon.match({ body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="Properties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="9" ParentId="5" Name="GetFolderByServerRelativeUrl"><Parameters><Parameter Type="String">/</Parameter></Parameters></Method><Identity Id="5" Name="38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275" /></ObjectPaths></Request>`
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly post payload when calling client.svc when getFolderPropertyBag and site is /sites/test', (done) => {
    const postRequestSpy: sinon.SinonSpy = stubAllPostRequests(new Promise((resolve, reject) => { return resolve(JSON.stringify([{
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7331.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
    }, 7, {
      "_ObjectType_": "SP.Web",
      "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      "ServerRelativeUrl": "\u002fsites\u002ftest"
    }])) }));

    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      key: 'vti_parentid',
      folder: '/'
    };

    cmdInstance.action({ options: options }, () => {
      try {
        const lastCall = postRequestSpy.lastCall;
        assert(lastCall.calledWith(sinon.match({ url: 'https://contoso.sharepoint.com/sites/test/_vti_bin/client.svc/ProcessQuery' })));
        assert(lastCall.calledWith(sinon.match({ headers: { 'X-RequestDigest': 'abc'}})));
        assert(lastCall.calledWith(sinon.match({ body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="Properties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="9" ParentId="5" Name="GetFolderByServerRelativeUrl"><Parameters><Parameter Type="String">/sites/test/</Parameter></Parameters></Method><Identity Id="5" Name="38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275" /></ObjectPaths></Request>`
        })));
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
    const prop = (command as any).formatProperty('vti_timecreated', 'true');
    assert.strictEqual(prop.key, 'vti_timecreated');
    assert.strictEqual(prop.value, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          webUrl: 'foo',
          key: 'abc'
        }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the url and key options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          webUrl: "https://contoso.sharepoint.com",
          key: 'abc'
        }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the url, key and folder options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          webUrl: "https://contoso.sharepoint.com",
          key: 123,
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
            webUrl: "https://contoso.sharepoint.com",
            key: 'abc'
          }
      });
    assert.strictEqual(actual, true);
  });
});