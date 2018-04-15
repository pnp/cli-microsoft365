import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./propertybag-set');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { IdentityResponse } from './propertybag-base';
import config from '../../../../config';

describe(commands.PROPERTYBAG_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
  const stubAllPostRequests = (
    requestObjectIdentityResp: any = null,
    folderObjectIdentityResp: any = null,
    setPropertyResp: any = null,
    effectiveBasePermissionsResp: any = null
  ): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

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

      // fake requestFolderObjectIdentity
      if (opts.body.indexOf('GetFolderByServerRelativeUrl') > -1) {
        if (folderObjectIdentityResp) {
          return folderObjectIdentityResp;
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

      // fake property set success for site and folder
      if (opts.body.indexOf('SetFieldValue') > -1) {
        if (setPropertyResp) {
          return setPropertyResp;
        } else {

          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7507.1203",
              "ErrorInfo": null,
              "TraceCorrelationId": "986d549e-d035-5000-2a28-c7306cd17024"
            }]));
        }
      }

      if (opts.body.indexOf('EffectiveBasePermissions') > -1) {
        if (effectiveBasePermissionsResp) {
          return effectiveBasePermissionsResp;
        } else {
          // effective base permissions (success case)
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7514.1204",
              "ErrorInfo": null,
              "TraceCorrelationId": "2d64579e-00e9-5000-71ce-fdad238b27fc"
            }, 7, {
              "_ObjectType_": "SP.Web",
              "_ObjectIdentity_": "2d64579e-00e9-5000-71ce-fdad238b27fc|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:692102df-335d-41e2-aa44-425b626037ea:web:f7fb12c3-ca68-4060-b1b0-c27a6bfffeb2",
              "EffectiveBasePermissions": {
                "_ObjectType_": "SP.BasePermissions",
                "High": 2147483647,
                "Low": 4294967295
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
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
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
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.PROPERTYBAG_SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {}, appCatalogUrl: 'https://contoso-admin.sharepoint.com' }, () => {
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
    cmdInstance.action({ options: { webUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert.equal(telemetry.name, commands.PROPERTYBAG_SET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { webUrl: 'https://contoso.sharepoint.com/sites/abc' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call setProperty when folder is not specified', (done) => {
    stubAllPostRequests();
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const setPropertySpy = sinon.spy((command as any), 'setProperty');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      debug: true,

    }
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002fabc"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(setPropertySpy.calledWith(objIdentity, options));
        assert(setPropertySpy.calledOnce === true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore((command as any)['setProperty']);
      }
    });
  });

  it('should call setProperty when folder is specified', (done) => {
    stubAllPostRequests(new Promise(resolve => {
      return resolve(JSON.stringify([{
        "SchemaVersion": "15.0.0.0",
        "LibraryVersion": "16.0.7331.1206",
        "ErrorInfo": null,
        "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
      }, 7, {
        "_ObjectType_": "SP.Web",
        "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
        "ServerRelativeUrl": "\u002f"
      }]));
    }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const setPropertySpy = sinon.spy((command as any), 'setProperty');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',

    }
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(setPropertySpy.calledWith(objIdentity, options));
        assert(setPropertySpy.calledOnce === true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore((command as any)['setProperty']);
      }
    });
  });

  it('should call setProperty when list folder is specified', (done) => {
    stubAllPostRequests();
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const setPropertySpy = sinon.spy((command as any), 'setProperty');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com/sites/abc',
      key: 'key1',
      value: 'value1',
      folder: '/Shared Documents',

    }
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/sites/abc/Shared Documents"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(setPropertySpy.calledWith(objIdentity, options));
        assert(setPropertySpy.calledOnce === true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore((command as any)['setProperty']);
      }
    });
  });

  it('should send correct property set request body when folder is not specified', (done) => {
    const requestStub: sinon.SinonStub = stubAllPostRequests();
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',

    }
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002fabc"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        const bodyPayload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">${(options as any).key}</Parameter><Parameter Type="String">${(options as any).value}</Parameter></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="AllProperties" /><Identity Id="198" Name="${objIdentity.objectIdentity}" /></ObjectPaths></Request>`
        assert(requestStub.calledWith(sinon.match({ body: bodyPayload })));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should send correct property set request body when folder is specified', (done) => {
    const requestStub: sinon.SinonStub = stubAllPostRequests();
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',

    }
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        const bodyPayload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">${(options as any).key}</Parameter><Parameter Type="String">${(options as any).value}</Parameter></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="Properties" /><Identity Id="198" Name="${objIdentity.objectIdentity}" /></ObjectPaths></Request>`
        assert(requestStub.calledWith(sinon.match({ body: bodyPayload })));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle requestObjectIdentity reject promise', (done) => {
    stubAllPostRequests(new Promise<any>((resolve, reject) => { return reject('requestObjectIdentity error'); }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',

    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('requestObjectIdentity error')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle requestObjectIdentity ClientSvc error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestObjectIdentity ClientSvc error" } }]);
    stubAllPostRequests(new Promise<any>((resolve, reject) => { return resolve(error); }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',

    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('requestObjectIdentity ClientSvc error')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle requestFolderObjectIdentity reject promise', (done) => {
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return reject('abc'); }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      debug: true
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('abc')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle requestFolderObjectIdentity ClientSvc error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestFolderObjectIdentity error" } }]);
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return resolve(error); }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      verbose: true,

    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('requestFolderObjectIdentity error')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle requestFolderObjectIdentity ClientSvc empty error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return resolve(error); }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      debug: true,

    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should requestFolderObjectIdentity reject promise if _ObjectIdentity_ not found', (done) => {
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return resolve('[{}]') }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_parentid',
      value: 'value1',

    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Cannot proceed. Folder _ObjectIdentity_ not found')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle isNoScriptSite = true', (done) => {
    stubAllPostRequests(null, null, null, new Promise<any>((resolve, reject) => {
      return resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0",
          "LibraryVersion": "16.0.7514.1204",
          "ErrorInfo": null,
          "TraceCorrelationId": "e811579e-009e-5000-ccf1-d233618f3d4f"
        }, 11, {
          "_ObjectType_": "SP.Web",
          "_ObjectIdentity_": "e811579e-009e-5000-ccf1-d233618f3d4f|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:692102df-335d-41e2-aa44-425b626037ea:web:f7fb12c3-ca68-4060-b1b0-c27a6bfffeb2", "EffectiveBasePermissions": {
            "_ObjectType_": "SP.BasePermissions",
            "High": 2147483647,
            "Low": 4294705151
          }
        }
      ]));
    }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      debug: true
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Site has NoScript enabled, and setting property bag values is not supported')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle getEffectiveBasePermissions reject promise', (done) => {
    stubAllPostRequests(null, null, null, new Promise<any>((resolve, reject) => { return reject('getEffectiveBasePermissions abc'); }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      debug: true
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('getEffectiveBasePermissions abc')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle getEffectiveBasePermissions ClientSvc error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "getEffectiveBasePermissions error" } }]);
    stubAllPostRequests(null, null, null, new Promise<any>((resolve, reject) => { return resolve(error); }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      verbose: true
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('getEffectiveBasePermissions error')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle getEffectiveBasePermissions ClientSvc empty error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, null, null, new Promise<any>((resolve, reject) => { return resolve(error); }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      debug: true
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should getEffectiveBasePermissions reject promise if EffectiveBasePermissions not found', (done) => {
    stubAllPostRequests(null, null, null, new Promise<any>((resolve, reject) => { return resolve('[{}]') }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_parentid',
      value: 'value1'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Cannot proceed. EffectiveBasePermissions not found')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle setProperty reject promise response', (done) => {
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return reject('setProperty promise error'); }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    const setPropertySpy = sinon.spy((command as any), 'setProperty');
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      verbose: true,

    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(setPropertySpy.calledOnce === true);
        assert(cmdInstanceLogSpy.calledWith(new CommandError('setProperty promise error')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle setProperty ClientSvc error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "setProperty error" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return resolve(error); }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      verbose: true,

    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('setProperty error')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('should correctly handle setProperty ClientSvc empty error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return resolve(error); }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    const options: Object = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      verbose: true,

    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
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
    sinon.stub(Command.prototype, 'options').callsFake(() => { return undefined; });
    const options = (command.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.equal(actual, "Missing required option url");
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          webUrl: 'foo',
          key: 'key1',
          value: 'value1'
        }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the property value option valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          webUrl: 'https://contoso.sharepoint.com',
          key: 'key1'
        }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the key option is not valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          webUrl: 'https://contoso.sharepoint.com'
        }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation when the url option specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          webUrl: 'https://contoso.sharepoint.com',
          key: 'key1',
          value: 'value1'
        }
    });
    assert.equal(actual, true);
  });

  it('passes validation when the url and folder options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          webUrl: 'https://contoso.sharepoint.com',
          key: 'key1',
          value: 'value1',
          folder: '/'
        }
    });
    assert.equal(actual, true);
  });

  it('doesn\'t fail validation if the optional folder option not specified', () => {
    const actual = (command.validate() as CommandValidate)(
      {
        options:
          {
            webUrl: 'https://contoso.sharepoint.com',
            key: 'key1',
            value: 'value1'
          }
      });
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.PROPERTYBAG_SET));
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
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        key: 'key1',
        value: 'value1',

      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(auth.getAccessToken);
      }
    });
  });
});