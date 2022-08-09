import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { IdentityResponse, sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./propertybag-remove');

describe(commands.PROPERTYBAG_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptOptions: any;
  const stubAllPostRequests = (
    requestObjectIdentityResp: any = null,
    folderObjectIdentityResp: any = null,
    removePropertyResp: any = null
  ): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      // fake requestObjectIdentity
      if (opts.data.indexOf('3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a') > -1) {
        if (requestObjectIdentityResp) {
          return requestObjectIdentityResp;
        }
        else {
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
      if (opts.data.indexOf('GetFolderByServerRelativeUrl') > -1) {
        if (folderObjectIdentityResp) {
          return folderObjectIdentityResp;
        }
        else {
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

      // fake deleteion success for site and folder
      if (opts.data.indexOf('SetFieldValue') > -1) {
        if (removePropertyResp) {
          return removePropertyResp;
        }
        else {

          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7507.1203",
              "ErrorInfo": null,
              "TraceCorrelationId": "986d549e-d035-5000-2a28-c7306cd17024"
            }]));
        }
      }

      return Promise.reject('Invalid request');
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
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
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      (command as any).removeProperty,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      spo.getRequestDigest
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PROPERTYBAG_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should remove property without prompting with confirmation argument', (done) => {
    stubAllPostRequests();

    command.action(logger, {
      options: {
        verbose: false,
        webUrl: 'https://contoso.sharepoint.com',
        key: 'key1',
        confirm: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should prompt before removing property when confirmation argument not passed', (done) => {
    command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        key: 'key1'
      }
    }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should abort property remove when prompt not confirmed', (done) => {
    const postCallsSpy: sinon.SinonStub = stubAllPostRequests();

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        key: 'key1'
      }
    }, () => {
      try {
        assert(postCallsSpy.notCalled === true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should remove property when prompt confirmed', (done) => {
    const postCallsSpy: sinon.SinonStub = stubAllPostRequests();
    const removePropertySpy = sinon.spy((command as any), 'removeProperty');

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        key: 'key1',
        debug: true
      }
    }, () => {
      try {
        assert(postCallsSpy.calledTwice === true);
        assert(removePropertySpy.calledOnce === true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call removeProperty when folder is not specified', (done) => {
    stubAllPostRequests();
    const removePropertySpy = sinon.spy((command as any), 'removeProperty');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      debug: true,
      confirm: true
    };
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002fabc"
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(removePropertySpy.calledWith(objIdentity, options));
        assert(removePropertySpy.calledOnce === true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call removeProperty when folder is specified', (done) => {
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
    const removePropertySpy = sinon.spy((command as any), 'removeProperty');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      folder: '/',
      confirm: true
    };
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/"
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(removePropertySpy.calledWith(objIdentity, options));
        assert(removePropertySpy.calledOnce === true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call removeProperty when list folder is specified', (done) => {
    stubAllPostRequests();
    const removePropertySpy = sinon.spy((command as any), 'removeProperty');
    const options = {
      webUrl: 'https://contoso.sharepoint.com/sites/abc',
      key: 'key1',
      folder: '/Shared Documents',
      confirm: true
    };
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/sites/abc/Shared Documents"
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(removePropertySpy.calledWith(objIdentity, options));
        assert(removePropertySpy.calledOnce === true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should send correct remove request data when folder is not specified', (done) => {
    const requestStub: sinon.SinonStub = stubAllPostRequests();
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      confirm: true
    };
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002fabc"
    };

    command.action(logger, { options: options } as any, () => {
      try {
        const bodyPayload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">${(options as any).key}</Parameter><Parameter Type="Null" /></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="AllProperties" /><Identity Id="198" Name="${objIdentity.objectIdentity}" /></ObjectPaths></Request>`;
        assert(requestStub.calledWith(sinon.match({ data: bodyPayload })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should send correct remove request data when folder is specified', (done) => {
    const requestStub: sinon.SinonStub = stubAllPostRequests();
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      folder: '/',
      confirm: true
    };
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/"
    };

    command.action(logger, { options: options } as any, () => {
      try {
        const bodyPayload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">${(options as any).key}</Parameter><Parameter Type="Null" /></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="Properties" /><Identity Id="198" Name="${objIdentity.objectIdentity}" /></ObjectPaths></Request>`;
        assert(requestStub.calledWith(sinon.match({ data: bodyPayload })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle requestObjectIdentity reject promise', (done) => {
    stubAllPostRequests(new Promise<any>((resolve, reject) => { return reject('requestObjectIdentity error'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      folder: '/',
      confirm: true
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('requestObjectIdentity error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle requestObjectIdentity ClientSvc error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestObjectIdentity ClientSvc error" } }]);
    stubAllPostRequests(new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      folder: '/',
      confirm: true
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('requestObjectIdentity ClientSvc error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle requestFolderObjectIdentity reject promise', (done) => {
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return reject('abc'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      folder: '/',
      confirm: true,
      debug: true
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('abc')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle requestFolderObjectIdentity ClientSvc error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestFolderObjectIdentity error" } }]);
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      verbose: true,
      confirm: true
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('requestFolderObjectIdentity error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle requestFolderObjectIdentity ClientSvc empty error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      debug: true,
      confirm: true
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should requestFolderObjectIdentity reject promise if _ObjectIdentity_ not found', (done) => {
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve('[{}]'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_parentid',
      confirm: true
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Cannot proceed. Folder _ObjectIdentity_ not found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle removeProperty reject promise response', (done) => {
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return reject('removeProperty promise error'); }));
    const removePropertySpy = sinon.spy((command as any), 'removeProperty');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      verbose: true,
      confirm: true
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert(removePropertySpy.calledOnce === true);
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('removeProperty promise error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle removeProperty ClientSvc error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "removeProperty error" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      verbose: true,
      confirm: true
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('removeProperty error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle removeProperty ClientSvc empty error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      verbose: true,
      confirm: true
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsVerboseOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsVerboseOption = true;
      }
    });
    assert(containsVerboseOption);
  });

  it('supports specifying folder', () => {
    const options = command.options;
    let containsScopeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[folder]') > -1) {
        containsScopeOption = true;
      }
    });
    assert(containsScopeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'foo',
        key: 'key1'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the key option is not specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the url option specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        key: 'key1'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the url and folder options specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        key: 'key1',
        folder: '/'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('doesn\'t fail validation if the optional folder option not specified', async () => {
    const actual = await command.validate(
      {
        options:
        {
          webUrl: 'https://contoso.sharepoint.com',
          key: 'key1'
        }
      }, commandInfo);
    assert.strictEqual(actual, true);
  });
});