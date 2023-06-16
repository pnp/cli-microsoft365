import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { IdentityResponse, spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./propertybag-set');

describe(commands.PROPERTYBAG_SET, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const stubAllPostRequests = (
    requestObjectIdentityResp: any = null,
    folderObjectIdentityResp: any = null,
    setPropertyResp: any = null,
    effectiveBasePermissionsResp: any = null
  ): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      // fake requestObjectIdentity
      if (opts.data.indexOf('3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a') > -1) {
        if (requestObjectIdentityResp) {
          return requestObjectIdentityResp;
        }
        else {
          return JSON.stringify([{
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7331.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
          }, 7, {
            "_ObjectType_": "SP.Web",
            "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
            "ServerRelativeUrl": "\u002fsites\u002fabc"
          }]);
        }
      }

      // fake requestFolderObjectIdentity
      if (opts.data.indexOf('GetFolderByServerRelativeUrl') > -1) {
        if (folderObjectIdentityResp) {
          return folderObjectIdentityResp;
        }
        else {
          return JSON.stringify([
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
          ]);
        }
      }

      // fake property set success for site and folder
      if (opts.data.indexOf('SetFieldValue') > -1) {
        if (setPropertyResp) {
          return setPropertyResp;
        }
        else {

          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7507.1203",
              "ErrorInfo": null,
              "TraceCorrelationId": "986d549e-d035-5000-2a28-c7306cd17024"
            }]);
        }
      }

      if (opts.data.indexOf('EffectiveBasePermissions') > -1) {
        if (effectiveBasePermissionsResp) {
          return effectiveBasePermissionsResp;
        }
        else {
          // effective base permissions (success case)
          return JSON.stringify([
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
          ]);
        }
      }

      throw 'Invalid request';
    });
  };

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      (command as any).setProperty,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PROPERTYBAG_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should call setProperty when folder is not specified', async () => {
    stubAllPostRequests();
    const setPropertySpy = sinon.spy((command as any), 'setProperty');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      debug: true
    };
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002fabc"
    };

    await command.action(logger, { options: options } as any);
    assert(setPropertySpy.calledWith(objIdentity, options));
    assert(setPropertySpy.calledOnce === true);
  });

  it('should call setProperty when folder is specified', async () => {
    stubAllPostRequests(JSON.stringify([{
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7331.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
    }, 7, {
      "_ObjectType_": "SP.Web",
      "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      "ServerRelativeUrl": "\u002f"
    }]));
    const setPropertySpy = sinon.spy((command as any), 'setProperty');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/'

    };
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/"
    };

    await command.action(logger, { options: options } as any);
    assert(setPropertySpy.calledWith(objIdentity, options));
    assert(setPropertySpy.calledOnce === true);
  });

  it('should call setProperty when list folder is specified', async () => {
    stubAllPostRequests();
    const setPropertySpy = sinon.spy((command as any), 'setProperty');
    const options = {
      webUrl: 'https://contoso.sharepoint.com/sites/abc',
      key: 'key1',
      value: 'value1',
      folder: '/Shared Documents'

    };
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/sites/abc/Shared Documents"
    };

    await command.action(logger, { options: options } as any);
    assert(setPropertySpy.calledWith(objIdentity, options));
    assert(setPropertySpy.calledOnce === true);
  });

  it('should send correct property set request data when folder is not specified', async () => {
    const requestStub: sinon.SinonStub = stubAllPostRequests();
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1'

    };
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002fabc"
    };

    await command.action(logger, { options: options } as any);
    const bodyPayload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">${(options as any).key}</Parameter><Parameter Type="String">${(options as any).value}</Parameter></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="AllProperties" /><Identity Id="198" Name="${objIdentity.objectIdentity}" /></ObjectPaths></Request>`;
    assert(requestStub.calledWith(sinon.match({ data: bodyPayload })));
  });

  it('should send correct property set request data when folder is specified', async () => {
    const requestStub: sinon.SinonStub = stubAllPostRequests();
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/'
    };
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/"
    };

    await command.action(logger, { options: options } as any);
    const bodyPayload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">${(options as any).key}</Parameter><Parameter Type="String">${(options as any).value}</Parameter></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="Properties" /><Identity Id="198" Name="${objIdentity.objectIdentity}" /></ObjectPaths></Request>`;
    assert(requestStub.calledWith(sinon.match({ data: bodyPayload })));
  });

  it('should correctly handle requestObjectIdentity reject promise', async () => {
    stubAllPostRequests(new Promise<any>((resolve, reject) => { return reject('requestObjectIdentity error'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/'
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('requestObjectIdentity error'));
  });

  it('should correctly handle requestObjectIdentity ClientSvc error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestObjectIdentity ClientSvc error" } }]);
    stubAllPostRequests(new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/'
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('requestObjectIdentity ClientSvc error'));
  });

  it('should correctly handle requestFolderObjectIdentity reject promise', async () => {
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return reject('abc'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      debug: true
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('abc'));
  });

  it('should correctly handle requestFolderObjectIdentity ClientSvc error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestFolderObjectIdentity error" } }]);
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      verbose: true
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('requestFolderObjectIdentity error'));
  });

  it('should correctly handle requestFolderObjectIdentity ClientSvc empty error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      debug: true
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('ClientSvc unknown error'));
  });

  it('should requestFolderObjectIdentity reject promise if _ObjectIdentity_ not found', async () => {
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve('[{}]'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_parentid',
      value: 'value1'
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('Cannot proceed. Folder _ObjectIdentity_ not found'));
  });

  it('should correctly handle isNoScriptSite = true', async () => {
    stubAllPostRequests(null, null, null, JSON.stringify([
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
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      debug: true
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('Site has NoScript enabled, and setting property bag values is not supported'));
  });

  it('should correctly handle getEffectiveBasePermissions reject promise', async () => {
    stubAllPostRequests(null, null, null, new Promise<any>((resolve, reject) => { return reject('getEffectiveBasePermissions abc'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      debug: true
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('getEffectiveBasePermissions abc'));
  });

  it('should correctly handle getEffectiveBasePermissions ClientSvc error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "getEffectiveBasePermissions error" } }]);
    stubAllPostRequests(null, null, null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      verbose: true
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('getEffectiveBasePermissions error'));
  });

  it('should correctly handle getEffectiveBasePermissions ClientSvc empty error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, null, null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      debug: true
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('ClientSvc unknown error'));
  });

  it('should getEffectiveBasePermissions reject promise if EffectiveBasePermissions not found', async () => {
    stubAllPostRequests(null, null, null, '[{}]');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folder: '/',
      key: 'vti_parentid',
      value: 'value1'
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('Cannot proceed. EffectiveBasePermissions not found'));
  });

  it('should correctly handle setProperty reject promise response', async () => {
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return reject('setProperty promise error'); }));
    const setPropertySpy = sinon.spy((command as any), 'setProperty');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      verbose: true
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('setProperty promise error'));
    assert(setPropertySpy.calledOnce === true);
  });

  it('should correctly handle setProperty ClientSvc error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "setProperty error" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      verbose: true
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('setProperty error'));
  });

  it('should correctly handle setProperty ClientSvc empty error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      key: 'key1',
      value: 'value1',
      folder: '/',
      verbose: true
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('ClientSvc unknown error'));
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
        key: 'key1',
        value: 'value1'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if value is not set', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        key: 'key1'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the key option is not set', async () => {
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
        key: 'key1',
        value: 'value1'
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
        value: 'value1',
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
          key: 'key1',
          value: 'value1'
        }
      }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
