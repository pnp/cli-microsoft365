import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { IdentityResponse, spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command, { options } from './propertybag-set.js';
import { SpoPropertyBagBaseCommand } from './propertybag-base.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.PROPERTYBAG_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const stubAllPostRequests = (
    requestObjectIdentityResp: any = null,
    folderObjectIdentityResp: any = null,
    setPropertyResp: any = null
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
                "_ObjectType_": "SP.PropertyValues", "vti_folderitemcount$  Int32": 0, "vti_level$  Int32": 1, "vti_parentid": "{1C5271C8-DB93-459E-9C18-68FC33EFD856}", "vti_winfileattribs": "00000012", "vti_candeleteversion": "true", "vti_foldersubfolderitemcount$  Int32": 0, "vti_timelastmodified": "/Date(2017,10,7,11,29,31,0)/", "vti_dirlateststamp": "/Date(2018,1,12,22,34,31,0)/", "vti_isscriptable": "false", "vti_isexecutable": "false", "vti_metainfoversion$  Int32": 1, "vti_isbrowsable": "true", "vti_timecreated": "/Date(2017,10,7,11,29,31,0)/", "vti_etag": "\"{DF4291DE-226F-4C39-BBCC-DF21915F5FC1},256\"", "vti_hassubdirs": "true", "vti_docstoreversion$  Int32": 256, "vti_rtag": "rt:DF4291DE-226F-4C39-BBCC-DF21915F5FC1@00000000256", "vti_docstoretype$  Int32": 1, "vti_replid": "rid:{DF4291DE-226F-4C39-BBCC-DF21915F5FC1}"
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

      throw 'Invalid request';
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      (command as any).setProperty,
      cli.getSettingWithDefaultValue,
      spo.getEffectiveBasePermissions
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
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
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002fabc"
    };
    const parsedOptions = commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', debug: true });

    await command.action(logger, { options: parsedOptions });
    assert(setPropertySpy.calledWith(objIdentity, parsedOptions));
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
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/"
    };
    const parsedOptions = commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', folder: '/' });

    await command.action(logger, { options: parsedOptions });
    assert(setPropertySpy.calledWith(objIdentity, parsedOptions));
    assert(setPropertySpy.calledOnce === true);
  });

  it('should call setProperty when list folder is specified', async () => {
    stubAllPostRequests();
    const setPropertySpy = sinon.spy((command as any), 'setProperty');
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/sites/abc/Shared Documents"
    };
    const parsedOptions = commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/abc', key: 'key1', value: 'value1', folder: '/Shared Documents' });

    await command.action(logger, { options: parsedOptions });
    assert(setPropertySpy.calledWith(objIdentity, parsedOptions));
    assert(setPropertySpy.calledOnce === true);
  });

  it('should send correct property set request data when folder is not specified', async () => {
    const requestStub: sinon.SinonStub = stubAllPostRequests();
    const objIdentity: IdentityResponse = {
      objectIdentity: "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
      serverRelativeUrl: "\u002fsites\u002fabc"
    };

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1' }) });
    const bodyPayload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">key1</Parameter><Parameter Type="String">value1</Parameter></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="AllProperties" /><Identity Id="198" Name="${objIdentity.objectIdentity}" /></ObjectPaths></Request>`;
    assert(requestStub.calledWith(sinon.match({ data: bodyPayload })));
  });

  it('should send correct property set request data when folder is specified', async () => {
    const requestStub: sinon.SinonStub = stubAllPostRequests();
    const objIdentity: IdentityResponse = {
      objectIdentity: "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1",
      serverRelativeUrl: "/"
    };

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', folder: '/' }) });
    const bodyPayload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">key1</Parameter><Parameter Type="String">value1</Parameter></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="Properties" /><Identity Id="198" Name="${objIdentity.objectIdentity}" /></ObjectPaths></Request>`;
    assert(requestStub.calledWith(sinon.match({ data: bodyPayload })));
  });

  it('should correctly handle requestObjectIdentity reject promise', async () => {
    stubAllPostRequests(new Promise<any>((resolve, reject) => { return reject('requestObjectIdentity error'); }));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', folder: '/' }) }),
      new CommandError('requestObjectIdentity error'));
  });

  it('should correctly handle requestObjectIdentity ClientSvc error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestObjectIdentity ClientSvc error" } }]);
    stubAllPostRequests(new Promise<any>((resolve) => { return resolve(error); }));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', folder: '/' }) }),
      new CommandError('requestObjectIdentity ClientSvc error'));
  });

  it('should correctly handle requestFolderObjectIdentity reject promise', async () => {
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return reject('abc'); }));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', folder: '/', debug: true }) }),
      new CommandError('abc'));
  });

  it('should correctly handle requestFolderObjectIdentity ClientSvc error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestFolderObjectIdentity error" } }]);
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve(error); }));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', folder: '/', verbose: true }) }),
      new CommandError('requestFolderObjectIdentity error'));
  });

  it('should correctly handle requestFolderObjectIdentity ClientSvc empty error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve(error); }));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', folder: '/', debug: true }) }),
      new CommandError('ClientSvc unknown error'));
  });

  it('should requestFolderObjectIdentity reject promise if _ObjectIdentity_ not found', async () => {
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve('[{}]'); }));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', folder: '/', key: 'vti_parentid', value: 'value1' }) }),
      new CommandError('Cannot proceed. Folder _ObjectIdentity_ not found'));
  });

  it('should show a tip when Access Denied error with E_ACCESSDENIED is thrown', async () => {
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return reject('Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED))'); }));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1' }) }),
      new CommandError('Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED))'));
    assert(log.some(l => l === 'Tip: If a site has NoScript enabled, setting the property bag value may result in exceptions.'));
  });

  it('should not show a tip when a non-Access Denied error is thrown', async () => {
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return reject('Some other error'); }));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1' }) }),
      new CommandError('Some other error'));
    assert(!log.some(l => l === 'Tip: If a site has NoScript enabled, setting the property bag value may result in exceptions.'));
  });

  it('should correctly handle setProperty reject promise response', async () => {
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return reject('setProperty promise error'); }));
    const setPropertySpy = sinon.spy((command as any), 'setProperty');

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', folder: '/', verbose: true }) }),
      new CommandError('setProperty promise error'));
    assert(setPropertySpy.calledOnce === true);
  });

  it('should correctly handle setProperty ClientSvc error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "setProperty error" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve) => { return resolve(error); }));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', folder: '/', verbose: true }) }),
      new CommandError('setProperty error'));
  });

  it('should correctly handle setProperty ClientSvc empty error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve) => { return resolve(error); }));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', folder: '/', verbose: true }) }),
      new CommandError('ClientSvc unknown error'));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', key: 'key1', value: 'value1' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation if value is not set', () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the key option is not set', () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when the url option specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when the url and folder options specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1', folder: '/' });
    assert.strictEqual(actual.success, true);
  });

  it('doesn\'t fail validation if the optional folder option not specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', key: 'key1', value: 'value1' });
    assert.strictEqual(actual.success, true);
  });

  it('should correctly detect a NoScript site', async () => {
    sinon.stub(spo, 'getEffectiveBasePermissions').resolves({
      high: 2147483647,
      low: 4294705151,
      has: () => false
    } as any);

    const identityResp: IdentityResponse = {
      objectIdentity: 'identity',
      serverRelativeUrl: '/'
    };

    const result = await SpoPropertyBagBaseCommand.isNoScriptSite('https://contoso.sharepoint.com', 'formDigest', identityResp, logger, false);
    assert.strictEqual(result, true);
  });

  it('should correctly detect a non-NoScript site', async () => {
    sinon.stub(spo, 'getEffectiveBasePermissions').resolves({
      high: 2147483647,
      low: 4294967295,
      has: () => true
    } as any);

    const identityResp: IdentityResponse = {
      objectIdentity: 'identity',
      serverRelativeUrl: '/'
    };

    const result = await SpoPropertyBagBaseCommand.isNoScriptSite('https://contoso.sharepoint.com', 'formDigest', identityResp, logger, false);
    assert.strictEqual(result, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', key: 'abc', value: 'def', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });
});

