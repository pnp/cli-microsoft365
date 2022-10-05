import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';

const command: Command = require('./folder-rename');

describe(commands.FOLDER_RENAME, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let stubAllPostRequests: any;

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

    stubAllPostRequests = (
      requestObjectIdentityResp: any = null,
      folderObjectIdentityResp: any = null,
      folderRenameResp: any = null
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
                "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7618.1204", "ErrorInfo": null, "TraceCorrelationId": "e52c649e-a019-5000-c38d-8d334a079fd2"
              }, 27, {
                "IsNull": false
              }, 28, {
                "_ObjectIdentity_": "e52c649e-a019-5000-c38d-8d334a079fd2|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:7f1c42fe-5933-430d-bafb-6c839aa87a5c:web:30a3906a-a55e-4f48-aaae-ecf45346bf53:folder:10c46485-5035-475f-a40f-d842bab30708"
              }, 29, {
                "_ObjectType_": "SP.Folder", "_ObjectIdentity_": "e52c649e-a019-5000-c38d-8d334a079fd2|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:7f1c42fe-5933-430d-bafb-6c839aa87a5c:web:30a3906a-a55e-4f48-aaae-ecf45346bf53:folder:10c46485-5035-475f-a40f-d842bab30708", "Name": "Test2", "ServerRelativeUrl": "\u002fsites\u002fabc\u002fShared Documents\u002fTest2"
              }
            ]));
          }
        }

        // fake folder rename/move success
        if (opts.data.indexOf('Name="MoveTo"') > -1) {
          if (folderRenameResp) {
            return folderRenameResp;
          }
          else {

            return Promise.resolve(JSON.stringify([
              {
                "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7618.1204", "ErrorInfo": null, "TraceCorrelationId": "e52c649e-5023-5000-c38d-86fa815f3928"
              }
            ]));
          }
        }

        return Promise.reject('Invalid request');
      });
    };
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
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FOLDER_RENAME), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should send correct folder remove request data', async () => {
    const requestStub: sinon.SinonStub = stubAllPostRequests();
    const options = {
      webUrl: 'https://contoso.sharepoint.com/sites/abc',
      folderUrl: '/Shared Documents/Test2',
      name: 'test1',
      verbose: true,
      debug: true
    };
    const folderObjectIdentity: string = "e52c649e-a019-5000-c38d-8d334a079fd2|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:7f1c42fe-5933-430d-bafb-6c839aa87a5c:web:30a3906a-a55e-4f48-aaae-ecf45346bf53:folder:10c46485-5035-475f-a40f-d842bab30708";

    await command.action(logger, { options: options } as any);
    const bodyPayload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="MoveTo" Id="32" ObjectPathId="26"><Parameters><Parameter Type="String">/sites/abc/Shared Documents/test1</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="26" Name="${folderObjectIdentity}" /></ObjectPaths></Request>`;
    assert.strictEqual(requestStub.lastCall.args[0].data, bodyPayload);
  });

  it('should not display anything when folder removed, but not verbose', async () => {
    stubAllPostRequests();
    const options = {
      webUrl: 'https://contoso.sharepoint.com/sites/abc',
      folderUrl: '/Shared Documents/Test2',
      name: 'test1'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(loggerLogSpy.called, false);
  });

  it('should correctly handle requestObjectIdentity reject promise', async () => {
    stubAllPostRequests(new Promise<any>((resolve, reject) => { return reject('requestObjectIdentity error'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
    };
    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('requestObjectIdentity error'));
  });

  it('should correctly handle requestObjectIdentity ClientSvc error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestObjectIdentity ClientSvc error" } }]);
    stubAllPostRequests(new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
    };
    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('requestObjectIdentity ClientSvc error'));
  });

  it('should correctly handle requestFolderObjectIdentity reject promise', async () => {
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return reject('abc 1'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
    };
    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('abc 1'));
  });

  it('should correctly handle requestFolderObjectIdentity ClientSvc error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestFolderObjectIdentity error" } }]);
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
    };
    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('requestFolderObjectIdentity error'));
  });

  it('should correctly handle requestFolderObjectIdentity ClientSvc empty error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('ClientSvc unknown error'));
  });

  it('should requestFolderObjectIdentity reject promise if _ObjectIdentity_ not found', async () => {
    stubAllPostRequests(null, new Promise<any>((resolve) => { return resolve('[{}]'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'abc'
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('Cannot proceed. Folder _ObjectIdentity_ not found'));
  });

  it('should correctly handle folder remove reject promise response', async () => {
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return reject('folder remove promise error'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'abc'
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('folder remove promise error'));
  });

  it('should correctly handle folder rename ClientSvc error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "File Not Found" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'abc'
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('File Not Found'));
  });

  it('should correctly handle ClientSvc empty error response', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('ClientSvc unknown error'));
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

  it('fails validation if the webUrl option is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'abc', folderUrl: '/Shared Documents/test', name: 'abc' } }, commandInfo);
    assert.strictEqual(actual, "abc is not a valid SharePoint Online site URL");
  });

  it('passes validation when the url option specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/test',
        name: 'abc'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});