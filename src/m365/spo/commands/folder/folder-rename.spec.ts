import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';

const command: Command = require('./folder-rename');

describe(commands.FOLDER_RENAME, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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

  it('should send correct folder remove request data', (done) => {
    const requestStub: sinon.SinonStub = stubAllPostRequests();
    const options = {
      webUrl: 'https://contoso.sharepoint.com/sites/abc',
      folderUrl: '/Shared Documents/Test2',
      name: 'test1',
      verbose: true,
      debug: true
    };
    const folderObjectIdentity: string = "e52c649e-a019-5000-c38d-8d334a079fd2|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:7f1c42fe-5933-430d-bafb-6c839aa87a5c:web:30a3906a-a55e-4f48-aaae-ecf45346bf53:folder:10c46485-5035-475f-a40f-d842bab30708";

    command.action(logger, { options: options } as any, () => {
      try {
        const bodyPayload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="MoveTo" Id="32" ObjectPathId="26"><Parameters><Parameter Type="String">/sites/abc/Shared Documents/test1</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="26" Name="${folderObjectIdentity}" /></ObjectPaths></Request>`;
        assert.strictEqual(requestStub.lastCall.args[0].data, bodyPayload);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should not display anything when folder removed, but not verbose', (done) => {
    stubAllPostRequests();
    const options = {
      webUrl: 'https://contoso.sharepoint.com/sites/abc',
      folderUrl: '/Shared Documents/Test2',
      name: 'test1'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.called, false);
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
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
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
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
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
    stubAllPostRequests(null, new Promise<any>((resolve, reject) => { return reject('abc 1'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
    };
    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('abc 1')));
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
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
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
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
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
      folderUrl: '/Shared Documents/test',
      name: 'abc'
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

  it('should correctly handle folder remove reject promise response', (done) => {
    stubAllPostRequests(null, null, new Promise<any>((resolve, reject) => { return reject('folder remove promise error'); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'abc'
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('folder remove promise error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle folder rename ClientSvc error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "File Not Found" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'abc'
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle ClientSvc empty error response', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
    stubAllPostRequests(null, null, new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      folderUrl: '/Shared Documents/test',
      name: 'test1',
      verbose: true
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
    const options = command.options();
    let containsVerboseOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsVerboseOption = true;
      }
    });
    assert(containsVerboseOption);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return []; });
    const options = command.options();
    sinonUtil.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('fails validation if the webUrl option is not valid', () => {
    const actual = command.validate({ options: { webUrl: 'abc' } });
    assert.strictEqual(actual, "abc is not a valid SharePoint Online site URL");
  });

  it('passes validation when the url option specified', () => {
    const actual = command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/test',
        name: 'abc'
      }
    });
    assert.strictEqual(actual, true);
  });
});