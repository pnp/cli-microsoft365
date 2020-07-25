import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./cdn-origin-remove');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.CDN_ORIGIN_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;
  let requests: any[];
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    auth.service.tenantId = 'abc';
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.body) {
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });
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
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    requests = [];
    promptOptions = undefined;
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      request.post,
      appInsights.trackEvent,
      (command as any).getRequestDigest
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
    auth.service.tenantId = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CDN_ORIGIN_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes existing CDN origin from the public CDN when Public type specified without prompting with confirmation argument', (done) => {
    cmdInstance.action({ options: { debug: false, origin: '*/cdn', confirm: true, type: 'Public' } }, () => {
      let deleteRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
          deleteRequestIssued = true;
        }
      });

      try {
        assert(deleteRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes existing CDN origin from the private CDN when Private type specified without prompting with confirmation argument', (done) => {
    cmdInstance.action({ options: { debug: false, origin: '*/cdn', confirm: true, type: 'Private' } }, () => {
      let deleteRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
          deleteRequestIssued = true;
        }
      });

      try {
        assert(deleteRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes existing CDN origin from the private CDN when Private type specified without prompting with confirmation argument (debug)', (done) => {
    cmdInstance.action({ options: { debug: true, origin: '*/cdn', confirm: true, type: 'Private' } }, () => {
      let deleteRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
          deleteRequestIssued = true;
        }
      });

      try {
        assert(deleteRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes existing CDN origin from the public CDN when no type specified without prompting with confirmation argument', (done) => {
    cmdInstance.action({ options: { debug: false, origin: '*/cdn', confirm: true } }, () => {
      let deleteRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
          deleteRequestIssued = true;
        }
      });

      try {
        assert(deleteRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing CDN origin when confirmation argument not passed', (done) => {
    cmdInstance.action({ options: { debug: true, origin: '*/cdn' } }, () => {
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

  it('aborts removing CDN origin when prompt not confirmed', (done) => {
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: true, origin: '*/cdn' } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes CDN origin when prompt confirmed', (done) => {
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: true, origin: '*/cdn' } }, () => {
      let doneResponse = false;
      log.forEach(l => {
        if (l &&
          typeof l === 'string' &&
          l.indexOf('DONE') > -1) {
          doneResponse = true;
        }
      });

      try {
        assert(doneResponse);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles an error when removing CDN origin', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.body) {
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([
              {
                "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
                  "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
                }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
              }
            ]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, origin: '*/cdn', confirm: true } }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
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

  it('correctly handles a random API error', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action({ options: { debug: false, origin: '*/cdn', confirm: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('supports suppressing confirmation prompt', () => {
    const options = (command.options() as CommandOption[]);
    let containsConfirmOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--confirm') > -1) {
        containsConfirmOption = true;
      }
    });
    assert(containsConfirmOption);
  });

  it('requires CDN origin name', () => {
    const options = (command.options() as CommandOption[]);
    let requiresCdnOriginName = false;
    options.forEach(o => {
      if (o.option.indexOf('<origin>') > -1) {
        requiresCdnOriginName = true;
      }
    });
    assert(requiresCdnOriginName);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return []; });
    const options = (command.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('accepts Public SharePoint Online CDN type', () => {
    const actual = (command.validate() as CommandValidate)({ options: { type: 'Public' } });
    assert.strictEqual(actual, true);
  });

  it('accepts Private SharePoint Online CDN type', () => {
    const actual = (command.validate() as CommandValidate)({ options: { type: 'Private' } });
    assert.strictEqual(actual, true);
  });

  it('rejects invalid SharePoint Online CDN type', () => {
    const type = 'foo';
    const actual = (command.validate() as CommandValidate)({ options: { type: type } });
    assert.strictEqual(actual, `${type} is not a valid CDN type. Allowed values are Public|Private`);
  });

  it('doesn\'t fail validation if the optional type option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.strictEqual(actual, true);
  });
});