import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./cdn-origin-add');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.CDN_ORIGIN_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  let requests: any[];

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
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddTenantCdnOrigin" Id="27" ObjectPathId="23"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="23" Name="abc" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([
              {
                "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "a05d299e-0036-4000-8546-cfc42dc07fd2"
              }, 42, [
                "*\u002fMASTERPAGE", "*\u002fSTYLE LIBRARY", "*\u002fCLIENTSIDEASSETS", "*\u002fCDN (configuration pending)"
              ]
            ]));
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
      }
    };
    requests = [];
  });

  afterEach(() => {
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
    assert.strictEqual(command.name.startsWith(commands.CDN_ORIGIN_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets CDN origin on the public CDN when Public type specified', (done) => {
    cmdInstance.action({ options: { debug: true, origin: '*/cdn', type: 'Public' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddTenantCdnOrigin" Id="27" ObjectPathId="23"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="23" Name="abc" /></ObjectPaths></Request>`) {
          setRequestIssued = true;
        }
      });

      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets CDN origin on the private CDN when Private type specified', (done) => {
    cmdInstance.action({ options: { debug: true, origin: '*/cdn', type: 'Private' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddTenantCdnOrigin" Id="27" ObjectPathId="23"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="23" Name="abc" /></ObjectPaths></Request>`) {
          setRequestIssued = true;
        }
      });

      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets CDN origin on the public CDN when no type specified', (done) => {
    cmdInstance.action({ options: { debug: false, origin: '*/cdn' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddTenantCdnOrigin" Id="27" ObjectPathId="23"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="23" Name="abc" /></ObjectPaths></Request>`) {
          setRequestIssued = true;
        }
      });

      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles trying to set CDN origin that has already been set', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

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
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddTenantCdnOrigin" Id="27" ObjectPathId="23"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="23" Name="abc" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([
              {
                "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
                  "ErrorMessage": "The library is already registered as a CDN origin.", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
                }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
              }
            ]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, origin: '*/cdn', type: 'Public' } }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'The library is already registered as a CDN origin.');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });
    cmdInstance.action({ options: { debug: true, origin: '*/cdn', type: 'Public' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in user input', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

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
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddTenantCdnOrigin" Id="27" ObjectPathId="23"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">&lt;*/CDN&gt;</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="23" Name="abc" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([
              {
                "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "a05d299e-0036-4000-8546-cfc42dc07fd2"
              }, 42, [
                "*\u002fMASTERPAGE", "*\u002fSTYLE LIBRARY", "*\u002fCLIENTSIDEASSETS", "*\u002fCDN (configuration pending)"
              ]
            ]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, origin: '<*/CDN>' } }, () => {
      let isDone = false;
      log.forEach(l => {
        if (l && typeof l === 'string' && l.indexOf('DONE')) {
          isDone = true;
        }
      });

      try {
        assert(isDone);
        done();
      }
      catch (e) {
        done(e);
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