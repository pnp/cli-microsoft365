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
const command: Command = require('./cdn-policy-set');

describe(commands.CDN_POLICY_SET, () => {
  let log: string[];
  let logger: Logger;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    auth.service.tenantId = 'abc';
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.data) {
          if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetTenantCdnPolicy" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Enum">0</Parameter><Parameter Type="String">WOFF</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });
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
    requests = [];
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      request.post,
      appInsights.trackEvent,
      spo.getRequestDigest
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
    auth.service.tenantId = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CDN_POLICY_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets IncludeFileExtensions CDN policy on the public CDN when Public type specified', (done) => {
    command.action(logger, { options: { debug: true, policy: 'IncludeFileExtensions', value: 'WOFF', type: 'Public' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetTenantCdnPolicy" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Enum">0</Parameter><Parameter Type="String">WOFF</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
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

  it('sets IncludeFileExtensions CDN policy on the private CDN when Private type specified', (done) => {
    command.action(logger, { options: { debug: true, policy: 'IncludeFileExtensions', value: 'WOFF', type: 'Private' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetTenantCdnPolicy" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Enum">0</Parameter><Parameter Type="String">WOFF</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
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

  it('sets IncludeFileExtensions CDN policy on the public CDN when no type specified', (done) => {
    command.action(logger, { options: { debug: true, policy: 'IncludeFileExtensions', value: 'WOFF' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetTenantCdnPolicy" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Enum">0</Parameter><Parameter Type="String">WOFF</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
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

  it('sets ExcludeRestrictedSiteClassifications CDN policy on the public CDN when no type specified', (done) => {
    command.action(logger, { options: { debug: false, policy: 'ExcludeRestrictedSiteClassifications', value: 'foo' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetTenantCdnPolicy" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Enum">1</Parameter><Parameter Type="String">foo</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
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

  it('escapes XML in user input', (done) => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.data) {
          if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetTenantCdnPolicy" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Enum">0</Parameter><Parameter Type="String">&lt;WOFF</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, policy: 'IncludeFileExtensions', value: '<WOFF' } }, () => {
      try {
        assert.strictEqual(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles an error when setting tenant CDN policy value', (done) => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.data) {
          if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetTenantCdnPolicy" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Enum">0</Parameter><Parameter Type="String">&lt;WOFF</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
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

    command.action(logger, { options: { debug: false, policy: 'IncludeFileExtensions', value: '<WOFF' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore(request.post);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('requires CDN policy name', () => {
    const options = command.options();
    let requiresCdnPolicyName = false;
    options.forEach(o => {
      if (o.option.indexOf('<policy>') > -1) {
        requiresCdnPolicyName = true;
      }
    });
    assert(requiresCdnPolicyName);
  });

  it('requires CDN policy value', () => {
    const options = command.options();
    let requiresCdnPolicyValue = false;
    options.forEach(o => {
      if (o.option.indexOf('<value>') > -1) {
        requiresCdnPolicyValue = true;
      }
    });
    assert(requiresCdnPolicyValue);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return []; });
    const options = command.options();
    sinonUtil.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('accepts Public SharePoint Online CDN type', () => {
    const actual = command.validate({ options: { type: 'Public', policy: 'IncludeFileExtensions' } });
    assert.strictEqual(actual, true);
  });

  it('accepts Private SharePoint Online CDN type', () => {
    const actual = command.validate({ options: { type: 'Private', policy: 'IncludeFileExtensions' } });
    assert.strictEqual(actual, true);
  });

  it('rejects invalid SharePoint Online CDN type', () => {
    const type = 'foo';
    const actual = command.validate({ options: { type: type, policy: 'IncludeFileExtensions' } });
    assert.strictEqual(actual, `${type} is not a valid CDN type. Allowed values are Public|Private`);
  });

  it('doesn\'t fail validation if the optional type option not specified', () => {
    const actual = command.validate({ options: { policy: 'IncludeFileExtensions' } });
    assert.strictEqual(actual, true);
  });

  it('accepts IncludeFileExtensions SharePoint Online CDN policy', () => {
    const actual = command.validate({ options: { policy: 'IncludeFileExtensions' } });
    assert.strictEqual(actual, true);
  });

  it('accepts ExcludeRestrictedSiteClassifications SharePoint Online CDN policy', () => {
    const actual = command.validate({ options: { policy: 'ExcludeRestrictedSiteClassifications' } });
    assert.strictEqual(actual, true);
  });

  it('rejects invalid SharePoint Online CDN policy', () => {
    const policy = 'foo';
    const actual = command.validate({ options: { policy: policy } });
    assert.strictEqual(actual, `${policy} is not a valid CDN policy. Allowed values are IncludeFileExtensions|ExcludeRestrictedSiteClassifications`);
  });
});