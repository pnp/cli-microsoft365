import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./cdn-set');

describe(commands.CDN_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.data) {
          if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });
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
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CDN_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('enables public CDN when Public type specified and enabled set to true', (done) => {
    command.action(logger, { options: { debug: false, enabled: 'true', type: 'Public' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables public CDN when Public type specified and enabled set to true (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'true', type: 'Public' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables public CDN when Public type specified and enabled set to true without default origins (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'true', type: 'Public', noDefaultOrigins: true } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables public CDN when no type specified and enabled set to true', (done) => {
    command.action(logger, { options: { debug: false, enabled: true } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables private CDN when Private type specified and enabled set to true (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'true', type: 'Private' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables private CDN when Private type specified and enabled set to true without default origins (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'true', type: 'Private', noDefaultOrigins: true } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables both CDN\'s when Both type specified and enabled set to true (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'true', type: 'Both' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="96" ObjectPathId="95" /><Method Name="SetTenantCdnEnabled" Id="97" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="98" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="99" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="100" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="95" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables both CDN\'s when Both type specified and enabled set to true, with default origins', (done) => {
    command.action(logger, { options: { enabled: 'true', type: 'Both', noDefaultOrigins: false } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="96" ObjectPathId="95" /><Method Name="SetTenantCdnEnabled" Id="97" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="98" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="99" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="100" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="95" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables both CDN\'s when Both type specified and enabled set to true, with origins (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'true', type: 'Both', noDefaultOrigins: false } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="96" ObjectPathId="95" /><Method Name="SetTenantCdnEnabled" Id="97" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="98" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="99" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="100" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="95" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables both CDN\'s when Both type specified and enabled set to true, without Default Origins', (done) => {
    command.action(logger, { options: { enabled: 'true', type: 'Both', noDefaultOrigins: true } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><Method Name="SetTenantCdnEnabled" Id="13" ObjectPathId="11"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="14" ObjectPathId="11"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables both CDN\'s when Both type specified and enabled set to true, without default origins (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'true', type: 'Both', noDefaultOrigins: true } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><Method Name="SetTenantCdnEnabled" Id="13" ObjectPathId="11"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="14" ObjectPathId="11"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disable both CDN\'s when Both type specified and enabled set to false, with default (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'false', type: 'Both', noDefaultOrigins: false } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="96" ObjectPathId="95" /><Method Name="SetTenantCdnEnabled" Id="97" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="98" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="95" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disable both CDN\'s when Both type specified and enabled set to false, without default origins (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'false', type: 'Both', noDefaultOrigins: true } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><Method Name="SetTenantCdnEnabled" Id="13" ObjectPathId="11"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="14" ObjectPathId="11"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disables both CDN\'s when Both type specified and enabled set to false', (done) => {
    command.action(logger, { options: { debug: false, enabled: 'false', type: 'Both' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="96" ObjectPathId="95" /><Method Name="SetTenantCdnEnabled" Id="97" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="98" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="95" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disables both CDN\'s when Both type specified and enabled set to false (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'false', type: 'Both' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="96" ObjectPathId="95" /><Method Name="SetTenantCdnEnabled" Id="97" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="98" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="95" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disables public CDN when Public type specified and enabled set to false (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'false', type: 'Public' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disables public CDN when Public type specified and enabled set to false and noDefaultOrigins is passed (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'false', type: 'Public', noDefaultOrigins: true } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disables public CDN when Public type specified and enabled set to false and noDefaultOrigins is passed', (done) => {
    command.action(logger, { options: { debug: false, enabled: 'false', type: 'Public', noDefaultOrigins: true } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disables public CDN when Public type specified and enabled set to false', (done) => {
    command.action(logger, { options: { debug: false, enabled: 'false', type: 'Public' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disables Private CDN when Private type specified and enabled set to false (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'false', type: 'Private' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disables Private CDN when Private type specified and enabled set to false and noDefaultOrigins is passed (debug)', (done) => {
    command.action(logger, { options: { debug: true, enabled: 'false', type: 'Private', noDefaultOrigins: true } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disables Private CDN when Private type specified and enabled set to false and noDefaultOrigins is passed', (done) => {
    command.action(logger, { options: { debug: false, enabled: 'false', type: 'Private', noDefaultOrigins: true } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disables Private CDN when Private type specified and enabled set to false', (done) => {
    command.action(logger, { options: { debug: false, enabled: 'false', type: 'Private' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('correctly handles an error when setting tenant CDN settings', (done) => {
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
          if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

    command.action(logger, { options: { debug: false, enabled: 'true' } } as any, (err?: any) => {
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
    const options = command.options;
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('requires tenant enabled state', () => {
    const options = command.options;
    let requiresOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<enabled>') > -1) {
        requiresOption = true;
      }
    });
    assert(requiresOption);
  });

  it('accepts Public SharePoint Online CDN type', async () => {
    const actual = await command.validate({ options: { type: 'Public', enabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts Private SharePoint Online CDN type', async () => {
    const actual = await command.validate({ options: { type: 'Private', enabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts Both SharePoint Online CDN type', async () => {
    const actual = await command.validate({ options: { type: 'Both', enabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects invalid SharePoint Online CDN type', async () => {
    const type = 'foo';
    const actual = await command.validate({ options: { type: type, enabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, `${type} is not a valid CDN type. Allowed values are Public|Private|Both`);
  });

  it('doesn\'t fail validation if the optional type option not specified', async () => {
    const actual = await command.validate({ options: { enabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts true SharePoint Online CDN enabled state', async () => {
    const actual = await command.validate({ options: { enabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('does not accept True SharePoint Online CDN enabled state', async () => {
    const actual = await command.validate({ options: { enabled: 'True' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('does not accept TRUE SharePoint Online CDN enabled state', async () => {
    const actual = await command.validate({ options: { enabled: 'TRUE' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('accepts false SharePoint Online CDN enabled state', async () => {
    const actual = await command.validate({ options: { enabled: 'false' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('does not accept False SharePoint Online CDN enabled state', async () => {
    const actual = await command.validate({ options: { enabled: 'False' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('does not accept FALSE SharePoint Online CDN enabled state', async () => {
    const actual = await command.validate({ options: { enabled: 'FALSE' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('rejects invalid SharePoint Online CDN enabled state', async () => {
    const enabled = 'foo';
    const actual = await command.validate({ options: { enabled: enabled } }, commandInfo);
    assert.strictEqual(actual, `${enabled} is not a valid boolean value. Allowed values are true|false`);
  });
});