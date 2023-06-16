import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./serviceprincipal-grant-revoke');

describe(commands.SERVICEPRINCIPAL_GRANT_REVOKE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SERVICEPRINCIPAL_GRANT_REVOKE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('revokes the specified permission grant (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Method Name="DeleteObject" Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionGrants" /><Method Id="13" ParentId="11" Name="GetByObjectId"><Parameters><Parameter Type="String">50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "63553a9e-101c-4000-d6f5-91ba841ffc9d"
          }, 66, {
            "IsNull": false
          }, 68, {
            "IsNull": false
          }, 70, {
            "IsNull": false
          }, 72, {
            "IsNull": false
          }, 73, {
            "IsNull": false
          }
        ]);
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, id: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg' } });
    assert(loggerLogToStderrSpy.called);
  });

  it('revokes the specified permission grant', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Method Name="DeleteObject" Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionGrants" /><Method Id="13" ParentId="11" Name="GetByObjectId"><Parameters><Parameter Type="String">50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "63553a9e-101c-4000-d6f5-91ba841ffc9d"
          }, 66, {
            "IsNull": false
          }, 68, {
            "IsNull": false
          }, 70, {
            "IsNull": false
          }, 72, {
            "IsNull": false
          }, 73, {
            "IsNull": false
          }
        ]);
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { id: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles error when revoking permission request', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      return JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": {
            "ErrorMessage": "The given key was not present in the dictionary.", "ErrorValue": null, "TraceCorrelationId": "8da23a9e-00d0-4000-c621-0ffad6315d99", "ErrorCode": -1, "ErrorTypeName": "System.Collections.Generic.KeyNotFoundException"
          }, "TraceCorrelationId": "8da23a9e-00d0-4000-c621-0ffad6315d99"
        }
      ]);
    });
    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('The given key was not present in the dictionary.'));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => { throw 'An error has occurred'; });
    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('An error has occurred'));
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('allows specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
