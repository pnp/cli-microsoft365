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
import commands from '../../commands.js';
import command from './cdn-get.js';

describe(commands.CDN_GET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    auth.connection.spoTenantId = 'abc';
    commandInfo = cli.getCommandInfo(command);
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
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
    auth.connection.spoTenantId = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CDN_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves the settings of the public CDN when type set to Public', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/contextinfo`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { FormDigestValue: 'abc' };
        }
      }

      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
        }
      }

      throw 'Invalid request ' + opts.url;
    });

    await command.action(logger, { options: { verbose: true, type: 'Public' } });
    assert(loggerLogToStderrSpy.calledWithExactly(`Public CDN at ${spoAdminUrl} is enabled`));
  });

  it('retrieves the settings of the private CDN when type set to Private', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/contextinfo`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { FormDigestValue: 'abc' };
        }
      }

      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, false]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'Private' } });
    assert(loggerLogSpy.calledOnceWithExactly(false));
  });

  it('retrieves the settings of the CDN when using verbose mode', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/contextinfo`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { FormDigestValue: 'abc' };
        }
      }

      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, false]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly(false));
  });

  it('retrieves the settings of the private CDN when type set to Private (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/contextinfo`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { FormDigestValue: 'abc' };
        }
      }

      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, false]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, type: 'Private' } });
    assert(loggerLogToStderrSpy.calledWithExactly(`Private CDN at ${spoAdminUrl} is disabled`));
  });

  it('retrieves the settings of the public CDN when no type set', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/contextinfo`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { FormDigestValue: 'abc' };
        }
      }

      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogToStderrSpy.calledWithExactly(`Public CDN at ${spoAdminUrl} is enabled`));
  });

  it('correctly handles an error when getting tenant CDN settings', async () => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/contextinfo`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { FormDigestValue: 'abc' };
        }
      }

      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.data) {
          if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
            return JSON.stringify([
              {
                "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
                  "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
                }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
              }
            ]);
          }
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError('An error has occurred'));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('accepts Public SharePoint Online CDN type', async () => {
    const actual = await command.validate({ options: { type: 'Public' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts Private SharePoint Online CDN type', async () => {
    const actual = await command.validate({ options: { type: 'Private' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects invalid SharePoint Online CDN type', async () => {
    const actual = await command.validate({ options: { type: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('doesn\'t fail validation if the optional type option not specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
