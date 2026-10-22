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
import command, { options } from './cdn-policy-list.js';

describe(commands.CDN_POLICY_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    auth.connection.spoTenantId = 'abc';
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
    loggerLogSpy = sinon.spy(logger, 'log');
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
    assert.strictEqual(command.name, commands.CDN_POLICY_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves the policies for the public CDN when type set to Public', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { FormDigestValue: 'abc' };
        }
      }

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnPolicies" Id="7" ObjectPathId="3"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="3" Name="abc" /></ObjectPaths></Request>`) {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "0c93299e-7040-4000-83c8-74ba9192542c" }, 7, ["IncludeFileExtensions;CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON", "ExcludeRestrictedSiteClassifications; ", "ExcludeIfNoScriptDisabled;False"]]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, cdnType: 'Public' }) });
    assert(loggerLogSpy.calledWith([{
      Policy: 'IncludeFileExtensions',
      Value: 'CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON'
    },
    { Policy: 'ExcludeRestrictedSiteClassifications', Value: ' ' },
    { Policy: 'ExcludeIfNoScriptDisabled', Value: 'False' }]));
  });

  it('retrieves the policies for the private CDN when type set to Private', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { FormDigestValue: 'abc' };
        }
      }

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnPolicies" Id="7" ObjectPathId="3"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="3" Name="abc" /></ObjectPaths></Request>`) {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "0c93299e-7040-4000-83c8-74ba9192542c" }, 7, ["IncludeFileExtensions;CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON", "ExcludeRestrictedSiteClassifications; ", "ExcludeIfNoScriptDisabled;False"]]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, cdnType: 'Private' }) });
    assert(loggerLogSpy.calledWith([{
      Policy: 'IncludeFileExtensions',
      Value: 'CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON'
    },
    { Policy: 'ExcludeRestrictedSiteClassifications', Value: ' ' },
    { Policy: 'ExcludeIfNoScriptDisabled', Value: 'False' }]));
  });

  it('retrieves the policies for the public CDN when no type set', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { FormDigestValue: 'abc' };
        }
      }

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnPolicies" Id="7" ObjectPathId="3"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="3" Name="abc" /></ObjectPaths></Request>`) {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "0c93299e-7040-4000-83c8-74ba9192542c" }, 7, ["IncludeFileExtensions;CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON", "ExcludeRestrictedSiteClassifications; ", "ExcludeIfNoScriptDisabled;False"]]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert(loggerLogSpy.calledWith([{
      Policy: 'IncludeFileExtensions',
      Value: 'CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON'
    },
    { Policy: 'ExcludeRestrictedSiteClassifications', Value: ' ' },
    { Policy: 'ExcludeIfNoScriptDisabled', Value: 'False' }]));
  });

  it('correctly handles an error when retrieving tenant CDN policies', async () => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { FormDigestValue: 'abc' };
        }
      }

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.data) {
          if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnPolicies" Id="7" ObjectPathId="3"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="3" Name="abc" /></ObjectPaths></Request>`) {
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError('An error has occurred'));
  });

  it('correctly handles a random API error', async () => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError('An error has occurred'));
  });

  it('passes validation with no options', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });

  it('accepts Public SharePoint Online CDN type', () => {
    const actual = commandOptionsSchema.safeParse({ cdnType: 'Public' });
    assert.strictEqual(actual.success, true);
  });

  it('accepts Private SharePoint Online CDN type', () => {
    const actual = commandOptionsSchema.safeParse({ cdnType: 'Private' });
    assert.strictEqual(actual.success, true);
  });

  it('rejects invalid SharePoint Online CDN type', () => {
    const actual = commandOptionsSchema.safeParse({ cdnType: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });
});
