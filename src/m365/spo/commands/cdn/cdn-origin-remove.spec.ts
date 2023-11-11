import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './cdn-origin-remove.js';

describe(commands.CDN_ORIGIN_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.active = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    auth.service.spoTenantId = 'abc';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.data) {
          if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
            return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]);
          }
        }
      }

      throw 'Invalid request';
    });
    commandInfo = Cli.getCommandInfo(command);
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
    requests = [];
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore(Cli.promptForConfirmation);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
    auth.service.spoUrl = undefined;
    auth.service.spoTenantId = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CDN_ORIGIN_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes existing CDN origin from the public CDN when Public type specified without prompting with confirmation argument', async () => {
    await command.action(logger, { options: { origin: '*/cdn', force: true, type: 'Public' } });
    let deleteRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        r.headers['X-RequestDigest'] &&
        r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
        deleteRequestIssued = true;
      }
    });

    assert(deleteRequestIssued);
  });

  it('removes existing CDN origin from the private CDN when Private type specified without prompting with confirmation argument', async () => {
    await assert.rejects(command.action(logger, { options: { origin: '*/cdn', force: true, type: 'Private' } }));
    let deleteRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        r.headers['X-RequestDigest'] &&
        r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
        deleteRequestIssued = true;
      }
    });

    assert(deleteRequestIssued);
  });

  it('removes existing CDN origin from the private CDN when Private type specified without prompting with confirmation argument (debug)', async () => {
    await assert.rejects(command.action(logger, { options: { debug: true, origin: '*/cdn', force: true, type: 'Private' } }));
    let deleteRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        r.headers['X-RequestDigest'] &&
        r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
        deleteRequestIssued = true;
      }
    });

    assert(deleteRequestIssued);
  });

  it('removes existing CDN origin from the public CDN when no type specified without prompting with confirmation argument', async () => {
    await command.action(logger, { options: { origin: '*/cdn', force: true } });
    let deleteRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        r.headers['X-RequestDigest'] &&
        r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
        deleteRequestIssued = true;
      }
    });

    assert(deleteRequestIssued);
  });

  it('prompts before removing CDN origin when confirmation argument not passed', async () => {
    await command.action(logger, { options: { debug: true, origin: '*/cdn' } });

    assert(promptIssued);
  });

  it('aborts removing CDN origin when prompt not confirmed', async () => {
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { debug: true, origin: '*/cdn' } });
    assert(requests.length === 0);
  });

  it('removes CDN origin when prompt confirmed', async () => {
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { debug: true, origin: '*/cdn' } });
  });

  it('correctly handles an error when removing CDN origin', async () => {
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
          if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">*/cdn</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="abc" /></ObjectPaths></Request>`) {
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

    await assert.rejects(command.action(logger, { options: { debug: true, origin: '*/cdn', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('correctly handles a random API error', async () => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { origin: '*/cdn', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('supports suppressing confirmation prompt', () => {
    const options = command.options;
    let containsConfirmOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--force') > -1) {
        containsConfirmOption = true;
      }
    });
    assert(containsConfirmOption);
  });

  it('requires CDN origin name', () => {
    const options = command.options;
    let requiresCdnOriginName = false;
    options.forEach(o => {
      if (o.option.indexOf('<origin>') > -1) {
        requiresCdnOriginName = true;
      }
    });
    assert(requiresCdnOriginName);
  });

  it('accepts Public SharePoint Online CDN type', async () => {
    const actual = await command.validate({ options: { type: 'Public', origin: '*/CDN' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts Private SharePoint Online CDN type', async () => {
    const actual = await command.validate({ options: { type: 'Private', origin: '*/CDN' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects invalid SharePoint Online CDN type', async () => {
    const type = 'foo';
    const actual = await command.validate({ options: { type: type, origin: '*/CDN' } }, commandInfo);
    assert.strictEqual(actual, `${type} is not a valid CDN type. Allowed values are Public|Private`);
  });

  it('doesn\'t fail validation if the optional type option not specified', async () => {
    const actual = await command.validate({ options: { origin: '*/CDN' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
