import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./orgassetslibrary-add');

describe(commands.ORGASSETSLIBRARY_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      spo.getRequestDigest
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ORGASSETSLIBRARY_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds a new library as org assets library (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddToOrgAssetsLibAndCdnWithType" Id="11" ObjectPathId="8"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets</Parameter><Parameter Type="Null" /><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify(
          [{
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19708.12061", "ErrorInfo": null, "TraceCorrelationId": "a0a8309f-4039-a000-ea81-9b8297eb43e0"
          }]
        ));
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, libraryUrl: 'https://contoso.sharepoint.com/siteassets' } });
  });

  it('adds a new library as org assets library with CDN Type (debug)', async () => {
    let orgAssetLibAddCallIssued = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddToOrgAssetsLibAndCdnWithType" Id="11" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets</Parameter><Parameter Type="Null" /><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        orgAssetLibAddCallIssued = true;

        return Promise.resolve(JSON.stringify(
          [{
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19708.12061", "ErrorInfo": null, "TraceCorrelationId": "a0a8309f-4039-a000-ea81-9b8297eb43e0"
          }]
        ));
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, libraryUrl: 'https://contoso.sharepoint.com/siteassets', cdnType: 'Public' } });
    assert(orgAssetLibAddCallIssued);
  });

  it('adds a new library as org assets library with CDN Type and thumbnailUrl (debug)', async () => {
    let orgAssetLibAddCallIssued = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddToOrgAssetsLibAndCdnWithType" Id="11" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets/logo.png</Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        orgAssetLibAddCallIssued = true;

        return Promise.resolve(JSON.stringify(
          [{
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19708.12061", "ErrorInfo": null, "TraceCorrelationId": "a0a8309f-4039-a000-ea81-9b8297eb43e0"
          }]
        ));
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, libraryUrl: 'https://contoso.sharepoint.com/siteassets', cdnType: 'Public', thumbnailUrl: 'https://contoso.sharepoint.com/siteassets/logo.png' } });
    assert(orgAssetLibAddCallIssued);
  });

  it('handles error if is already present', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddToOrgAssetsLibAndCdnWithType" Id="11" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets/logo.png</Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19708.12061", "ErrorInfo": {
                "ErrorMessage": "This library is already an organization assets library.", "ErrorValue": null, "TraceCorrelationId": "aba8309f-d0d9-a000-ea81-916572c2fbeb", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "aba8309f-d0d9-a000-ea81-916572c2fbeb"
            }
          ]
        ));
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: true, libraryUrl: 'https://contoso.sharepoint.com/siteassets', cdnType: 'Public', thumbnailUrl: 'https://contoso.sharepoint.com/siteassets/logo.png' } } as any),
      new CommandError(`This library is already an organization assets library.`));
  });

  it('handles error getting request', async () => {
    const svcListRequest = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
              "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
            }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true
      }
    } as any), new CommandError('An error has occurred'));
    assert(svcListRequest.called);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if the libraryUrl is not valid', async () => {
    const actual = await command.validate({ options: { libraryUrl: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the thumbnail is not valid', async () => {
    const actual = await command.validate({ options: { libraryUrl: 'https://contoso.sharepoint.com/siteassets', thumbnailUrl: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the libraryUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { libraryUrl: 'https://contoso.sharepoint.com/siteassets' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
