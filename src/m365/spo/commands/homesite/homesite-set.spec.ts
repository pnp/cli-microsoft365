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
const command: Command = require('./homesite-set');

describe(commands.HOMESITE_SET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
      telemetry.trackEvent,
      pid.getProcessName,
      spo.getRequestDigest
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HOMESITE_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets the specified site as the Home Site', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="57" ObjectPathId="56" /><Method Name="SetSPHSite" Id="58" ObjectPathId="56"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Work</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="56" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": null, "TraceCorrelationId": "e4f2e59e-c0a9-0000-3dd0-1d8ef12cc742"
            }, 57, {
              "IsNull": false
            }, 58, "The Home site has been set to https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fWork."
          ]
        ));
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/Work"
      }
    } as any);
    assert(loggerLogSpy.calledWith('The Home site has been set to https://contoso.sharepoint.com/sites/Work.'));
  });

  it('sets the specified site as the Home Site and sets the Viva Connections default experience to True', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="ValidateMultipleHomeSitesParameterExists" Id="85" ObjectPathId="81"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Name="ValidateVivaHomeParameterExists" Id="86" ObjectPathId="81"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="SetSPHSiteWithConfigurations" Id="87" ObjectPathId="81"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Work</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="81" Name="b6e793a0-e066-6000-3c4a-cb1f897402b4|908bed80-a04a-4433-b4a0-883d9847d110:d872ec63-6bea-4678-9429-078f4fa93560&#xA;Tenant" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": null, "TraceCorrelationId": "e4f2e59e-c0a9-0000-3dd0-1d8ef12cc742"
            }, 86, {
              "IsNull": false
            }, 87, "The Home site has been set to https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fWork and the Viva Connections default experience to True. It may take some time for the change to apply. Check aka.ms/homesites for details."
          ]
        ));
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/Work",
        VivaConnectionsDefaultStart: true
      }
    } as any);
    assert(loggerLogSpy.calledWith('The Home site has been set to https://contoso.sharepoint.com/sites/Work and the Viva Connections default experience to True. It may take some time for the change to apply. Check aka.ms/homesites for details.'));
  });

  it('correctly handles error when setting the Home Site', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="57" ObjectPathId="56" /><Method Name="SetSPHSite" Id="58" ObjectPathId="56"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Work</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="56" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": {
                "ErrorMessage": "The provided site url can't be set as a Home site. Check aka.ms\u002fhomesites for cmdlet requirements.", "ErrorValue": null, "TraceCorrelationId": "f1f2e59e-3047-0000-3dd0-1f48be47bbc2", "ErrorCode": -2146232832, "ErrorTypeName": "Microsoft.SharePoint.SPException"
              }, "TraceCorrelationId": "f1f2e59e-3047-0000-3dd0-1f48be47bbc2"
            }
          ]
        ));
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/Work"
      }
    } as any), new CommandError(`The provided site url can't be set as a Home site. Check aka.ms\u002fhomesites for cmdlet requirements.`));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/Work"
      }
    } as any), new CommandError(`An error has occurred`));
  });

  it('fails validation if the siteUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the siteUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
