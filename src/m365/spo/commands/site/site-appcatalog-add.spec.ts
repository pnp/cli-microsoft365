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
const command: Command = require('./site-appcatalog-add');

describe(commands.SITE_APPCATALOG_ADD, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_APPCATALOG_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates site collection app catalog (debug=false)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="38" ObjectPathId="37" /><ObjectPath Id="40" ObjectPathId="39" /><ObjectPath Id="42" ObjectPathId="41" /><ObjectPath Id="44" ObjectPathId="43" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Constructor Id="37" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="39" ParentId="37" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method><Property Id="41" ParentId="39" Name="RootWeb" /><Property Id="43" ParentId="41" Name="TenantAppCatalog" /><Property Id="45" ParentId="43" Name="SiteCollectionAppCatalogsSites" /><Method Id="47" ParentId="45" Name="Add"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
            }, 38, {
              "IsNull": false
            }, 40, {
              "IsNull": false
            }, 42, {
              "IsNull": false
            }, 44, {
              "IsNull": false
            }, 46, {
              "IsNull": false
            }, 48, {
              "IsNull": false
            }
          ]
          ));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/site' } }, () => {
      try {
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates site collection app catalog (debug=true)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="38" ObjectPathId="37" /><ObjectPath Id="40" ObjectPathId="39" /><ObjectPath Id="42" ObjectPathId="41" /><ObjectPath Id="44" ObjectPathId="43" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Constructor Id="37" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="39" ParentId="37" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method><Property Id="41" ParentId="39" Name="RootWeb" /><Property Id="43" ParentId="41" Name="TenantAppCatalog" /><Property Id="45" ParentId="43" Name="SiteCollectionAppCatalogsSites" /><Method Id="47" ParentId="45" Name="Add"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
            }, 38, {
              "IsNull": false
            }, 40, {
              "IsNull": false
            }, 42, {
              "IsNull": false
            }, 44, {
              "IsNull": false
            }, 46, {
              "IsNull": false
            }, 48, {
              "IsNull": false
            }
          ]
          ));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/site' } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when creating site collection app catalog', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="38" ObjectPathId="37" /><ObjectPath Id="40" ObjectPathId="39" /><ObjectPath Id="42" ObjectPathId="41" /><ObjectPath Id="44" ObjectPathId="43" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Constructor Id="37" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="39" ParentId="37" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method><Property Id="41" ParentId="39" Name="RootWeb" /><Property Id="43" ParentId="41" Name="TenantAppCatalog" /><Property Id="45" ParentId="43" Name="SiteCollectionAppCatalogsSites" /><Method Id="47" ParentId="45" Name="Add"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": {
                "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "d2d0389e-a040-4000-b24b-d16b0546a03c", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
              }, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
            }
          ]
          ));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/site' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Unknown Error")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/site' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying site url', () => {
    const options = command.options;
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--url') > -1) {
        assert(true);
        return;
      }
    }
    assert(false);
  });

  it('passes validation when the url option specified', async () => {
    const actual = await command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/site'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the url option not specified', async () => {
    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.notDeepEqual(actual, true);
  });

  it('fails validation when the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options: {
        url: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});