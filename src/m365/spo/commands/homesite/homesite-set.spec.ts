import commands from '../../commands';
import Command, { CommandValidate, CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
const command: Command = require('./homesite-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import appInsights from '../../../../appInsights';
import config from '../../../../config';
import * as chalk from 'chalk';

describe(commands.HOMESITE_SET, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => {
      return {
        FormDigestValue: 'ABC'
      };
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      (command as any).getRequestDigest
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

  it('sets the specified site as the Home Site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="57" ObjectPathId="56" /><Method Name="SetSPHSite" Id="58" ObjectPathId="56"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Work</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="56" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

    cmdInstance.action({
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/Work"
      }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith('The Home site has been set to https://contoso.sharepoint.com/sites/Work.'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets the specified site as the Home Site (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="57" ObjectPathId="56" /><Method Name="SetSPHSite" Id="58" ObjectPathId="56"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Work</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="56" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

    cmdInstance.action({
      options: {
        debug: true,
        siteUrl: "https://contoso.sharepoint.com/sites/Work"
      }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when setting the Home Site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="57" ObjectPathId="56" /><Method Name="SetSPHSite" Id="58" ObjectPathId="56"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Work</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="56" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

    cmdInstance.action({
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/Work"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The provided site url can't be set as a Home site. Check aka.ms\u002fhomesites for cmdlet requirements.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => Promise.reject('An error has occurred'));

    cmdInstance.action({
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/Work"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the siteUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the siteUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});