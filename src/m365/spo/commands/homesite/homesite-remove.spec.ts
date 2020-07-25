import commands from '../../commands';
import Command, { CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
const command: Command = require('./homesite-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import appInsights from '../../../../appInsights';
import config from '../../../../config';
import * as chalk from 'chalk';

describe(commands.HOMESITE_REMOVE, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let promptOptions: any;

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
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    promptOptions = undefined;
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth,
      request.post,
      (command as any).getRequestDigest
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HOMESITE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the Home Site when confirm option is not passed', (done) => {
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {

      try {
        let promptIssued = false;

        if (promptOptions && promptOptions.type === 'confirm') {
          promptIssued = true;
        }

        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing Home Site when confirm option is not passed and prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'post');

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };

    cmdInstance.action({ options: {} }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the Home Site when prompt confirmed', (done) => {
    let homeSiteRemoveCallIssued = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><Method Name="RemoveSPHSite" Id="29" ObjectPathId="27" /></Actions><ObjectPaths><Constructor Id="27" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {

        homeSiteRemoveCallIssued = true;

        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": null, "TraceCorrelationId": "e4f2e59e-c0a9-0000-3dd0-1d8ef12cc742"
            }, 57, {
              "IsNull": false
            }, 58, "The Home site has been removed."
          ]
        ));
      }

      return Promise.reject('Invalid request');
    })

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(homeSiteRemoveCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the Home Site when prompt confirmed (debug)', (done) => {
    let homeSiteRemoveCallIssued = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><Method Name="RemoveSPHSite" Id="29" ObjectPathId="27" /></Actions><ObjectPaths><Constructor Id="27" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {

        homeSiteRemoveCallIssued = true;

        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": null, "TraceCorrelationId": "e4f2e59e-c0a9-0000-3dd0-1d8ef12cc742"
            }, 57, {
              "IsNull": false
            }, 58, "The Home site has been removed."
          ]
        ));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(homeSiteRemoveCallIssued && cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when removing the Home Site (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><Method Name="RemoveSPHSite" Id="29" ObjectPathId="27" /></Actions><ObjectPaths><Constructor Id="27" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": {
                "ErrorMessage": "The requested operation is part of an experimental feature that is not supported in the current environment.", "ErrorValue": null, "TraceCorrelationId": "75b6e89e-f072-8000-892f-75866252852a", "ErrorCode": -2146232832, "ErrorTypeName": "Microsoft.SharePoint.SPExperimentalFeatureException"
              }, "TraceCorrelationId": "f1f2e59e-3047-0000-3dd0-1f48be47bbc2"
            }
          ]
        ));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, confirm: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The requested operation is part of an experimental feature that is not supported in the current environment.`)));
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
        confirm: true
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
