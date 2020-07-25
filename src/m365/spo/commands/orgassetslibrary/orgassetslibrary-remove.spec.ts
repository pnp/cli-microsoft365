import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./orgassetslibrary-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import config from '../../../../config';
import * as chalk from 'chalk';

describe(commands.ORGASSETSLIBRARY_REMOVE, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc'
    }));
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
      auth.restoreAuth,
      appInsights.trackEvent,
      (command as any).getRequestDigest
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ORGASSETSLIBRARY_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the Org Assets Library when confirm option is not passed', (done) => {
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

  it('aborts removing the Org Assets Library when confirm option is not passed and prompt not confirmed', (done) => {
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

  it('removes the Org Assets Library when prompt confirmed', (done) => {
    let orgAssetLibRemoveCallIssued = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><Method Name="RemoveFromOrgAssets" Id="10" ObjectPathId="8"><Parameters><Parameter Type="String">/sites/branding/assets</Parameter><Parameter Type="Guid">{00000000-0000-0000-0000-000000000000}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {

        orgAssetLibRemoveCallIssued = true;

        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19520.12061", "ErrorInfo": null, "TraceCorrelationId": "f4e1279f-100c-9000-7ea4-40fa74757476"
            }, 9, {
              "IsNull": false
            }
          ]
        ));
      }

      return Promise.reject('Invalid request');
    })

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { libraryUrl: '/sites/branding/assets' } }, () => {
      try {
        assert(orgAssetLibRemoveCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the Org Assets Library bypassing confirmation (debug)', (done) => {
    let orgAssetLibRemoveCallIssued = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><Method Name="RemoveFromOrgAssets" Id="10" ObjectPathId="8"><Parameters><Parameter Type="String">/sites/branding/assets</Parameter><Parameter Type="Guid">{00000000-0000-0000-0000-000000000000}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {

        orgAssetLibRemoveCallIssued = true;

        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19520.12061", "ErrorInfo": null, "TraceCorrelationId": "f4e1279f-100c-9000-7ea4-40fa74757476"
            }, 9, {
              "IsNull": false
            }
          ]
        ));
      }

      return Promise.reject('Invalid request');
    })

    cmdInstance.action({ options: { libraryUrl: '/sites/branding/assets', debug: true, verbose: true, confirm: true } }, () => {
      try {
        assert(orgAssetLibRemoveCallIssued && cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the Org Assets Library when prompt confirmed', (done) => {
    let orgAssetLibRemoveCallIssued = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><Method Name="RemoveFromOrgAssets" Id="10" ObjectPathId="8"><Parameters><Parameter Type="String">/sites/branding/assets</Parameter><Parameter Type="Guid">{00000000-0000-0000-0000-000000000000}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {

        orgAssetLibRemoveCallIssued = true;

        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19520.12061", "ErrorInfo": null, "TraceCorrelationId": "f4e1279f-100c-9000-7ea4-40fa74757476"
            }, 9, {
              "IsNull": false
            }
          ]
        ));
      }

      return Promise.reject('Invalid request');
    })

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { libraryUrl: '/sites/branding/assets', output: 'json' } }, () => {
      try {
        assert(orgAssetLibRemoveCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when removing a non-existing Org Asset Library', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><Method Name="RemoveFromOrgAssets" Id="10" ObjectPathId="8"><Parameters><Parameter Type="String">/sites/branding/assets</Parameter><Parameter Type="Guid">{00000000-0000-0000-0000-000000000000}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {

        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19520.12061", "ErrorInfo": {
                "ErrorMessage": "Run Add-SPOOrgAssetsLibrary first to set up the organization assets library feature for your organization.", "ErrorValue": null, "TraceCorrelationId": "5fe2279f-40d7-9000-7e58-51033180e44d", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "5fe2279f-40d7-9000-7e58-51033180e44d"
            }
          ]
        ));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { libraryUrl: '/sites/branding/assets', debug: true, confirm: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Run Add-SPOOrgAssetsLibrary first to set up the organization assets library feature for your organization.`)));
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
