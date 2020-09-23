import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./orgassetslibrary-add');

describe(commands.ORGASSETSLIBRARY_ADD, () => {
  let log: any[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;

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
    logger = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    loggerSpy = sinon.spy(logger, 'log');
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
    assert.strictEqual(command.name.startsWith(commands.ORGASSETSLIBRARY_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds a new library as org assets library (debug)', (done) => {
    let orgAssetLibAddCallIssued = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddToOrgAssetsLibAndCdnWithType" Id="11" ObjectPathId="8"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets</Parameter><Parameter Type="Null" /><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        orgAssetLibAddCallIssued = true;

        return Promise.resolve(JSON.stringify(
          [{
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19708.12061", "ErrorInfo": null, "TraceCorrelationId": "a0a8309f-4039-a000-ea81-9b8297eb43e0"
          }]
        ));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, libraryUrl: 'https://contoso.sharepoint.com/siteassets' } }, () => {
      try {
        assert(orgAssetLibAddCallIssued && loggerSpy.calledWith(chalk.green('DONE')));

        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('adds a new library as org assets library with CDN Type (debug)', (done) => {
    let orgAssetLibAddCallIssued = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddToOrgAssetsLibAndCdnWithType" Id="11" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets</Parameter><Parameter Type="Null" /><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        orgAssetLibAddCallIssued = true;

        return Promise.resolve(JSON.stringify(
          [{
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19708.12061", "ErrorInfo": null, "TraceCorrelationId": "a0a8309f-4039-a000-ea81-9b8297eb43e0"
          }]
        ));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, libraryUrl: 'https://contoso.sharepoint.com/siteassets', cdnType: 'Public' } }, () => {
      try {
        assert(orgAssetLibAddCallIssued && loggerSpy.calledWith(chalk.green('DONE')));

        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('adds a new library as org assets library with CDN Type and thumbnailUrl (debug)', (done) => {
    let orgAssetLibAddCallIssued = false;

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddToOrgAssetsLibAndCdnWithType" Id="11" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets/logo.png</Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        orgAssetLibAddCallIssued = true;

        return Promise.resolve(JSON.stringify(
          [{
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19708.12061", "ErrorInfo": null, "TraceCorrelationId": "a0a8309f-4039-a000-ea81-9b8297eb43e0"
          }]
        ));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, libraryUrl: 'https://contoso.sharepoint.com/siteassets', cdnType: 'Public', thumbnailUrl: 'https://contoso.sharepoint.com/siteassets/logo.png' } }, () => {
      try {
        assert(orgAssetLibAddCallIssued && loggerSpy.calledWith(chalk.green('DONE')));

        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('handles error if is already present', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddToOrgAssetsLibAndCdnWithType" Id="11" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets</Parameter><Parameter Type="String">https://contoso.sharepoint.com/siteassets/logo.png</Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

    command.action(logger, { options: { debug: true, libraryUrl: 'https://contoso.sharepoint.com/siteassets', cdnType: 'Public', thumbnailUrl: 'https://contoso.sharepoint.com/siteassets/logo.png' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`This library is already an organization assets library.`)));

        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('handles error getting request', (done) => {
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

    command.action(logger, {
      options: {
        debug: true
      }
    } as any, (err?: any) => {
      try {
        assert(svcListRequest.called);
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => Promise.reject('An error has occurred'));

    command.action(logger, { options: {} } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the libraryUrl is not valid', () => {
    const actual = command.validate({ options: { libraryUrl: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the thumbnail is not valid', () => {
    const actual = command.validate({ options: { libraryUrl: 'https://contoso.sharepoint.com/siteassets', thumbnailUrl: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the libraryUrl option is a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { libraryUrl: 'https://contoso.sharepoint.com/siteassets' } });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});
