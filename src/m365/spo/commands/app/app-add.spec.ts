import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./app-add');

describe(commands.APP_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    requests = [];
    sinon.stub(request, 'get').resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      fs.readFileSync
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APP_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds new app to the tenant app catalog', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.resolve('{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}');
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: false, filePath: 'spfx.sppkg' } }, () => {
      try {
        assert(loggerLogSpy.calledWith("bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new app to the tenant app catalog (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.resolve('{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0 &&
          r.headers.binaryStringRequestBody &&
          r.data) {
          correctRequestIssued = true;
        }
      });

      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('adds new app to a site app catalog (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.resolve('{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg', scope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/sitecollectionappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0 &&
          r.headers.binaryStringRequestBody &&
          r.data) {
          correctRequestIssued = true;
        }
      });

      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('returns all info about the added app in the JSON output mode', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.resolve('{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: false, filePath: 'spfx.sppkg', output: 'json' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(JSON.parse('{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('correctly handles failure when the app already exists in the tenant app catalog', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.reject({
            error: JSON.stringify({ "odata.error": { "code": "-2130575257, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "A file with the name AppCatalog/spfx.sppkg already exists. It was last modified by i:0#.f|membership|admin@contoso.onmi on 24 Nov 2017 12:50:43 -0800." } } })
          });
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('A file with the name AppCatalog/spfx.sppkg already exists. It was last modified by i:0#.f|membership|admin@contoso.onmi on 24 Nov 2017 12:50:43 -0800.')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('correctly handles failure when the app already exists in the site app catalog', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.reject({
            error: JSON.stringify({ "odata.error": { "code": "-2130575257, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "A file with the name AppCatalog/spfx.sppkg already exists. It was last modified by i:0#.f|membership|admin@contoso.onmi on 24 Nov 2017 12:50:43 -0800." } } })
          });
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg', scope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('A file with the name AppCatalog/spfx.sppkg already exists. It was last modified by i:0#.f|membership|admin@contoso.onmi on 24 Nov 2017 12:50:43 -0800.')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.reject({ error: 'An error has occurred' });
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('correctly handles random API error when sitecollection', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.reject({ error: 'An error has occurred' });
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg', scope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('correctly handles random API error (string error)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.reject('An error has occurred');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('correctly handles random API error when sitecollection (string error)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.reject('An error has occurred');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg', scope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('handles promise error while getting tenant appcatalog', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, {
      options: {
        debug: true, filePath: 'spfx.sppkg'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error while getting tenant appcatalog', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
              "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "18091989-62a6-4cad-9717-29892ee711bc", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.ServerException"
            }, "TraceCorrelationId": "18091989-62a6-4cad-9717-29892ee711bc"
          }
        ]));
      }
      if ((opts.url as string).indexOf('contextinfo') > -1) {
        return Promise.resolve('abc');
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true, filePath: 'spfx.sppkg'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('fails validation on invalid scope', async () => {
    const actual = await command.validate({ options: { scope: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation on valid \'tenant\' scope', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({ options: { scope: 'tenant', filePath: 'abc' } }, commandInfo);
    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'Tenant\' scope', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({ options: { scope: 'Tenant', filePath: 'abc' } }, commandInfo);
    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'SiteCollection\' scope', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({ options: { scope: 'SiteCollection', appCatalogUrl: 'https://contoso.sharepoint.com', filePath: 'abc' } }, commandInfo);
    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('submits to tenant app catalog when scope not specified', (done) => {
    // setup call to fake requests...
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.resolve('{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { filePath: 'spfx.sppkg' } }, () => {
      let correctAppCatalogUsed = false;
      requests.forEach(r => {
        if (r.url.indexOf('/tenantappcatalog/') > -1) {
          correctAppCatalogUsed = true;
        }
      });

      try {
        assert(correctAppCatalogUsed);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('submits to tenant app catalog when scope \'tenant\' specified ', (done) => {
    // setup call to fake requests...
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.resolve('{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { scope: 'tenant', filePath: 'spfx.sppkg' } }, () => {
      let correctAppCatalogUsed = false;
      requests.forEach(r => {
        if (r.url.indexOf('/tenantappcatalog/') > -1) {
          correctAppCatalogUsed = true;
        }
      });

      try {
        assert(correctAppCatalogUsed);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('submits to sitecollection app catalog when scope \'sitecollection\' specified ', (done) => {
    // setup call to fake requests...
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.resolve('{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { scope: 'sitecollection', filePath: 'spfx.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com' } }, () => {
      let correctAppCatalogUsed = false;
      requests.forEach(r => {
        if (r.url.indexOf('/sitecollectionappcatalog/') > -1) {
          correctAppCatalogUsed = true;
        }
      });

      try {
        assert(correctAppCatalogUsed);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post,
          fs.readFileSync
        ]);
      }
    });
  });

  it('fails validation if file path doesn\'t exist', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = await command.validate({ options: { filePath: 'abc' } }, commandInfo);
    sinonUtil.restore(fs.existsSync);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if file path points to a directory', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => true);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = await command.validate({ options: { filePath: 'abc' } }, commandInfo);
    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid scope is specified', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({ options: { filePath: 'abc', scope: 'foo' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when path points to a valid file', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({ options: { filePath: 'abc' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('passes validation when no scope is specified', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({ options: { filePath: 'abc' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'tenant\'', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({ options: { filePath: 'abc', scope: 'tenant' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });


  it('should fail when \'sitecollection\' scope, but no appCatalogUrl specified', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({ options: { filePath: 'abc', scope: 'sitecollection' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });

  it('should not fail when \'tenant\' scope, but also appCatalogUrl specified', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({ options: { filePath: 'abc', scope: 'tenant', appCatalogUrl: 'https://contoso.sharepoint.com' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but bad appCatalogUrl format specified', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({ options: { filePath: 'abc', scope: 'sitecollection', appCatalogUrl: 'contoso.sharepoint.com' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });
});