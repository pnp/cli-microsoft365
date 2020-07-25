import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import config from '../../../../config';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./cdn-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.CDN_GET, () => {
  let log: any[];
  let cmdInstance: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    auth.service.tenantId = 'abc';
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
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
    auth.service.tenantId = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CDN_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves the settings of the public CDN when type set to Public', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([{"SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.7025.1207","ErrorInfo":null,"TraceCorrelationId":"3d92299e-e019-4000-c866-de7d45aa9628"},12,true]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, type: 'Public' } }, () => {
      let correctLogStatement = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('Public CDN at') > -1 && l.indexOf('enabled') > -1) {
          correctLogStatement = true;
        }
      });
      try {
        assert(correctLogStatement);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves the settings of the private CDN when type set to Private', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([{"SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.7025.1207","ErrorInfo":null,"TraceCorrelationId":"3d92299e-e019-4000-c866-de7d45aa9628"},12,false]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, type: 'Private' } }, () => {
      let correctLogStatement = false;
      log.forEach(l => {
        if (l === false) {
          correctLogStatement = true;
        }
      });
      try {
        assert(correctLogStatement);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves the settings of the private CDN when type set to Private (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([{"SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.7025.1207","ErrorInfo":null,"TraceCorrelationId":"3d92299e-e019-4000-c866-de7d45aa9628"},12,false]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, type: 'Private' } }, () => {
      let correctLogStatement = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('disabled') > -1) {
          correctLogStatement = true;
        }
      });
      try {
        assert(correctLogStatement);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves the settings of the public CDN when no type set', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([{"SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.7025.1207","ErrorInfo":null,"TraceCorrelationId":"3d92299e-e019-4000-c866-de7d45aa9628"},12,true]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true } }, () => {
      let correctLogStatement = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('Public CDN at') > -1 && l.indexOf('enabled') > -1) {
          correctLogStatement = true;
        }
      });
      try {
        assert(correctLogStatement);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles an error when getting tenant CDN settings', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.body) {
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="abc" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([
              {
                "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
                  "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
                }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
              }
            ]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('supports specifying CDN type', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[type]') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return []; });
    const options = (command.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('accepts Public SharePoint Online CDN type', () => {
    const actual = (command.validate() as CommandValidate)({ options: { type: 'Public' } });
    assert.strictEqual(actual, true);
  });

  it('accepts Private SharePoint Online CDN type', () => {
    const actual = (command.validate() as CommandValidate)({ options: { type: 'Private' } });
    assert.strictEqual(actual, true);
  });

  it('rejects invalid SharePoint Online CDN type', () => {
    const type = 'foo';
    const actual = (command.validate() as CommandValidate)({ options: { type: type } });
    assert.strictEqual(actual, `${type} is not a valid CDN type. Allowed values are Public|Private`);
  });

  it('doesn\'t fail validation if the optional type option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.strictEqual(actual, true);
  });
});