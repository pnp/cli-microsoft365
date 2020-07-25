import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./hubsite-rights-grant');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.HUBSITE_RIGHTS_GRANT, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
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
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HUBSITE_RIGHTS_GRANT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('grants rights on the specified site design to the specified principal', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('grants rights on the specified site design to the specified principal (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('grants rights on the specified site design to the specified principals', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object><Object Type="String">user</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin,user', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('grants rights on the specified site design to the specified principals (email)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales</Parameter><Parameter Type="Array"><Object Type="String">admin@contoso.onmicrosoft.com</Object><Object Type="String">user@contoso.onmicrosoft.com</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin@contoso.onmicrosoft.com,user@contoso.onmicrosoft.com', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('grants rights on the specified site design to the specified principals separated with an extra space', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object><Object Type="String">user</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin, user', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in user input', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales&gt;</Parameter><Parameter Type="Array"><Object Type="String">admin&gt;</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales>', principals: 'admin>', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": {
              "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "86be439e-80c4-5000-fcf8-b746bccdc4e7", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
            }, "TraceCorrelationId": "86be439e-80c4-5000-fcf8-b746bccdc4e7"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin', rights: 'Join' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin', rights: 'Join' } }, (err?: any) => {
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
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying hub site url', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--url') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying principals', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--principals') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying rights', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--rights') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if url is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'abc', principals: 'admin', rights: 'Join' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified rights value is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/sales', principals: 'PattiF', rights: 'Invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required parameters are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/sales', principals: 'PattiF', rights: 'Join' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if all required parameters are valid (multiple principals)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/sales', principals: 'PattiF,AdeleV', rights: 'Join' } });
    assert.strictEqual(actual, true);
  });
});