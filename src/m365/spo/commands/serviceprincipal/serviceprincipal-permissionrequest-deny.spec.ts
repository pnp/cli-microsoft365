import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./serviceprincipal-permissionrequest-deny');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_DENY, () => {
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
    assert.strictEqual(command.name.startsWith(commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_DENY), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('denies the specified permission request (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="160" ObjectPathId="159" /><ObjectPath Id="162" ObjectPathId="161" /><ObjectPath Id="164" ObjectPathId="163" /><Method Name="Deny" Id="165" ObjectPathId="163" /></Actions><ObjectPaths><Constructor Id="159" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="161" ParentId="159" Name="PermissionRequests" /><Method Id="163" ParentId="161" Name="GetById"><Parameters><Parameter Type="Guid">{4dc4c043-25ee-40f2-81d3-b3bf63da7538}</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "1c643a9e-40b1-4000-c0ac-2fae75aa36ca"
          }, 211, {
            "IsNull": false
          }, 213, {
            "IsNull": false
          }, 215, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, requestId: '4dc4c043-25ee-40f2-81d3-b3bf63da7538' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('denies the specified permission request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="160" ObjectPathId="159" /><ObjectPath Id="162" ObjectPathId="161" /><ObjectPath Id="164" ObjectPathId="163" /><Method Name="Deny" Id="165" ObjectPathId="163" /></Actions><ObjectPaths><Constructor Id="159" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="161" ParentId="159" Name="PermissionRequests" /><Method Id="163" ParentId="161" Name="GetById"><Parameters><Parameter Type="Guid">{4dc4c043-25ee-40f2-81d3-b3bf63da7538}</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "1c643a9e-40b1-4000-c0ac-2fae75aa36ca"
          }, 211, {
            "IsNull": false
          }, 213, {
            "IsNull": false
          }, 215, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, requestId: '4dc4c043-25ee-40f2-81d3-b3bf63da7538' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when denying permission request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "A permission request with the ID f0feaecf-24be-402b-a080-3a55738ec56a could not be found.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "Microsoft.SharePoint.Client.ResourceNotFoundException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, requestId: 'f0feaecf-24be-402b-a080-3a55738ec56a' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('A permission request with the ID f0feaecf-24be-402b-a080-3a55738ec56a could not be found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => Promise.reject('An error has occurred'));
    cmdInstance.action({ options: { debug: false, requestId: 'f0feaecf-24be-402b-a080-3a55738ec56a' } }, (err?: any) => {
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

  it('allows specifying requestId', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--requestId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the requestId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { requestId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the requestId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { requestId: '4dc4c043-25ee-40f2-81d3-b3bf63da7538' } });
    assert.strictEqual(actual, true);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });
});