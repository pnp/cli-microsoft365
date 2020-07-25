import commands from '../../commands';
import Command, { CommandValidate, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./listitem-record-undeclare');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LISTITEM_RECORD_UNDECLARE, () => {
  let log: any[];
  let cmdInstance: any;
  const expectedId = 147;
  let actualId = 0;
  let postFakes = (opts: any) => {
    if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {

      // requestObjectIdentity mock
      if (opts.body.indexOf('Name="Current"') > -1) {

        if ((opts.url as string).indexOf('rejectme.com') > -1) {

          return Promise.reject('Failed request')

        }

        if ((opts.url as string).indexOf('returnerror.com') > -1) {

          return Promise.resolve(JSON.stringify(
            [{ "ErrorInfo": "error occurred" }]
          ))

        }

        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7618.1204",
              "ErrorInfo": null,
              "TraceCorrelationId": "3e3e629e-30cc-5000-9f31-cf83b8e70021"
            },
            {
              "_ObjectType_": "SP.Web",
              "_ObjectIdentity_": "d704ae73-d5ed-459e-80b0-b8103c5fb6e0|8f2be65d-f195-4699-b0de-24aca3384ba9:site:0ead8b78-89e5-427f-b1bc-6e5a77ac191c:web:4c076c07-e3f1-49a8-ad01-dbb70b263cd7",
              "ServerRelativeUrl": "\\u002fsites\\u002fprojectx"
            }
          ])
        )

      }
      if (opts.body.indexOf('Name="UndeclareItemAsRecord') > -1) {

        actualId = expectedId;
        return Promise.resolve();
      }
    }
    return Promise.reject('Invalid request');
  }

  let getFakes = (opts: any) => {
    if ((opts.url as string).indexOf('/id') > -1) {
      return Promise.resolve({ value: "f64041f2-9818-4b67-92ff-3bc5dbbef27e" });
    }
    return Promise.reject('Invalid request');
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc'
    }));
    auth.service.connected = true;
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
      request.post,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_RECORD_UNDECLARE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails to get _ObjecttIdentity_ when an error is returned by the _ObjectIdentity_ CSOM request', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    actualId = 0;
    let options: any = {
      debug: false,
      listTitle: 'Demo List',
      id: 147,
      webUrl: 'https://returnerror.com/sites/project-y',
    }

    cmdInstance.action({ options: options }, () => {
      try {
        done();
      }
      catch (e) {
        assert(actualId !== expectedId);
        done(e);
      }
    });

  });
  it('correctly undeclares list item as a record when listTitle is passes', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    command.allowUnknownOptions();

    let options: any = {
      debug: true,
      listTitle: 'Demo List',
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x'
    }
    cmdInstance.action({ options: options }, () => {
      try {
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });
  it('correctly undeclares list item as a record when listId is passed', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    command.allowUnknownOptions();

    let options: any = {
      debug: true,
      listId: '770fe148-1d72-480e-8cde-f9d3832798b6',
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x'
    }
    cmdInstance.action({ options: options }, () => {
      try {
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });
  it('fails to undeclare a list item as a record when \'reject me\' values are used', (done) => {
    actualId = 0;

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = {
      debug: false,
      listTitle: 'Demo List',
      id: 47,
      webUrl: 'https://rejectme.com/sites/project-y',
    }

    cmdInstance.action({ options: options }, () => {
      try {
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

  it('supports specifying URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if both id and title options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1 } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', id: 1 } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: 1 } });
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', id: 1 } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: 1 } });
    assert(actual);
  });

  it('fails validation if both id and title options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents', id: 1 } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'abc' } });
    assert.notStrictEqual(actual, true);
  });
});
