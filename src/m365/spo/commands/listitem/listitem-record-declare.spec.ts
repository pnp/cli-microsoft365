import commands from '../../commands';
import Command from '../../../../Command';
import { CommandValidate, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./listitem-record-declare');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LISTITEM_RECORD_DECLARE, () => {
  let log: any[];
  let cmdInstance: any;
  let declareItemAsRecordFakeCalled = false;

  let postFakes = (opts: any) => {
    if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {

      // requestObjectIdentity mock
      if (opts.body.indexOf('Name="Current"') > -1) {

        if ((opts.url as string).indexOf('rejectme.sharepoint.com') > -1) {
          return Promise.reject('Failed request')
        }

        if ((opts.url as string).indexOf('returnerror.sharepoint.com') > -1) {
          return Promise.reject(JSON.stringify(
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

      if (opts.body.indexOf('Name="DeclareItemAsRecord') > -1
        || opts.body.indexOf('Name="DeclareItemAsRecordWithDeclarationDate') > -1) {

        if ((opts.url as string).indexOf('alreadydeclared') > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8713.1223", "ErrorInfo": {
                "ErrorMessage": "This item has already been declared a record.", "ErrorValue": null, "TraceCorrelationId": "9d66cc9e-e0fa-8000-1225-3a9b7ff9284d", "ErrorCode": -2146232832, "ErrorTypeName": "Microsoft.SharePoint.SPException"
              }, "TraceCorrelationId": "9d66cc9e-e0fa-8000-1225-3a9b7ff9284d"
            }
          ]));
        }

        declareItemAsRecordFakeCalled = true;
        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.8713.1221",
              "ErrorInfo": null,
              "TraceCorrelationId": "9d20cc9e-7077-8000-1225-32482bc95341"
            }
          ])
        );

      }
    }
    return Promise.reject('Invalid request');
  }

  let getFakes = (opts: any) => {
    if ((opts.url as string).indexOf('/_api/web/lists') > -1 &&
      (opts.url as string).indexOf('$select=Id') > -1) {
      cmdInstance.log('faked!');
      return Promise.resolve({
        Id: '81f0ecee-75a8-46f0-b384-c8f4f9f31d99'
      });
    }
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
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_RECORD_DECLARE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('declares a record using list title is specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: true,
      listTitle: 'Test List',
      id: 147,
      webUrl: `https://contoso.sharepoint.com/sites/project-y/`,
    };

    declareItemAsRecordFakeCalled = false;
    cmdInstance.action({ options: options }, () => {
      try {
        assert(declareItemAsRecordFakeCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('declares a record using list id is passed as an option', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      listId: '99a14fe8-781c-3ce1-a1d5-c6e6a14561da',
      id: 147,
      webUrl: `https://contoso.sharepoint.com/sites/project-y/`,
      debug: true,
    };

    declareItemAsRecordFakeCalled = false;
    cmdInstance.action({ options: options }, () => {
      try {
        assert(declareItemAsRecordFakeCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('declares a record when specifying a date in debug mode', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: true,
      listId: '99a14fe8-781c-3ce1-a1d5-c6e6a14561da',
      id: 147,
      date: '2019-03-14',
      webUrl: `https://contoso.sharepoint.com/sites/project-y/`,
    };

    declareItemAsRecordFakeCalled = false;
    cmdInstance.action({ options: options }, () => {
      try {
        assert(declareItemAsRecordFakeCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('declares a record when specifying a date', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      listId: '99a14fe8-781c-3ce1-a1d5-c6e6a14561da',
      id: 147,
      date: '2019-03-14',
      webUrl: `https://contoso.sharepoint.com/sites/project-y/`,
    };

    declareItemAsRecordFakeCalled = false;
    cmdInstance.action({ options: options }, () => {
      try {
        assert(declareItemAsRecordFakeCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('it reports an error correctly when an item is already declared', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: true,
      listId: '99a14fe8-781c-3ce1-a1d5-c6e6a14561da',
      id: 147,
      date: '2019-03-14',
      webUrl: `https://alreadydeclared.sharepoint.com/sites/project-y/`,
    };

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'This item has already been declared a record.');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get _ObjecttIdentity_ when an error is returned by the _ObjectIdentity_ CSOM request', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = {
      debug: true,
      listId: '99a14fe8-781c-3ce1-a1d5-c6e6a14561da',
      id: 147,
      date: '2019-03-14',
      webUrl: `https://returnerror.sharepoint.com/sites/project-y/`
    }

    declareItemAsRecordFakeCalled = false;
    cmdInstance.action({ options: options }, () => {
      try {
        assert.notStrictEqual(declareItemAsRecordFakeCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('fails to declare a list item as a record when an error is returned', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = {
      debug: true,
      listTitle: 'Test List',
      id: 147,
      webUrl: 'https://rejectme.sharepoint.com/sites/project-y',
    }

    declareItemAsRecordFakeCalled = false;
    cmdInstance.action({ options: options }, () => {
      try {
        assert.notStrictEqual(declareItemAsRecordFakeCalled, true);
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

  it('fails validation if listTitle and listId option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle and listId are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Test List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', listTitle: 'Test List', id: '1' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the item ID is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Test List', id: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the item ID is not a positive number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Test List', id: '-1' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and numerical ID specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Test List', id: '1' } });
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo', id: '1' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: '1', debug: true } });
    assert(actual);
  });

  it('fails validation if the date passed in is not in ISO format', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: '1', date: 'foo', debug: true } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the date passed in is in ISO format', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: '1', date: 'foo', debug: true } });
    assert(actual);
  });
});