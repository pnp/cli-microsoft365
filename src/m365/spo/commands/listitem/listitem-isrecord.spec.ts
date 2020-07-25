import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./listitem-isrecord');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LISTITEM_ISRECORD, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  let postFakes = (opts: any) => {
    // requestObjectIdentity mock
    if (opts.body.indexOf('Name="Current"') > -1) {
      if ((opts.url as string).indexOf('returnerror.sharepoint.com') > -1) {
        cmdInstance.log("Returns error from requestObjectIdentity")
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

    // IsRecord request mocks
    if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
      // Unsuccessful response for when the item does not exist
      if ((opts.url as string).indexOf('itemdoesnotexist.sharepoint.com') > -1) {
        return Promise.resolve(JSON.stringify(
          [
            {
              "ErrorInfo": { "ErrorMessage": "Item does not exist. It may have been deleted by another user.", "ErrorValue": null, "TraceCorrelationId": "fedae69e-4077-8000-f13a-d4a607aefc32", "ErrorCode": -2130575338, "ErrorTypeName": "Microsoft.SharePoint.SPException" },
              "LibraryVersion": "16.0.9005.1214",
              "SchemaVersion": "15.0.0.0",
              "TraceCorrelationId": "fedae69e-4077-8000-f13a-d4a607aefc32",
            }]));
      }

      // Successful response
      return Promise.resolve(JSON.stringify(
        [
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.9005.1214", "ErrorInfo": null, "TraceCorrelationId": "9ec8e69e-d001-8000-f13a-d5e03849cd96"
          }, 32, true
        ]
      ));
    }
    return Promise.reject('Invalid request');
  }

  let getFakes = (opts: any) => {
    // Get list mock
    if ((opts.url as string).indexOf('/_api/web/lists') > -1 &&
      (opts.url as string).indexOf('$select=Id') > -1) {
      cmdInstance.log('faked!');
      return Promise.resolve({
        Id: '81f0ecee-75a8-46f0-b384-c8f4f9f31d99'
      });
    }
    if ((opts.url as string).indexOf('?select=Id') > -1) {
      return Promise.resolve({ value: "f64041f2-9818-4b67-92ff-3bc5dbbef27e" });
    }
    return Promise.reject('Invalid request');
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub((command as any), 'getRequestDigest').callsFake(() => Promise.resolve({
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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth,
      (command as any).getRequestDigest
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_ISRECORD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('throws an error when requesting a record for an item that does not exist', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: true,
      listTitle: 'Test List',
      id: 147,
      webUrl: `https://itemdoesnotexist.sharepoint.com/sites/project-y/`,
      verbose: true
    };

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Item does not exist. It may have been deleted by another user.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('test a record with list title passed in as an option', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: true,
      listTitle: 'Test List',
      id: 147,
      webUrl: `https://contoso.sharepoint.com/sites/project-y/`,
      verbose: true
    };

    cmdInstance.action({ options: options }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith("Getting list id..."));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('test a record with list id passed in as an option', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      listId: '99a14fe8-781c-3ce1-a1d5-c6e6a14561da',
      id: 147,
      webUrl: `https://contoso.sharepoint.com/sites/project-y/`,
      debug: true,
      verbose: true
    };

    cmdInstance.action({ options: options }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith("List Id passed in as an argument."));
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

    cmdInstance.action({ options: options }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith("Returns error from requestObjectIdentity"));
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
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '1' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle and listId are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '1', listTitle: 'Test List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
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
});
