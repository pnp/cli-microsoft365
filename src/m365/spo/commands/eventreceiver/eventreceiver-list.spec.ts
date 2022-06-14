import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./eventreceiver-list');

describe(commands.EVENTRECEIVER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const eventReceiverResponseJson = [
    {
      "ReceiverAssembly": "Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
      "ReceiverClass": "Microsoft.SharePoint.Internal.SitePages.Sharing.PageSharingEventReceiver",
      "ReceiverId": "625b1f4c-2869-457f-8b41-bed72059bb2b",
      "ReceiverName": "Microsoft.SharePoint.Internal.SitePages.Sharing.PageSharingEventReceiver",
      "SequenceNumber": 10000,
      "Synchronization": 1,
      "EventType": 309,
      "ReceiverUrl": null
    },
    {
      "ReceiverAssembly": "Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
      "ReceiverClass": "Microsoft.SharePoint.Internal.SitePages.Sharing.PageSharingEventReceiver",
      "ReceiverId": "41ad359e-ac6a-4a5e-8966-a85492ca4f52",
      "ReceiverName": "Microsoft.SharePoint.Internal.SitePages.Sharing.PageSharingEventReceiver",
      "SequenceNumber": 10000,
      "Synchronization": 1,
      "EventType": 310,
      "ReceiverUrl": null
    }
  ];

  const eventReceiverValue = {
    value: eventReceiverResponseJson
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.EVENTRECEIVER_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['ReceiverId', 'ReceiverName']);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', () => {
    const actual = command.validate({ options: { webUrl: 'site.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is set to site and one of the list properties is filled in', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', scope: 'site', listTitle: 'Documents' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the list ID is not a valid GUID', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list id', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list title', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list url', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listUrl: 'sites/hr-life/Lists/breakInheritance' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if title and id are specified together', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and id and url are specified together', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', listUrl: 'sites/hr-life/Lists/breakInheritance' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and url are specified together', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listUrl: 'sites/hr-life/Lists/breakInheritance' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is invalid value', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('correctly handles list not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers`) > -1) {
        return Promise.reject({
          error: {
            "odata.error": {
              "code": "-1, System.ArgumentException",
              "message": {
                "lang": "en-US",
                "value": "List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'."
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all web event receivers', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/eventreceivers`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all site event receivers', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/site/eventreceivers`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all list event receivers queried by title', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all list event receivers queried by url', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fportal%2FShared%20Documents')/eventreceivers`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listUrl: 'Shared Documents' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all list event receivers queried by guid', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'b17bd74f-d1b1-42bf-a21d-f865a903acc3')/eventreceivers`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listId: 'b17bd74f-d1b1-42bf-a21d-f865a903acc3' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

});