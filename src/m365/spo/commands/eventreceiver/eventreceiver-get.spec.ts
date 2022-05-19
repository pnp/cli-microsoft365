import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./eventreceiver-get');

describe(commands.EVENTRECEIVER_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const eventReceiverResponseJson = [
    {
      "ReceiverAssembly": "",
      "ReceiverClass": "",
      "ReceiverId": "c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec",
      "ReceiverName": "PnP Test Receiver",
      "SequenceNumber": 30846,
      "Synchronization": 1,
      "EventType": 1,
      "ReceiverUrl": "https://pnp.github.io"
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
    assert.strictEqual(command.name.startsWith(commands.EVENTRECEIVER_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets();
    assert.deepStrictEqual(optionSets, [['name', 'id']]);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', () => {
    const actual = command.validate({ options: { webUrl: 'site.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is set to site and one of the list properties is filled in', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', scope: 'site', listTitle: 'Documents'} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the list ID is not a valid GUID', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listId: 'abc'} });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with id', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with name', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list id', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list title', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list url', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listUrl: 'sites/hr-life/Lists/breakInheritance' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if title and id are specified together', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and id and url are specified together', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', listUrl: 'sites/hr-life/Lists/breakInheritance' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and url are specified together', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List', listUrl: 'sites/hr-life/Lists/breakInheritance' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is invalid value', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('correctly handles list not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
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

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', name: 'PnP Test Receiver' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves web event receiver using name as option', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'PnP Test Receiver' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves site event receiver using name as option', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/site/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site', name: 'PnP Test Receiver' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list event receiver retrieved by list title using name as option', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', name: 'PnP Test Receiver' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list event receivers queried by url using name as option', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fportal%2FShared%20Documents')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listUrl: 'Shared Documents', name: 'PnP Test Receiver' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list event receivers queried by list id using name as option', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'b17bd74f-d1b1-42bf-a21d-f865a903acc3')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listId: 'b17bd74f-d1b1-42bf-a21d-f865a903acc3', name: 'PnP Test Receiver' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves web event receiver using id as option', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/eventreceivers?$filter=receiverid eq (guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves site event receiver using id as option', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/site/eventreceivers?$filter=receiverid eq (guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list event receiver retrieved by list title using id as option', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers?$filter=receiverid eq (guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list event receivers queried by url using id as option', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fportal%2FShared%20Documents')/eventreceivers?$filter=receiverid eq (guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listUrl: 'Shared Documents', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves list event receivers queried by list id using id as option', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'b17bd74f-d1b1-42bf-a21d-f865a903acc3')/eventreceivers?$filter=receiverid eq (guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return Promise.resolve(eventReceiverValue);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listId: 'b17bd74f-d1b1-42bf-a21d-f865a903acc3', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' } }, () => {
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