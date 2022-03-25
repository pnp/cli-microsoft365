import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./site-recyclebinitem-list');

describe(commands.SITE_RECYCLEBINITEM_LIST, () => {

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_RECYCLEBINITEM_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'Title', 'DirName']);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { siteUrl: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com' } });
    assert(actual);
  });

  it('fails validation if type is not an allowed value', () => {
    const actual = command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', type: 'something' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if type is an allowed value', () => {
    const actual = command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', type: 'listItems' } });
    assert(actual);
  });

  it('command handles retrieving all items from recycle bin reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/site/RecycleBin?$filter=(ItemState eq 1)') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com'
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all items from recycle bin', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/site/RecycleBin?$filter=(ItemState eq 1)') > -1) {
        return Promise.resolve(
          {
            "value": [{
              "AuthorEmail": "test.onmicrosoft.com",
              "AuthorName": "test test",
              "DeletedByEmail": "test.onmicrosoft.com",
              "DeletedByName": "test test",
              "DeletedDate": "2021-11-20T20:48:16Z",
              "DeletedDateLocalFormatted": "11/20/2021 12:48 PM",
              "DirName": "sites/test/Shared Documents",
              "DirNamePath": {
                "DecodedUrl": "sites/test/Shared Documents"
              },
              "Id": "ae6f97a7-280e-48d6-b481-0ea986c323da",
              "ItemState": 1,
              "ItemType": 1,
              "LeafName": "Document.docx",
              "LeafNamePath": {
                "DecodedUrl": "Document.docx"
              },
              "Size": "41939",
              "Title": "Document.docx"
            }]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          [{
            "AuthorEmail": "test.onmicrosoft.com",
            "AuthorName": "test test",
            "DeletedByEmail": "test.onmicrosoft.com",
            "DeletedByName": "test test",
            "DeletedDate": "2021-11-20T20:48:16Z",
            "DeletedDateLocalFormatted": "11/20/2021 12:48 PM",
            "DirName": "sites/test/Shared Documents",
            "DirNamePath": {
              "DecodedUrl": "sites/test/Shared Documents"
            },
            "Id": "ae6f97a7-280e-48d6-b481-0ea986c323da",
            "ItemState": 1,
            "ItemType": 1,
            "LeafName": "Document.docx",
            "LeafNamePath": {
              "DecodedUrl": "Document.docx"
            },
            "Size": "41939",
            "Title": "Document.docx"
          }]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all items from secondary recycle bin', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/site/RecycleBin?$filter=(ItemState eq 2)') > -1) {
        return Promise.resolve(
          {
            "value": [{
              "AuthorEmail": "test.onmicrosoft.com",
              "AuthorName": "test test",
              "DeletedByEmail": "test.onmicrosoft.com",
              "DeletedByName": "test test",
              "DeletedDate": "2021-11-20T20:48:16Z",
              "DeletedDateLocalFormatted": "11/20/2021 12:48 PM",
              "DirName": "sites/test/Shared Documents",
              "DirNamePath": {
                "DecodedUrl": "sites/test/Shared Documents"
              },
              "Id": "ae6f97a7-280e-48d6-b481-0ea986c323da",
              "ItemState": 2,
              "ItemType": 1,
              "LeafName": "Document.docx",
              "LeafNamePath": {
                "DecodedUrl": "Document.docx"
              },
              "Size": "41939",
              "Title": "Document.docx"
            }]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        output: 'json',
        secondary: true,
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          [{
            "AuthorEmail": "test.onmicrosoft.com",
            "AuthorName": "test test",
            "DeletedByEmail": "test.onmicrosoft.com",
            "DeletedByName": "test test",
            "DeletedDate": "2021-11-20T20:48:16Z",
            "DeletedDateLocalFormatted": "11/20/2021 12:48 PM",
            "DirName": "sites/test/Shared Documents",
            "DirNamePath": {
              "DecodedUrl": "sites/test/Shared Documents"
            },
            "Id": "ae6f97a7-280e-48d6-b481-0ea986c323da",
            "ItemState": 2,
            "ItemType": 1,
            "LeafName": "Document.docx",
            "LeafNamePath": {
              "DecodedUrl": "Document.docx"
            },
            "Size": "41939",
            "Title": "Document.docx"
          }]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all items from recycle bin filtered by type', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/site/RecycleBin?$filter=(ItemState eq 1) and (ItemType eq 1)') > -1) {
        return Promise.resolve(
          {
            "value": [{
              "AuthorEmail": "test.onmicrosoft.com",
              "AuthorName": "test test",
              "DeletedByEmail": "test.onmicrosoft.com",
              "DeletedByName": "test test",
              "DeletedDate": "2021-11-20T20:48:16Z",
              "DeletedDateLocalFormatted": "11/20/2021 12:48 PM",
              "DirName": "sites/test/Shared Documents",
              "DirNamePath": {
                "DecodedUrl": "sites/test/Shared Documents"
              },
              "Id": "ae6f97a7-280e-48d6-b481-0ea986c323da",
              "ItemState": 1,
              "ItemType": 5,
              "LeafName": "Document.docx",
              "LeafNamePath": {
                "DecodedUrl": "Document.docx"
              },
              "Size": "41939",
              "Title": "Document.docx"
            }]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        output: 'json',
        type: 'files',
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          [{
            "AuthorEmail": "test.onmicrosoft.com",
            "AuthorName": "test test",
            "DeletedByEmail": "test.onmicrosoft.com",
            "DeletedByName": "test test",
            "DeletedDate": "2021-11-20T20:48:16Z",
            "DeletedDateLocalFormatted": "11/20/2021 12:48 PM",
            "DirName": "sites/test/Shared Documents",
            "DirNamePath": {
              "DecodedUrl": "sites/test/Shared Documents"
            },
            "Id": "ae6f97a7-280e-48d6-b481-0ea986c323da",
            "ItemState": 1,
            "ItemType": 5,
            "LeafName": "Document.docx",
            "LeafNamePath": {
              "DecodedUrl": "Document.docx"
            },
            "Size": "41939",
            "Title": "Document.docx"
          }]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('does not retrieve items from recycle bin filtered by type', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/site/RecycleBin?$filter=(ItemState eq 1)') > -1) {
        return Promise.resolve(
          {
            "value": []
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        output: 'json',
        type: 'something',
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});