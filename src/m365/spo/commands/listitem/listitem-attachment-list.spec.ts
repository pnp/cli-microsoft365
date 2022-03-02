import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./listitem-attachment-list');

describe(commands.LISTITEM_ATTACHMENT_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  const itemId = 147;

  const getFakes = (opts: any) => {
    if ((opts.url as string).indexOf('/_api/web/lists') > -1) {
      return Promise.resolve(
        {
          "AttachmentFiles": [
            {
              "FileName": "my_file.docx",
              "ServerRelativeUrl": "/sites/project-x/Lists/Demo List/Attachments/1/my_file.docx"
            },
            {
              "FileName": "my_workbook.xlsx",
              "ServerRelativeUrl": "/sites/project-x/Lists/Demo List/Attachments/1/my_workbook.xlsx"
            }
          ]
        }
      );
    }
    return Promise.reject('Invalid request');
  };

  const jsonOutput = [
    {
      "FileName": "my_file.docx",
      "ServerRelativeUrl": "/sites/project-x/Lists/Demo List/Attachments/1/my_file.docx"
    },
    {
      "FileName": "my_workbook.xlsx",
      "ServerRelativeUrl": "/sites/project-x/Lists/Demo List/Attachments/1/my_workbook.xlsx"
    }
  ];

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
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
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_ATTACHMENT_LIST), true);
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

  it('supports specifying URL', () => {
    const options = command.options();
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if listTitle and listId option not specified', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', itemId: itemId } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle and listId are specified together', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', itemId: itemId } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo', listTitle: 'Demo List', itemId: itemId } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', itemId: itemId } });
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo', itemId: itemId } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', itemId: itemId } });
    assert(actual);
  });

  it('fails validation if the specified itemId is not a number', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', itemId: 'a' } });
    assert.notStrictEqual(actual, true);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['FileName', 'ServerRelativeUrl']);
  });

  it('returns attachments associated to a list item by listId', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    const options: any = {
      debug: true,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
      itemId: itemId
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns attachments associated to a list item by listTitle', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    const options: any = {
      debug: true,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      listTitle: 'Demo List',
      itemId: itemId
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    const options: any = {
      debug: false,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
      itemId: itemId,
      output: "json"
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles No attachments found (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve({ AttachmentFiles: [] });
    });

    const options: any = {
      debug: true,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
      itemId: itemId
    };

    command.action(logger, { options: options }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('No attachments found'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});
