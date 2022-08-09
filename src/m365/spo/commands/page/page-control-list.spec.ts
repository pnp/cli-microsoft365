import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
import { ClientSidePage } from './clientsidepages';
import { mockControlListDataOutput, mockControlListData, mockControlListDataWithUnknownType, mockControlListDataWithUnknownTypeOutput, mockControlListDataWithText, mockControlListDataWithTextOutput } from './page-control-list.mock';
const command: Command = require('./page-control-list');

describe(commands.PAGE_CONTROL_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
      request.get,
      ClientSidePage.fromHtml
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
    assert.strictEqual(command.name.startsWith(commands.PAGE_CONTROL_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'type', 'title']);
  });

  it('lists controls on the modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {      
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlListData);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(mockControlListDataOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists controls on the modern page (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlListData);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(mockControlListDataOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists controls on the modern page when the specified page name doesn\'t contain extension', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlListData);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(mockControlListDataOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles empty columns and unknown control types', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlListDataWithUnknownType);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(mockControlListDataWithUnknownTypeOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles text web part correctly', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlListDataWithText);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(mockControlListDataWithTextOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles page when CanvasContent1 is not defined', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve({ CanvasContent1: null });
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any, () => {
      try {
        assert([]);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles page not found', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({ error: {
        "odata.error": {
          "code": "-2130575338, Microsoft.SharePoint.SPException",
          "message": {
            "lang": "en-US",
            "value": "The file /sites/team-a/SitePages/home1.aspx does not exist."
          }
        }
      } });
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The file /sites/team-a/SitePages/home1.aspx does not exist.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when retrieving pages', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any, (err?: any) => {
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
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', name: 'home.aspx' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL and name is specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});