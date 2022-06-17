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
import { mockControlGetData, mockControlGetDataEmptyColumn, mockControlGetDataEmptyColumnOutput, mockControlGetDataOutput, mockControlGetDataWithoutAnId, mockControlGetDataWithText, mockControlGetDataWithTextOutput, mockControlGetDataWithUnknownType, mockControlGetDataWithUnknownTypeOutput } from './page-control-get.mock';
const command: Command = require('./page-control-get');

describe(commands.PAGE_CONTROL_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
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
    assert.strictEqual(command.name.startsWith(commands.PAGE_CONTROL_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about the control on a modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlGetData);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(mockControlGetDataOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the control on a modern page (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlGetData);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(mockControlGetDataOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the control on a modern page when the specified page name doesn\'t contain extension', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlGetData);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home', id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(mockControlGetDataOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles empty columns', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlGetDataEmptyColumn);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: '88f7b5b2-83a8-45d1-bc61-c11425f233e3' } }, () => {
      try {
        assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(mockControlGetDataEmptyColumnOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles text controls', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlGetDataWithText);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: '1212fc8d-dd6b-408a-8d5d-9f1cc787efbb' } }, () => {
      try {
        assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(mockControlGetDataWithTextOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles unknown types', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlGetDataWithUnknownType);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(mockControlGetDataWithUnknownTypeOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('correctly handles control not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlGetData);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e6' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles control not found (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlGetData);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e6' } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('Control with ID 3ede60d3-dc2c-438b-b5bf-cc40bb2351e6 not found on page home.aspx'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles control not found and no ID was specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve(mockControlGetDataWithoutAnId);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e6' } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('Control with ID 3ede60d3-dc2c-438b-b5bf-cc40bb2351e6 not found on page home.aspx'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles control not found on a page when CanvasContent1 is not defined', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return Promise.resolve({ CanvasContent1: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e6' } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('Control with ID 3ede60d3-dc2c-438b-b5bf-cc40bb2351e6 not found on page home.aspx'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles page not found', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-2130575338, Microsoft.SharePoint.SPException",
            "message": {
              "lang": "en-US",
              "value": "The file /sites/team-a/SitePages/home1.aspx does not exist."
            }
          }
        }
      });
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

  it('fails validation if the specified id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc', name: 'home.aspx', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', name: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL and name and id are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});