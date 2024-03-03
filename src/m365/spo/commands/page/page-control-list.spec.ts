import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import { ClientSidePage } from './clientsidepages.js';
import command from './page-control-list.js';
import { mockControlListData, mockControlListDataOutput, mockControlListDataWithText, mockControlListDataWithTextOutput, mockControlListDataWithUnknownType, mockControlListDataWithUnknownTypeOutput } from './page-control-list.mock.js';

describe(commands.PAGE_CONTROL_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
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
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_CONTROL_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'type', 'title']);
  });

  it('lists controls on the modern page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return mockControlListData;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx' } });
    assert(loggerLogSpy.calledWith(mockControlListDataOutput));
  });

  it('lists controls on the modern page (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return mockControlListData;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx' } });
    assert(loggerLogSpy.calledWith(mockControlListDataOutput));
  });

  it('lists controls on the modern page when the specified page name doesn\'t contain extension', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return mockControlListData;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home' } });
    assert(loggerLogSpy.calledWith(mockControlListDataOutput));
  });

  it('handles empty columns and unknown control types', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return mockControlListDataWithUnknownType;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx' } });
    assert(loggerLogSpy.calledWith(mockControlListDataWithUnknownTypeOutput));
  });

  it('handles text web part correctly', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return mockControlListDataWithText;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx' } } as any);
    assert(loggerLogSpy.calledWith(mockControlListDataWithTextOutput));
  });

  it('correctly handles page when CanvasContent1 is not defined', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      return { CanvasContent1: null };
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx' } } as any);
    assert([]);
  });

  it('correctly handles page not found', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw {
        error: {
          "odata.error": {
            "code": "-2130575338, Microsoft.SharePoint.SPException",
            "message": {
              "lang": "en-US",
              "value": "The file /sites/team-a/SitePages/home1.aspx does not exist."
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx' } } as any),
      new CommandError('The file /sites/team-a/SitePages/home1.aspx does not exist.'));
  });

  it('correctly handles OData error when retrieving pages', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx' } } as any),
      new CommandError('An error has occurred'));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', pageName: 'home.aspx' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL and name is specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', pageName: 'home.aspx' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
