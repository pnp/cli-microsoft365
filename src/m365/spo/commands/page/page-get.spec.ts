import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { controlsMock, pageListItemMock, sectionMock, classicPage } from './page-get.mock';
const command: Command = require('./page-get');

describe(commands.PAGE_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['commentsDisabled', 'numSections', 'numControls', 'title', 'layoutType']);
  });

  it('gets information about a modern page including all returned properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return pageListItemMock;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
        return controlsMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', output: 'json' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].numControls, sectionMock.numControls);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].numSections, sectionMock.numSections);
  });

  it('gets information about a modern page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return pageListItemMock;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
        return controlsMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any);
    assert(loggerLogSpy.calledWith({
      ...pageListItemMock,
      canvasContentJson: controlsMock.CanvasContent1,
      ...sectionMock
    }));
  });

  it('gets information about a modern page on root of tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/SitePages/home.aspx')`) > -1) {
        return pageListItemMock;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
        return controlsMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx', output: 'json' } } as any);
    assert(loggerLogSpy.calledWith({
      ...pageListItemMock,
      canvasContentJson: controlsMock.CanvasContent1,
      ...sectionMock
    }));
  });

  it('gets information about a modern page when the specified page name doesn\'t contain extension', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return pageListItemMock;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
        return controlsMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home', output: 'json' } } as any);
    assert(loggerLogSpy.calledWith({
      ...pageListItemMock,
      canvasContentJson: controlsMock.CanvasContent1,
      ...sectionMock
    }));
  });

  it('check if section and control HTML parsing gets skipped for metadata only mode', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return pageListItemMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home', metadataOnly: true, output: 'json' } });
    assert(loggerLogSpy.calledWith(pageListItemMock));
  });

  it('shows error when the specified page is a classic page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return classicPage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any),
      new CommandError('Page home.aspx is not a modern page.'));
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

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any),
      new CommandError('The file /sites/team-a/SitePages/home1.aspx does not exist.'));
  });

  it('correctly handles OData error when retrieving pages', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any),
      new CommandError('An error has occurred'));
  });

  it('supports specifying metadataOnly flag', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--metadataOnly') {
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
