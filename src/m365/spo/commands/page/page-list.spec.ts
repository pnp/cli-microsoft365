import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './page-list.js';
import { fileApiListData, mockPageListData, mockPagesListOutput } from './page-list.mock.js';

describe(commands.PAGE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    commandInfo = Cli.getCommandInfo(command);
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Name', 'Title']);
  });

  it('lists all modern pages', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages?$orderby=Title`) > -1) {
        return { value: mockPageListData };
      }

      if ((opts.url as string).indexOf(`/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`) > -1) {
        return { value: fileApiListData };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a' } });
    assert(loggerLogSpy.calledWith(mockPagesListOutput));
  });

  it('lists all modern pages (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages?$orderby=Title`) > -1) {
        return { value: mockPageListData };
      }

      if ((opts.url as string).indexOf(`/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`) > -1) {
        return { value: fileApiListData };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } });
    assert(loggerLogSpy.calledWith(mockPagesListOutput));
  });

  it('correctly handles no modern pages', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages?$orderby=Title`) > -1) {
        return { value: [] };
      }

      if ((opts.url as string).indexOf(`/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`) > -1) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles OData error when retrieving pages', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a' } } as any),
      new CommandError('An error has occurred'));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
