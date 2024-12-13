import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
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
import command from './page-default-get.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { formatting } from '../../../../utils/formatting.js';
import { pageListItemMock } from './page-get.mock.js';

describe(commands.PAGE_DEFAULT_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  const siteUrl = 'https://contoso.sharepoint.com/sites/Marketing';
  const serverRelativeUrl = urlUtil.getServerRelativeSiteUrl(siteUrl);
  const page = {
    WelcomePage: '/SitePages/Home.aspx'
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_DEFAULT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if siteUrl is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when the siteUrl is a valid SharePoint URL and name is specified', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: siteUrl });
    assert.strictEqual(actual.success, true);
  });

  it('gets the home page details for a specific site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/RootFolder?$select=WelcomePage`) {
        return page;
      }

      if (opts.url === `${siteUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${serverRelativeUrl}/${formatting.encodeQueryParameter(page.WelcomePage)}')?$expand=ListItemAllFields/ClientSideApplicationId,ListItemAllFields/PageLayoutType,ListItemAllFields/CommentsDisabled`) {
        return pageListItemMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, verbose: true } });
    assert(loggerLogSpy.calledWith({ ...pageListItemMock }));
  });

  it('correctly handles API error', async () => {
    const errorMessage = 'The file /sites/Marketing/SitePages/Welcome.aspx does not exist.';

    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          message: {
            lang: 'en-US',
            value: errorMessage
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl } }),
      new CommandError(errorMessage));
  });
});