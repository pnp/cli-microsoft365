import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './report-siteversionstorageusage-status-get.js';

describe(commands.REPORT_SITEVERSIONSTORAGEUSAGE_STATUS_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const siteUrl = 'https://contoso.sharepoint.com/sites/marketing';
  const fileUrl = '/sites/marketing/Reports/FileUsage.csv';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
    auth.connection.active = true;
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.REPORT_SITEVERSIONSTORAGEUSAGE_STATUS_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if site URL is not a valid URL', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'foo', fileUrl: fileUrl });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation if fileUrl is not specified', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: siteUrl });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with valid options', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: siteUrl, fileUrl: fileUrl });
    assert.strictEqual(actual.success, true);
  });

  it('retrieves the status of the version storage usage report', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/site/GetProgressForFileVersionExpirationReport(reportFileUrl='${formatting.encodeQueryParameter(fileUrl)}')`) {
        return { value: JSON.stringify({ status: 'completed' }) };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: siteUrl, fileUrl: fileUrl, verbose: true }) });
    assert(loggerLogSpy.calledOnceWithExactly({ status: 'completed' }));
  });

  it('correctly handles error when the file is not in the specified site', async () => {
    const diffSiteFileUrl = '/teams/sales/shared documents/FileUsage.csv';

    sinon.stub(request, 'post').rejects({
      error: {
        'odata.error': {
          code: '-1, System.ArgumentException',
          message: { lang: 'en-US', value: `Invalid URL: ${diffSiteFileUrl}.` }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: siteUrl, fileUrl: diffSiteFileUrl }) }),
      new CommandError(`Invalid URL: ${diffSiteFileUrl}.`));
  });
});
