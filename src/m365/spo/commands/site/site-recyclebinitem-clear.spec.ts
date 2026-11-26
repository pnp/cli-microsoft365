import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './site-recyclebinitem-clear.js';

describe(commands.SITE_RECYCLEBINITEM_CLEAR, () => {

  let log: any[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_RECYCLEBINITEM_CLEAR);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'foo', force: true });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com', force: true });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('prompts before removing the items from the recycle bin when force option not passed', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        siteUrl: 'https://contoso.sharepoint.com'
      })
    });

    assert(promptIssued);
  });

  it('aborts removing the items from the recycle bin when force option not passed and prompt not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/RecycleBin/DeleteAll`) {
        return {
          'odata.null': true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        siteUrl: 'https://contoso.sharepoint.com'
      })
    });

    assert(postStub.notCalled);
  });

  it('removes all items from the first-stage recycle bin with force option', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/RecycleBin/DeleteAll`) {
        return {
          'odata.null': true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        siteUrl: 'https://contoso.sharepoint.com',
        force: true
      })
    });
  });

  it('removes all items from the first-stage recycle bin without confirmation', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/RecycleBin/DeleteAll`) {
        return {
          'odata.null': true
        };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        siteUrl: 'https://contoso.sharepoint.com'
      })
    });

    assert(postStub.called);
  });

  it('removes all items from the second-stage recycle bin with force option', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/site/RecycleBin/DeleteAllSecondStageItems`) {
        return {
          'odata.null': true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        siteUrl: 'https://contoso.sharepoint.com',
        secondary: true,
        force: true
      })
    });

    assert(postStub.called);
  });

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          message: {
            value: "The files cannot be removed from the second-stage recycle bin."
          }
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async () => {
      return error;
    });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com', force: true }) }),
      new CommandError(error.error['odata.error'].message.value)
    );
  });
});
