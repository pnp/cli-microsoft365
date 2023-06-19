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
const command: Command = require('./site-recyclebinitem-remove');

describe(commands.SITE_RECYCLEBINITEM_REMOVE, () => {

  let log: any[];
  let logger: Logger;
  let promptOptions: any;
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_RECYCLEBINITEM_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo', ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1', confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if ids is not a valid guid array string', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,foo', confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all options are passed', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1', confirm: true } }, commandInfo);
    assert(actual);
  });

  it('prompts before removing the items from the recycle bin when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1'
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the items from the recycle bin when confirm option not passed and prompt not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves();
    await command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1'
      }
    });

    assert(postStub.notCalled);
  });

  it('removes items from the recycle bin with ids and confirm option', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/site/RecycleBin/DeleteByIds`) {
        return {
          'odata.null': true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1',
        confirm: true
      }
    });
  });

  it('removes items from the recycle bin with ids', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/site/RecycleBin/DeleteByIds`) {
        return {
          'odata.null': true
        };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        verbose: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1'
      }
    });
  });

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          message: {
            value: 'Specified argument was out of the range of valid values'
          }
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/site/RecycleBin/DeleteByIds`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1',
        confirm: true
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });
});