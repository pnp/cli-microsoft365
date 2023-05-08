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
const command: Command = require('./site-recyclebinitem-move');

describe(commands.SITE_RECYCLEBINITEM_MOVE, () => {

  let log: any[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_RECYCLEBINITEM_MOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo', all: true, confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', all: true, confirm: true } }, commandInfo);
    assert(actual);
  });

  it('fails validation if ids is not a valid guid array string', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '85528dee-00d5-4c38-a6ba-e2abace32f63, foo', confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if ids has a valid value', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '85528dee-00d5-4c38-a6ba-e2abace32f63, aecb840f-20e9-4ff8-accf-5df8eaad31a1', confirm: true } }, commandInfo);
    assert(actual);
  });

  it('prompts before moving the items to the second-stage recycle bin when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com',
        all: true
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts moving the items to the second-stage recycle bin when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    await command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com',
        all: true
      }
    });

    assert(postSpy.notCalled);
  });

  it('moves items to the second-stage recycle bin with ids and confirm option', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/recycleBin/MoveAllToSecondStage`) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '85528dee-00d5-4c38-a6ba-e2abace32f63, aecb840f-20e9-4ff8-accf-5df8eaad31a1',
        confirm: true
      }
    });
  });

  it('moves all items to the second-stage recycle bin with all option', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/recycleBin/MoveAllToSecondStage`) {
        return {
          "odata.null": true
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
        siteUrl: 'https://contoso.sharepoint.com',
        all: true
      }
    });

    assert(postSpy.called);
  });

  it('moves all items to the second-stage recycle bin with all and confirm option', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/recycleBin/MoveAllToSecondStage`) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        siteUrl: 'https://contoso.sharepoint.com',
        all: true,
        confirm: true
      }
    });

    assert(postSpy.called);
  });

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          message: {
            value: "Value does not fall within the expected range."
          }
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async () => {
      throw error;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com', all: true, confirm: true } } as any), new CommandError(error.error['odata.error'].message.value));
  });
});
