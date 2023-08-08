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
import command from './web-roleinheritance-break.js';

describe(commands.WEB_ROLEINHERITANCE_BREAK, () => {
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
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
    assert.strictEqual(command.name.startsWith(commands.WEB_ROLEINHERITANCE_BREAK), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should prompt before breaking when confirmation argument not passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/subsite/_api/web/breakroleinheritance(true)') {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: "https://contoso.sharepoint.com/subsite" } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
  });

  it('breaks inheritance successfully when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/subsite/_api/web/breakroleinheritance(true)') {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite"
      }
    });
  });

  it('does not break inheritance when prompt declined', async () => {
    const sinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/subsite/_api/web/breakroleinheritance(true)') {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite"
      }
    });

    assert(sinonStub.notCalled);
  });

  it('breaks role inheritance on web and clear all permissions', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/subsite/_api/web/breakroleinheritance(false)') {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/subsite',
        clearExistingPermissions: true,
        force: true
      }
    });
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'post').callsFake(async () => { throw { error: { message: errorMessage } }; });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/subsite',
        force: true
      }
    }), new CommandError(errorMessage));
  });
});
