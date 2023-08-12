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
import command from './roledefinition-remove.js';

describe(commands.ROLEDEFINITION_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptOptions: any;

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
    requests = [];
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROLEDEFINITION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if id is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if id is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '1' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('remove role definitions handles reject request correctly', async () => {
    const err = 'request rejected';
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/roledefinitions') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 1,
        force: true
      }
    }), new CommandError(err));
  });

  it('prompts before removing role definition when confirmation argument not passed', async () => {
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com', id: 1 } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing role definition when prompt not confirmed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com', id: 1 } });
    assert(requests.length === 0);
  });

  it('removes the role definition when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/roledefinitions(1)`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return '';
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com', id: 1 } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/roledefinitions(1)`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the role definition without confirm prompt', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/roledefinitions(1)`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return '';
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com', id: 1, force: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/roledefinitions(1)`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });
});
