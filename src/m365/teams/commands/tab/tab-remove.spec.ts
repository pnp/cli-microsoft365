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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./tab-remove');

describe(commands.TAB_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
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
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TAB_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when valid channelId, teamId and id is specified', async () => {
    const actual = await command.validate({
      options: {
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the teamId , channelId and id are not provided', async () => {
    const actual = await command.validate({
      options: {

      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the channelId is not valid channelId', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: 'invalid',
        id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid',
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
  it('fails validation if the id is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        id: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });


  it('prompts before removing the specified tab when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        debug: false,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('prompts before removing the specified tab when confirm option not passed (debug)', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified tab when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    });
    assert(postSpy.notCalled);
  });

  it('aborts removing the specified tab when confirm option not passed and prompt not confirmed (debug)', async () => {
    const postSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified tab by id when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`tabs/d66b8110-fcad-49e8-8159-0d488ddb7656`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    });
  });


  it('removes the specified tab without prompting when confirmed specified (debug)', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`tabs/d66b8110-fcad-49e8-8159-0d488ddb7656`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        id: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        confirm: true
      }
    });
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, { options: { 
      debug: false,
      channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
      teamId: '00000000-0000-0000-0000-000000000000',
      tabId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      confirm: true } } as any), new CommandError('An error has occurred'));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

});
