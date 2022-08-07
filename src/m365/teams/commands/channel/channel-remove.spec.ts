import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./channel-remove');

describe(commands.CHANNEL_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      request.get,
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CHANNEL_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when valid id & teamId is specified', async () => {
    const actual = await command.validate({
      options: {
        id: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name & teamId is specified', async () => {
    const actual = await command.validate({
      options: {
        name: 'Channel Name',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id and name are not provided', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not valid id', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        id: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid',
        id: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both name and id are provided', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        id: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        name: 'channelname'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails to remove channel when channel does not exists', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`channels?$filter=displayName eq 'name'`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {
      debug: true,
      teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      name: 'name',
      confirm: true } } as any), new CommandError('The specified channel does not exist in the Microsoft Teams team'));
  });

  it('prompts before removing the specified channel when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        debug: false,
        id: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('prompts before removing the specified channel when confirm option not passed (debug)', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        id: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified channel when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: {
        debug: true,
        id: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    });

    assert(postSpy.notCalled);
  });

  it('aborts removing the specified channel when confirm option not passed and prompt not confirmed (debug)', async () => {
    const postSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: {
        debug: true,
        id: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    });

    assert(postSpy.notCalled);
  });

  it('removes specified channel when channelId is passed with confirm option', async () => {
    sinon.stub(request, 'delete').returns(Promise.resolve());

    await command.action(logger, {
      options: {
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        confirm: true
      }
    });
  });

  it('removes the specified channel by name when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`channels?$filter=displayName eq 'name'`) > -1) {
        return Promise.resolve({
          value: [
            {
              "id": "19:f3dcbb1674574677abcae89cb626f1e6@thread.skype",
              "displayName": "name",
              "description": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19:f3dcbb1674574677abcae89cb626f1e6%40thread.skype/%F0%9F%92%A1+Ideas?groupId=d66b8110-fcad-49e8-8159-0d488ddb7656&tenantId=eff8592e-e14a-4ae8-8771-d96d5c549e1c"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/d66b8110-fcad-49e8-8159-0d488ddb7656/channels/19%3Af3dcbb1674574677abcae89cb626f1e6%40thread.skype`) {
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
        name: 'name',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    });
  });

  it('removes the specified channel by name without prompt', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`channels?$filter=displayName eq 'channelName'`) > -1) {
        return Promise.resolve({
          value: [
            {
              "id": "19:f3dcbb1674574677abcae89cb626f1e6@thread.skype",
              "displayName": "channelName",
              "description": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19:f3dcbb1674574677abcae89cb626f1e6%40thread.skype/%F0%9F%92%A1+Ideas?groupId=d66b8110-fcad-49e8-8159-0d488ddb7656&tenantId=eff8592e-e14a-4ae8-8771-d96d5c549e1c"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/d66b8110-fcad-49e8-8159-0d488ddb7656/channels/19%3Af3dcbb1674574677abcae89cb626f1e6%40thread.skype`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'channelName',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        confirm: true
      }
    });
  });

  it('should handle Microsoft graph error response', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/d66b8110-fcad-49e8-8159-0d488ddb7656/channels/19%3Af3dcbb1674574677abcae89cb626f1e6%40thread.skype`) {
        return Promise.reject({
          "error": {
            "code": "ItemNotFound",
            "message": "Failed to execute Skype backend request GetThreadS2SRequest.",
            "innerError": {
              "request-id": "5a563fc6-6df2-4cd9-b0b8-9810f1110714",
              "date": "2019-08-28T19:18:30"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await assert.rejects(command.action(logger, { options: {
      id: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
      teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656' } } as any), new CommandError('Failed to execute Skype backend request GetThreadS2SRequest.'));
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
