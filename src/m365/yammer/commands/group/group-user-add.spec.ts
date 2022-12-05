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
const command: Command = require('./group-user-add');

describe(commands.GROUP_USER_ADD, () => {
  let log: string[];
  let logger: Logger;
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
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GROUP_USER_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    await assert.rejects(command.action(logger, { options: { debug: false } } as any), new CommandError('An error has occurred.'));
  });

  it('passes validation with parameters', async () => {
    const actual = await command.validate({ options: { groupId: 10123123 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('groupId must be a number', async () => {
    const actual = await command.validate({ options: { groupId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('id must be a number', async () => {
    const actual = await command.validate({ options: { groupId: 10, id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
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

  it('calls the service if the current user is added to the group', async () => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, { options: { debug: true, groupId: 1231231 } });

    assert(requestPostedStub.called);
  });

  it('calls the service if the user 989998789 is added to the group 1231231', async () => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, groupId: 1231231, id: 989998789 } });

    assert(requestPostedStub.called);
  });

  it('calls the service if the user suzy@contoso.com is added to the group 1231231', async () => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, groupId: 1231231, email: "suzy@contoso.com" } });

    assert(requestPostedStub.called);
  });
}); 