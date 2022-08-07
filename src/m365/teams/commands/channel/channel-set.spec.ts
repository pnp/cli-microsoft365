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
const command: Command = require('./channel-set');

describe(commands.CHANNEL_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
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
    assert.strictEqual(command.name.startsWith(commands.CHANNEL_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly validates the arguments', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Reviews',
        newName: 'Gen',
        description: 'this is a new description'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid',
        name: 'Reviews',
        newName: 'Gen',
        description: 'this is a new description'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when name is General', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'General',
        newName: 'Reviews',
        description: 'this is a new description'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails to patch channel updates for the Microsoft Teams team when channel does not exists', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`channels?$filter=displayName eq 'Latest'`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {
      debug: true,
      teamId: '00000000-0000-0000-0000-000000000000',
      name: 'Latest',
      newName: 'New Review',
      description: 'New Review' } } as any), new CommandError('The specified channel does not exist in the Microsoft Teams team'));
  });

  it('correctly patches channel updates for the Microsoft Teams team', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`channels?$filter=displayName eq 'Review'`) > -1) {
        return Promise.resolve({
          value:
            [
              {
                "id": "19:8a53185a51ac44a3aef27397c3dfebfc@thread.skype",
                "displayName": "Review",
                "description": "Updated by CLI"
              }]
        });
      }
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (((opts.url as string).indexOf(`channels/19:8a53185a51ac44a3aef27397c3dfebfc@thread.skype`) > -1) &&
        JSON.stringify(opts.data) === JSON.stringify({ displayName: "New Review", description: "New Review" })
      ) {
        return Promise.resolve({});
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: false,
        teamId: '00000000-0000-0000-0000-000000000000',
        name: 'Review',
        newName: 'New Review',
        description: 'New Review'
      }
    } as any);
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