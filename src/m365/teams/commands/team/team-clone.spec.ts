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
const command: Command = require('./team-clone');

describe(commands.TEAM_CLONE, () => {
  let log: string[];
  let logger: Logger;
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.TEAM_CLONE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the id is not a valid GUID.', async () => {
    const actual = await command.validate({
      options: {
        id: 'invalid',
        partsToClone: "apps,tabs,settings,channels,members"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation on invalid visibility', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        name: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: 'abc'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation on valid \'private\' visibility', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: 'private'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'public\' visibility', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: 'public'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the input is correct with mandatory parameters', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        partsToClone: "apps,tabs,settings,channels,members"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the input is correct with mandatory and optional parameters', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        partsToClone: "apps,tabs,settings,channels,members",
        description: "Self help community for library",
        visibility: "public",
        classification: "public"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if visibility is set to private', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: "abc"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if partsToClone is set to invalid value', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        partsToClone: "abc"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if visibility is set to private', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: "private"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [{ options: ['id', 'name'] }]);
  });

  it('creates a clone of a Microsoft Teams team with mandatory parameters', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/15d7a78e-fd77-4599-97a5-dbb6372846c5/clone`) {
        return Promise.resolve({
          "location": "/teams('f9526e6a-1d0d-4421-8882-88a70975a00c')/operations('6cf64f96-08c3-4173-9919-eaf7684aae9a')"
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: false,
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        name: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members"
      }
    } as any);
  });

  it('creates a clone of a Microsoft Teams team with optional parameters (debug)', async () => {
    const sinonStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/15d7a78e-fd77-4599-97a5-dbb6372846c5/clone`) {
        return Promise.resolve({
          "location": "/teams('f9526e6a-1d0d-4421-8882-88a70975a00c')/operations('6cf64f96-08c3-4173-9919-eaf7684aae9a')"
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        name: 'Library Assist',
        partsToClone: 'apps,tabs,settings,channels,members',
        description: 'abc',
        visibility: 'public',
        classification: 'label'
      }
    } as any);
    assert.strictEqual(sinonStub.lastCall.args[0].url, 'https://graph.microsoft.com/v1.0/teams/15d7a78e-fd77-4599-97a5-dbb6372846c5/clone');
    assert.strictEqual(sinonStub.lastCall.args[0].data.displayName, 'Library Assist');
    assert.strictEqual(sinonStub.lastCall.args[0].data.partsToClone, 'apps,tabs,settings,channels,members');
    assert.strictEqual(sinonStub.lastCall.args[0].data.description, 'abc');
    assert.strictEqual(sinonStub.lastCall.args[0].data.visibility, 'public');
    assert.strictEqual(sinonStub.lastCall.args[0].data.classification, 'label');
    assert.notStrictEqual(sinonStub.lastCall.args[0].data.mailNickname.length, 0);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        name: 'Library Assist',
        partsToClone: 'apps,tabs,settings,channels,members',
        description: 'abc',
        visibility: 'public',
        classification: 'label'
      }
    } as any), new CommandError('An error has occurred'));
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
