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
const command: Command = require('./funsettings-set');

describe(commands.FUNSETTINGS_SET, () => {
  let log: string[];
  let logger: Logger;
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FUNSETTINGS_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets allowGiphy settings to false', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowGiphy: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: { teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowGiphy: false }
    } as any);
  });

  it('sets allowGiphy settings to true', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowGiphy: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: { teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowGiphy: true }
    } as any);
  });

  it('sets giphyContentRating to moderate', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            giphyContentRating: 'moderate'
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: { teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', giphyContentRating: 'moderate' }
    } as any);
  });

  it('sets giphyContentRating to strict', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            giphyContentRating: 'strict'
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: { teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', giphyContentRating: 'strict' }
    } as any);
  });

  it('sets allowStickersAndMemes to true', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowStickersAndMemes: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: { teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowStickersAndMemes: true }
    } as any);
  });

  it('sets allowStickersAndMemes to false', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowStickersAndMemes: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: { teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowStickersAndMemes: false }
    } as any);
  });


  it('sets allowCustomMemes to true', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowCustomMemes: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: { teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: true }
    } as any);
  });

  it('sets allowCustomMemes to false', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowCustomMemes: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: { debug: true, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: false }
    } as any);
  });

  it('sets allowCustomMemes to false (debug)', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowCustomMemes: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: { debug: true, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: false }
    } as any);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'patch').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        allowGiphy: true,
        giphyContentRating: "moderate",
        allowStickersAndMemes: false,
        allowCustomMemes: true
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if teamId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: { teamId: 'invalid' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when teamId is a valid GUID', async () => {
    const actual = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when giphyContentRating is moderate or strict', async () => {
    const actualModerate = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'moderate' }
    }, commandInfo);

    const actualStrict = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'strict' }
    }, commandInfo);

    const actual = actualModerate && actualStrict;
    assert.strictEqual(actual, true);
  });

  it('fails validation when giphyContentRating is not moderate or strict', async () => {
    const actual = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'somethingelse' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when allowStickersAndMemes is a valid boolean', async () => {
    const actualTrue = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowStickersAndMemes: true }
    }, commandInfo);

    const actualFalse = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowStickersAndMemes: false }
    }, commandInfo);

    const actual = actualTrue && actualFalse;
    assert.strictEqual(actual, true);
  });
});