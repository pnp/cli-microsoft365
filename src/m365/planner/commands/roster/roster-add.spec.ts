import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { powerPlatform } from '../../../../utils/powerPlatform';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./roster-add');

describe(commands.ROSTER_ADD, () => {
  //#region Mocked Responses
  const rosterResponse = {
    "id": "-uQKZUR-qUSdmRUb051Z85gADf6O",
    "assignedSensitivityLabel": null
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      powerPlatform.getDynamicsInstanceApiUrl
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
    assert.strictEqual(command.name.startsWith(commands.ROSTER_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds a new roster', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === 'https://graph.microsoft.com/beta/planner/rosters')) {
        return rosterResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(rosterResponse));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(async () => { throw 'An error has occurred'; });

    await assert.rejects(command.action(logger, {
      options: {}
    }), new CommandError("An error has occurred"));
  });
});
