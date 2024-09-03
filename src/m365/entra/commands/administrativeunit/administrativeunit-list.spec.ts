import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { misc } from '../../../../utils/misc.js';
import { MockRequests } from '../../../../utils/MockRequest.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';
import command from './administrativeunit-list.js';

export const mocks = {
  administrativeUnits: {
    request: {
      url: `https://graph.microsoft.com/v1.0/directory/administrativeUnits`
    },
    response: {
      body: {
        value: [
          {
            id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
            displayName: 'European Division',
            visibility: 'HiddenMembership'
          },
          {
            id: 'a25b4c5e-e8b7-4f02-a23d-0965b6415098',
            displayName: 'Asian Division',
            visibility: null
          }
        ]
      }
    }
  }
} satisfies MockRequests;

describe(commands.ADMINISTRATIVEUNIT_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.ADMINISTRATIVEUNIT_LIST]);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'visibility']);
  });

  it(`should get a list of administrative units`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.administrativeUnits.request.url) {
        return misc.deepClone(mocks.administrativeUnits.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {}
    });

    assert(
      loggerLogSpy.calledWith([
        {
          id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
          displayName: 'European Division',
          visibility: 'HiddenMembership'
        },
        {
          id: 'a25b4c5e-e8b7-4f02-a23d-0965b6415098',
          displayName: 'Asian Division',
          visibility: null
        }
      ])
    );
  });

  it('handles error when retrieving administrative units list failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.administrativeUnits.request.url) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: {} } as any),
      new CommandError('An error has occurred')
    );
  });
});