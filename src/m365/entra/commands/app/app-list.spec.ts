import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-list.js';
import aadCommands from '../../aadCommands.js';
import { MockRequests } from '../../../../utils/MockRequest.js';
import { misc } from '../../../../utils/misc.js';

export const mocks = {
  getApps: {
    request: {
      url: `https://graph.microsoft.com/v1.0/applications`
    },
    response: {
      body: {
        value: [
          {
            id: '340a4aa3-1af6-43ac-87d8-189819003952',
            appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
            displayName: 'My App 1',
            description: 'My second app',
            signInAudience: 'My Audience'
          },
          {
            id: '340a4aa3-1af6-43ac-87d8-189819003953',
            appId: '9b1b1e42-794b-4c71-93ac-5ed92488b670',
            displayName: 'My App 2',
            description: 'My second app',
            signInAudience: 'My Audience'
          }
        ]
      }
    }
  }
} satisfies MockRequests;

describe(commands.APP_LIST, () => {
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
    assert.strictEqual(command.name, commands.APP_LIST);
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
    assert.deepStrictEqual(alias, [aadCommands.APP_LIST, commands.APPREGISTRATION_LIST]);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['appId', 'id', 'displayName', 'signInAudience']);
  });

  it(`should get a list of Microsoft Entra app registrations`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.getApps.request.url) {
        return misc.deepClone(mocks.getApps.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {}
    });

    assert(
      loggerLogSpy.calledWith([
        {
          id: '340a4aa3-1af6-43ac-87d8-189819003952',
          appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          displayName: 'My App 1',
          description: 'My second app',
          signInAudience: 'My Audience'
        },
        {
          id: '340a4aa3-1af6-43ac-87d8-189819003953',
          appId: '9b1b1e42-794b-4c71-93ac-5ed92488b670',
          displayName: 'My App 2',
          description: 'My second app',
          signInAudience: 'My Audience'
        }
      ])
    );
  });

  it('handles error when retrieving app list failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications`) {
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
