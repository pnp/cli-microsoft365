import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./app-list');

describe(commands.APP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => {});
    sinon.stub(pid, 'getProcessName').callsFake(() => undefined);
    auth.service.connected = true;
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
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.APP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['appId', 'id', 'displayName', 'signInAudience']);
  });

  it(`should get a list of Azure AD app registrations`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications`) {
        return {
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
        };
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

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
