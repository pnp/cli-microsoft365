import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-list.js';

describe(commands.APP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['appId', 'id', 'displayName', 'signInAudience']);
  });

  it('passes validation when no properties are specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when properties are specified', () => {
    const actual = commandOptionsSchema.safeParse({ properties: 'id,displayName' });
    assert.strictEqual(actual.success, true);
  });

  it(`should get a list of Microsoft Entra app registrations`, async () => {
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
      options: commandOptionsSchema.parse({})
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

  it(`should get a list of Microsoft Entra app registrations with specified properties`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications?$select=id,displayName`) {
        return {
          value: [
            {
              id: '340a4aa3-1af6-43ac-87d8-189819003952',
              displayName: 'My App 1'
            },
            {
              id: '340a4aa3-1af6-43ac-87d8-189819003953',
              displayName: 'My App 2'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({ properties: 'id,displayName' })
    });

    assert(
      loggerLogSpy.calledWith([
        {
          id: '340a4aa3-1af6-43ac-87d8-189819003952',
          displayName: 'My App 1'
        },
        {
          id: '340a4aa3-1af6-43ac-87d8-189819003953',
          displayName: 'My App 2'
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
      command.action(logger, {
        options: commandOptionsSchema.parse({})
      }),
      new CommandError('An error has occurred')
    );
  });
});
