import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import commands from '../../commands.js';
import command from './administrativeunit-list.js';

describe(commands.ADMINISTRATIVEUNIT_LIST, () => {
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
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'visibility']);
  });

  it(`should get a list of administrative units`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({});
    await command.action(logger, { options: parsedSchema.data });

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

  it(`should get a list of administrative units with specified properties`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$select=id,displayName`) {
        return {
          value: [
            {
              id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
              displayName: 'European Division'
            },
            {
              id: 'a25b4c5e-e8b7-4f02-a23d-0965b6415098',
              displayName: 'Asian Division'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({ properties: 'id,displayName' });
    await command.action(logger, { options: parsedSchema.data });

    assert(
      loggerLogSpy.calledWith([
        {
          id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
          displayName: 'European Division'
        },
        {
          id: 'a25b4c5e-e8b7-4f02-a23d-0965b6415098',
          displayName: 'Asian Division'
        }
      ])
    );
  });

  it('handles error when retrieving administrative units list failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    const parsedSchema = commandOptionsSchema.safeParse({});
    await assert.rejects(
      command.action(logger, { options: parsedSchema.data }),
      new CommandError('An error has occurred')
    );
  });
});