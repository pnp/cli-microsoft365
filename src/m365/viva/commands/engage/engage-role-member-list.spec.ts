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
import command from './engage-role-member-list.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { vivaEngageRole } from '../../../../utils/vivaEngageRole.js';

describe(commands.ENGAGE_ROLE_MEMBER_LIST, () => {
  const roleId = 'ec759127-089f-4f91-8dfc-03a30b51cb38';
  const roleName = 'Network Admin';

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
    assert.strictEqual(command.name, commands.ENGAGE_ROLE_MEMBER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'userId']);
  });

  it('fails validation if roleId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if roleId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: roleId
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if roleName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleName: roleName
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both roleId and roleName are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: roleId,
      roleName: roleName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither roleId nor roleName is specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.notStrictEqual(actual.success, true);
  });

  it(`should get a list of users assigned to a Viva Engage role specified by id`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/roles/${roleId}/members`) {
        return {
          "value": [
            {
              "id": "1f5595b2-aa07-445d-9801-a45ea18160b2",
              "createdDateTime": "2024-05-22T15:43:08.368Z",
              "userId": "1f5595b2-aa07-445d-9801-a45ea18160b2"
            },
            {
              "id": "717f1683-00fa-488c-b68d-5d0051f6bcfa",
              "createdDateTime": "2025-07-16T02:51:22.602Z",
              "userId": "717f1683-00fa-488c-b68d-5d0051f6bcfa"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { verbose: true, roleId: roleId }
    });

    assert(
      loggerLogSpy.calledWith([
        {
          "id": "1f5595b2-aa07-445d-9801-a45ea18160b2",
          "createdDateTime": "2024-05-22T15:43:08.368Z",
          "userId": "1f5595b2-aa07-445d-9801-a45ea18160b2"
        },
        {
          "id": "717f1683-00fa-488c-b68d-5d0051f6bcfa",
          "createdDateTime": "2025-07-16T02:51:22.602Z",
          "userId": "717f1683-00fa-488c-b68d-5d0051f6bcfa"
        }
      ])
    );
  });

  it(`should get a list of users assigned to a Viva Engage role specified by names`, async () => {
    sinon.stub(vivaEngageRole, 'getRoleIdByName').resolves(roleId);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/roles/${roleId}/members`) {
        return {
          "value": [
            {
              "id": "1f5595b2-aa07-445d-9801-a45ea18160b2",
              "createdDateTime": "2024-05-22T15:43:08.368Z",
              "userId": "1f5595b2-aa07-445d-9801-a45ea18160b2"
            },
            {
              "id": "717f1683-00fa-488c-b68d-5d0051f6bcfa",
              "createdDateTime": "2025-07-16T02:51:22.602Z",
              "userId": "717f1683-00fa-488c-b68d-5d0051f6bcfa"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { verbose: true, roleName: roleName }
    });

    assert(
      loggerLogSpy.calledWith([
        {
          "id": "1f5595b2-aa07-445d-9801-a45ea18160b2",
          "createdDateTime": "2024-05-22T15:43:08.368Z",
          "userId": "1f5595b2-aa07-445d-9801-a45ea18160b2"
        },
        {
          "id": "717f1683-00fa-488c-b68d-5d0051f6bcfa",
          "createdDateTime": "2025-07-16T02:51:22.602Z",
          "userId": "717f1683-00fa-488c-b68d-5d0051f6bcfa"
        }
      ])
    );
  });

  it('handles error when retrieving Viva Engage roles failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/roles/${roleId}/members`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: { roleId: roleId } } as any),
      new CommandError('An error has occurred')
    );
  });
});