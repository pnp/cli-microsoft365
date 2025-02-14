
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
import command from './engage-community-user-list.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';

describe(commands.ENGAGE_COMMUNITY_USER_LIST, () => {
  const communityId = 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIzNjAyMDAxMTAwOSJ9';
  const communityDisplayName = 'All company';
  const entraGroupId = 'b6c35b51-ebca-445c-885a-63a67d24cb53';
  const membersAPIResult = [
    {
      "id": "1deb8814-8130-451d-8fcb-849dc7ed47e5",
      "businessPhones": [
        "123-555-1215"
      ],
      "displayName": "Samu Tolonen",
      "givenName": "Samu",
      "jobTitle": "IT Manager",
      "mail": null,
      "mobilePhone": "123-555-6645",
      "officeLocation": "123455",
      "preferredLanguage": null,
      "surname": "Tolonen",
      "userPrincipalName": "Samu.Tolonen@contoso.onmicrosoft.com"
    }
  ];
  const membersResult = [{ ...membersAPIResult[0], roles: ["Member"] }];
  const adminsAPIResult = [
    {
      "id": "da634de7-d23c-4419-ab83-fcd395b4ebd0",
      "businessPhones": [
        "123-555-1215"
      ],
      "displayName": "Anton Johansen",
      "givenName": "Anton",
      "jobTitle": "IT Manager",
      "mail": null,
      "mobilePhone": "123-555-6645",
      "officeLocation": "123455",
      "preferredLanguage": null,
      "surname": "Johansen",
      "userPrincipalName": "Anton.Johansen@contoso.onmicrosoft.com"
    }
  ];
  const adminsResult = [{ ...adminsAPIResult[0], roles: ["Admin"] }];
  const community = {
    id: communityId,
    displayName: communityDisplayName,
    privacy: 'Public',
    groupId: entraGroupId
  };

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
      request.get,
      vivaEngage.getCommunityById
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENGAGE_COMMUNITY_USER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if entraGroupId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      entraGroupId: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if communityId, communityDisplayName or entraGroupId are not specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if communityId, communityDisplayName and entraGroupId are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      communityId: communityId,
      communityDisplayName: communityDisplayName,
      entraGroupId: entraGroupId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if incorrect role value is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      communityId: communityId,
      role: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if communityId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      communityId: communityId
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if entraGroupId is specified with a proper GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      entraGroupId: entraGroupId
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if communityDisplayName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      communityDisplayName: communityDisplayName
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if role is specified with a proper value', () => {
    const actual = commandOptionsSchema.safeParse({
      communityId: communityId,
      role: 'Admin'
    });
    assert.strictEqual(actual.success, true);
  });

  it('correctly gets the list of users in the community by entraGroupId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/owners`) {
        return { value: adminsAPIResult };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/members`) {
        return { value: membersAPIResult };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { entraGroupId: entraGroupId, verbose: true } });
    assert(loggerLogSpy.calledWith([...adminsResult, ...membersResult]));
  });

  it('correctly gets the list of users in the community by communityId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/owners`) {
        return { value: adminsAPIResult };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/members`) {
        return { value: membersAPIResult };
      }

      throw 'Invalid request';
    });

    sinon.stub(vivaEngage, 'getCommunityById').resolves(community);

    await command.action(logger, { options: { communityId: communityId, verbose: true } });
    assert(loggerLogSpy.calledWith([...adminsResult, ...membersResult]));
  });

  it('correctly gets the list of users in the community by communityName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/owners`) {
        return { value: adminsAPIResult };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/members`) {
        return { value: membersAPIResult };
      }

      throw 'Invalid request';
    });

    sinon.stub(vivaEngage, 'getCommunityByDisplayName').resolves(community);

    await command.action(logger, { options: { communityDisplayName: communityDisplayName, verbose: true } });
    assert(loggerLogSpy.calledWith([...adminsResult, ...membersResult]));
  });

  it('correctly gets the list of members in the community by entraGroupId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/owners`) {
        return { value: adminsAPIResult };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/members`) {
        return { value: membersAPIResult };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { entraGroupId: entraGroupId, role: 'Member' } });
    assert(loggerLogSpy.calledWith(membersResult));
  });

  it('correctly gets the list of admins in the community by communityId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/owners`) {
        return { value: adminsAPIResult };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/members`) {
        return { value: membersAPIResult };
      }

      throw 'Invalid request';
    });

    sinon.stub(vivaEngage, 'getCommunityById').resolves(community);

    await command.action(logger, { options: { communityId: communityId, role: 'Admin' } });
    assert(loggerLogSpy.calledWith(adminsResult));
  });

  it('correctly handles error', async () => {
    const errorMessage = 'Bad request.';
    sinon.stub(request, 'get').rejects({
      error: {
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { id: 'invalid', verbose: true } }),
      new CommandError(errorMessage));
  });
});