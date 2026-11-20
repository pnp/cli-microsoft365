import assert from 'assert';
import sinon from 'sinon';
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
import command, { options } from './channel-list.js';

describe(commands.CHANNEL_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CHANNEL_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both teamId and teamName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '00000000-0000-0000-0000-000000000000',
      teamName: 'Team Name'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both channelId and channelName options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the teamId is not a valid guid.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '00000000-0000'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when invalid type specified', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '00000000-0000-0000-0000-000000000000',
      type: 'Invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName']);
  });

  it('validates for a correct input.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '00000000-0000-0000-0000-000000000000'
    });
    assert.strictEqual(actual.success, true);
  });

  it('rejects invalid channel type', () => {
    const type = 'foo';
    const actual = commandOptionsSchema.safeParse({ teamId: '00000000-0000-0000-0000-000000000000', type: type });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('correctly lists all channels in a Microsoft teams team by team id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return {
          value: [
            {
              "id": "19:17de660d16844149ab3f0240405f9316@thread.skype",
              "displayName": "General",
              "description": "Test group for office cli commands",
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            },
            {
              "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
              "displayName": "Development",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            },
            {
              "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
              "displayName": "Social",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: '00000000-0000-0000-0000-000000000000'
      })
    });
    assert(loggerLogSpy.calledWith(
      [
        {
          "id": "19:17de660d16844149ab3f0240405f9316@thread.skype",
          "displayName": "General",
          "description": "Test group for office cli commands",
          "isFavoriteByDefault": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
        },
        {
          "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
          "displayName": "Development",
          "description": null,
          "isFavoriteByDefault": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
        },
        {
          "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
          "displayName": "Social",
          "description": null,
          "isFavoriteByDefault": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
        }
      ]
    ));
  });

  it('fails when team name does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/teams?$filter=displayName eq '`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 0,
          "value": []
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ teamName: 'Team Name' }) }), new CommandError("The specified team 'Team Name' does not exist."));
  });

  it('correctly lists all channels in a Microsoft teams team with specified type parameter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels?$filter=membershipType eq 'private'`) {
        return {
          value: [
            {
              "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
              "displayName": "Development",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8",
              "membershipType": "private"
            },
            {
              "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
              "displayName": "Social",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8",
              "membershipType": "private"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: '00000000-0000-0000-0000-000000000000',
        type: 'private'
      })
    });
    assert(loggerLogSpy.calledWith(
      [
        {
          "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
          "displayName": "Development",
          "description": null,
          "isFavoriteByDefault": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8",
          "membershipType": "private"
        },
        {
          "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
          "displayName": "Social",
          "description": null,
          "isFavoriteByDefault": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8",
          "membershipType": "private"
        }
      ]
    ));
  });

  it('correctly lists all channels in a Microsoft teams team by team name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/teams?$filter=displayName eq '`) > -1) {
        return {
          "value": [
            {
              "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
              "id": "00000000-0000-0000-0000-000000000000",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2020-10-11T09:35:26Z",
              "creationOptions": [
                "Team",
                "ExchangeProvisioningFlags:3552"
              ],
              "description": "Team Description",
              "displayName": "Team Name",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "TeamName@contoso.com",
              "mailEnabled": true,
              "mailNickname": "TeamName",
              "membershipRule": null,
              "membershipRuleProcessingState": null,
              "onPremisesDomainName": null,
              "onPremisesLastSyncDateTime": null,
              "onPremisesNetBiosName": null,
              "onPremisesSamAccountName": null,
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "preferredLanguage": null,
              "proxyAddresses": [
                "SPO:SPO_97df7113-c3f3-447f-8010-9f88eb0fc7f1@SPO_00000000-0000-0000-0000-000000000000",
                "SMTP:TeamName@contoso.com"
              ],
              "renewedDateTime": "2020-10-11T09:35:26Z",
              "resourceBehaviorOptions": [
                "HideGroupInOutlook",
                "SubscribeMembersToCalendarEventsDisabled",
                "WelcomeEmailDisabled"
              ],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-1927732186-1159088485-2915259540-28248825",
              "theme": null,
              "visibility": "Private",
              "onPremisesProvisioningErrors": []
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return {
          value: [
            {
              "id": "19:17de660d16844149ab3f0240405f9316@thread.skype",
              "displayName": "General",
              "description": "Test group for office cli commands",
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            },
            {
              "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
              "displayName": "Development",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            },
            {
              "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
              "displayName": "Social",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        teamName: 'Team Name'
      })
    });
    assert(loggerLogSpy.calledWith(
      [
        {
          "id": "19:17de660d16844149ab3f0240405f9316@thread.skype",
          "displayName": "General",
          "description": "Test group for office cli commands",
          "isFavoriteByDefault": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
        },
        {
          "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
          "displayName": "Development",
          "description": null,
          "isFavoriteByDefault": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
        },
        {
          "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
          "displayName": "Social",
          "description": null,
          "isFavoriteByDefault": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
        }
      ]
    ));
  });

  it('correctly lists all channels in a Microsoft teams team (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return {
          value: [{ "id": "19:17de660d16844149ab3f0240405f9316@thread.skype", "displayName": "General", "description": "Test group for office cli commands", "isFavoriteByDefault": null, "email": "", "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8" }, { "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype", "displayName": "Development", "description": null, "isFavoriteByDefault": null, "email": "", "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8" }, { "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype", "displayName": "Social", "description": null, "isFavoriteByDefault": null, "email": "", "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8" }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, teamId: "00000000-0000-0000-0000-000000000000" }) });
    assert(loggerLogSpy.calledWith([{ "id": "19:17de660d16844149ab3f0240405f9316@thread.skype", "displayName": "General", "description": "Test group for office cli commands", "isFavoriteByDefault": null, "email": "", "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8" }, { "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype", "displayName": "Development", "description": null, "isFavoriteByDefault": null, "email": "", "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8" }, { "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype", "displayName": "Social", "description": null, "isFavoriteByDefault": null, "email": "", "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8" }]));
  });

  it('outputs all data in json output mode', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return {
          value: [
            {
              "id": "19:17de660d16844149ab3f0240405f9316@thread.skype",
              "displayName": "General",
              "description": "Test group for office cli commands",
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            },
            {
              "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
              "displayName": "Development",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            },
            {
              "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
              "displayName": "Social",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        output: 'json',
        teamId: '00000000-0000-0000-0000-000000000000'
      })
    });
    assert(loggerLogSpy.calledWith(
      [
        {
          "id": "19:17de660d16844149ab3f0240405f9316@thread.skype",
          "displayName": "General",
          "description": "Test group for office cli commands",
          "isFavoriteByDefault": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
        },
        {
          "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
          "displayName": "Development",
          "description": null,
          "isFavoriteByDefault": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
        },
        {
          "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
          "displayName": "Social",
          "description": null,
          "isFavoriteByDefault": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
        }
      ]
    ));
  });

  it('correctly handles error when retrieving all teams', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: '00000000-0000-0000-0000-000000000000'
      })
    }), new CommandError('An error has occurred'));
  });
});
