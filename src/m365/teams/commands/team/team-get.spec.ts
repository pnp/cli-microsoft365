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
import command from './team-get.js';
import { teams } from '../../../../utils/teams.js';

describe(commands.TEAM_GET, () => {
  const teamId = '1caf7dcd-7e83-4c3a-94f7-932a1299c844';
  const teamName = 'Finance';
  const teamResponse: any = {
    id: teamId,
    createdDateTime: '2017-11-29T03:27:05Z',
    displayName: teamName,
    description: 'This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.',
    classification: null,
    specialization: 'none',
    visibility: 'Public',
    webUrl: 'https://teams.microsoft.com/l/team/19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01%40thread.tacv2/conversations?groupId=1caf7dcd-7e83-4c3a-94f7-932a1299c844&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35',
    isArchived: false,
    isMembershipLimitedToOwners: false,
    discoverySettings: {
      showInTeamsSearchAndSuggestions: false
    },
    memberSettings: {
      allowCreateUpdateChannels: true,
      allowCreatePrivateChannels: true,
      allowDeleteChannels: true,
      allowAddRemoveApps: true,
      allowCreateUpdateRemoveTabs: true,
      allowCreateUpdateRemoveConnectors: true
    },
    guestSettings: {
      allowCreateUpdateChannels: false,
      allowDeleteChannels: false
    },
    messagingSettings: {
      allowUserEditMessages: true,
      allowUserDeleteMessages: true,
      allowOwnerDeleteMessages: true,
      allowTeamMentions: true,
      allowChannelMentions: true
    },
    funSettings: {
      allowGiphy: true,
      giphyContentRating: 'moderate',
      allowStickersAndMemes: true,
      allowCustomMemes: true
    }
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
      cli.getSettingWithDefaultValue,
      teams.getTeamByDisplayName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TEAM_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the teamId is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: teamId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails to get team information due to wrong team id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}`) {
        throw {
          error: {
            code: 'NotFound',
            message: `No team found with Group Id ${teamId}`,
            innerError: {
              message: `No team found with Group Id ${teamId}`,
              code: 'ItemNotFound',
              innerError: {},
              date: '2021-09-23T01:26:41',
              'request-id': '717697d2-b63d-422f-863c-d74d0c1c8c6f'
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: teamId
      }
    } as any), new CommandError(`No team found with Group Id ${teamId}`));
  });

  it('fails when team name does not exist', async () => {
    sinon.stub(teams, 'getTeamByDisplayName').rejects(new Error(`The specified team '${teamName}' does not exist.`));

    await assert.rejects(command.action(logger, {
      options: {
        name: teamName
      }
    } as any), new CommandError(`The specified team '${teamName}' does not exist.`));
  });

  it('retrieves information about the specified Microsoft Team', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}`) {
        return teamResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } });
    assert(loggerLogSpy.calledWith(teamResponse));
  });

  it('retrieves information about the specified Microsoft Teams team by name', async () => {
    sinon.stub(teams, 'getTeamByDisplayName').withArgs(teamName).resolves(teamResponse);

    await command.action(logger, { options: { name: teamName } });
    assert(loggerLogSpy.calledWith(teamResponse));
  });
});
