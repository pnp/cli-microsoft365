import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { odata } from '../../../../utils/odata';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as TeamGetCommand from './team-get';
const command: Command = require('./team-app-list');

describe(commands.TEAM_APP_LIST, () => {
  const teamId = '0ad55b5d-6a79-467b-ad21-d4bef7948a79';
  const teamName = 'Contoso Team';
  const jsonResponse = JSON.parse(`[{"id":"MGFkNTViNWQtNmE3OS00NjdiLWFkMjEtZDRiZWY3OTQ4YTc5IyMxNGQ2OTYyZC02ZWViLTRmNDgtODg5MC1kZTU1NDU0YmIxMzY=","teamsApp":{"id":"14d6962d-6eeb-4f48-8890-de55454bb136","externalId":null,"displayName":"Activity","distributionMethod":"store"},"teamsAppDefinition":{"id":"MTRkNjk2MmQtNmVlYi00ZjQ4LTg4OTAtZGU1NTQ1NGJiMTM2IyMxLjAjI1B1Ymxpc2hlZA==","teamsAppId":"14d6962d-6eeb-4f48-8890-de55454bb136","displayName":"Activity","version":"1.0","publishingState":"published","shortDescription":"Activity app bar entry.","description":"Activity app bar entry.","lastModifiedDateTime":null,"createdBy":null}},{"id":"MGFkNTViNWQtNmE3OS00NjdiLWFkMjEtZDRiZWY3OTQ4YTc5IyMyMGMzNDQwZC1jNjdlLTQ0MjAtOWY4MC0wZTUwYzM5NjkzZGY=","teamsApp":{"id":"20c3440d-c67e-4420-9f80-0e50c39693df","externalId":null,"displayName":"Calling","distributionMethod":"store"},"teamsAppDefinition":{"id":"MjBjMzQ0MGQtYzY3ZS00NDIwLTlmODAtMGU1MGMzOTY5M2RmIyMxLjAjI1B1Ymxpc2hlZA==","teamsAppId":"20c3440d-c67e-4420-9f80-0e50c39693df","displayName":"Calling","version":"1.0","publishingState":"published","shortDescription":"Calling app bar entry.","description":"Calling app bar entry.","lastModifiedDateTime":null,"createdBy":null}},{"id":"MGFkNTViNWQtNmE3OS00NjdiLWFkMjEtZDRiZWY3OTQ4YTc5IyMyYTg0OTE5Zi01OWQ4LTQ0NDEtYTk3NS0yYThjMjY0M2I3NDE=","teamsApp":{"id":"2a84919f-59d8-4441-a975-2a8c2643b741","externalId":null,"displayName":"Teams","distributionMethod":"store"},"teamsAppDefinition":{"id":"MmE4NDkxOWYtNTlkOC00NDQxLWE5NzUtMmE4YzI2NDNiNzQxIyMxLjAjI1B1Ymxpc2hlZA==","teamsAppId":"2a84919f-59d8-4441-a975-2a8c2643b741","displayName":"Teams","version":"1.0","publishingState":"published","shortDescription":"Teams app bar entry.","description":"Teams app bar entry.","lastModifiedDateTime":null,"createdBy":null}}]`);
  const friendlyResponse = [{ "id": "MGFkNTViNWQtNmE3OS00NjdiLWFkMjEtZDRiZWY3OTQ4YTc5IyMxNGQ2OTYyZC02ZWViLTRmNDgtODg5MC1kZTU1NDU0YmIxMzY=", "displayName": "Activity", "distributionMethod": "store" }, { "id": "MGFkNTViNWQtNmE3OS00NjdiLWFkMjEtZDRiZWY3OTQ4YTc5IyMyMGMzNDQwZC1jNjdlLTQ0MjAtOWY4MC0wZTUwYzM5NjkzZGY=", "displayName": "Calling", "distributionMethod": "store" }, { "id": "MGFkNTViNWQtNmE3OS00NjdiLWFkMjEtZDRiZWY3OTQ4YTc5IyMyYTg0OTE5Zi01OWQ4LTQ0NDEtYTk3NS0yYThjMjY0M2I3NDE=", "displayName": "Teams", "distributionMethod": "store" }];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      odata.getAllItems,
      Cli.executeCommandWithOutput
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
    assert.strictEqual(command.name.startsWith(commands.TEAM_APP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { teamId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the teamId is a valid GUID', async () => {
    const actual = await command.validate({ options: { teamId: teamId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails when team does not exist in tenant', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === TeamGetCommand) {
        throw 'The specified team does not exist in the Microsoft Teams';
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { teamName: teamName, verbose: true } }),
      new CommandError('The specified team does not exist in the Microsoft Teams'));
  });

  it('lists team apps for team specified by name with output json', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === TeamGetCommand) {
        return { "stdout": JSON.stringify({ id: teamId }) };
      }

      throw 'Invalid request';
    });

    sinon.stub(odata, 'getAllItems').callsFake(async (url: string): Promise<any> => {
      if (url === `https://graph.microsoft.com/v1.0/teams/${teamId}/installedApps?$expand=teamsApp,teamsAppDefinition`) {
        return jsonResponse;
      }

      throw 'Invalid response';
    });

    await command.action(logger, { options: { teamName: teamName, verbose: true, output: 'json' } });
    assert(loggerLogSpy.calledWith(jsonResponse));
  });

  it('lists team apps for team specified by id with output csv', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string): Promise<any> => {
      if (url === `https://graph.microsoft.com/v1.0/teams/${teamId}/installedApps?$expand=teamsApp,teamsAppDefinition`) {
        return jsonResponse;
      }

      throw 'Invalid response';
    });

    await command.action(logger, { options: { teamId: teamId, verbose: true, output: 'csv' } });
    assert(loggerLogSpy.calledWith(friendlyResponse));
  });
});
