import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./messagingsettings-list');

describe(commands.MESSAGINGSETTINGS_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.MESSAGINGSETTINGS_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists messaging settings for a Microsoft Team', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=messagingSettings`) {
        return Promise.resolve({
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900" } });
    assert(loggerLogSpy.calledWith({
      "allowUserEditMessages": true,
      "allowUserDeleteMessages": true,
      "allowOwnerDeleteMessages": true,
      "allowTeamMentions": true,
      "allowChannelMentions": true
    }));
  });

  it('lists messaging settings for a Microsoft Team (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=messagingSettings`) {
        return Promise.resolve({
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", debug: true } });
    assert(loggerLogSpy.calledWith({
      "allowUserEditMessages": true,
      "allowUserDeleteMessages": true,
      "allowOwnerDeleteMessages": true,
      "allowTeamMentions": true,
      "allowChannelMentions": true
    }));
  });

  it('correctly handles error when listing messaging settings', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900" } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if teamId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when a valid teamId is specified', async () => {
    const actual = await command.validate({
      options: {
        teamId: '2609af39-7775-4f94-a3dc-0dd67657e900'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('lists all properties for output json', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=messagingSettings`) {
        return Promise.resolve({
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", output: 'json' } });
    assert(loggerLogSpy.calledWith({
      "allowUserEditMessages": true,
      "allowUserDeleteMessages": true,
      "allowOwnerDeleteMessages": true,
      "allowTeamMentions": true,
      "allowChannelMentions": true
    }));
  });
});