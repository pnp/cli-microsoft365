import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./membersettings-list');

describe(commands.MEMBERSETTINGS_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MEMBERSETTINGS_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists member settings for a Microsoft Team', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=memberSettings`) {
        return {
          "memberSettings": {
            "allowCreateUpdateChannels": true,
            "allowDeleteChannels": true,
            "allowAddRemoveApps": true,
            "allowCreateUpdateRemoveTabs": true,
            "allowCreateUpdateRemoveConnectors": true
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900" } });
    assert(loggerLogSpy.calledWith({
      "allowCreateUpdateChannels": true,
      "allowDeleteChannels": true,
      "allowAddRemoveApps": true,
      "allowCreateUpdateRemoveTabs": true,
      "allowCreateUpdateRemoveConnectors": true
    }));
  });

  it('lists member settings for a Microsoft Team (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=memberSettings`) {
        return {
          "memberSettings": {
            "allowCreateUpdateChannels": true,
            "allowDeleteChannels": true,
            "allowAddRemoveApps": true,
            "allowCreateUpdateRemoveTabs": true,
            "allowCreateUpdateRemoveConnectors": true
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", debug: true } });
    assert(loggerLogSpy.calledWith({
      "allowCreateUpdateChannels": true,
      "allowDeleteChannels": true,
      "allowAddRemoveApps": true,
      "allowCreateUpdateRemoveTabs": true,
      "allowCreateUpdateRemoveConnectors": true
    }));
  });

  it('fails validation if teamId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when teamId is valid', async () => {
    const actual = await command.validate({
      options: {
        teamId: '2609af39-7775-4f94-a3dc-0dd67657e900'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('lists all properties for output json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=memberSettings`) {
        return {
          "memberSettings": {
            "allowCreateUpdateChannels": true,
            "allowDeleteChannels": true,
            "allowAddRemoveApps": true,
            "allowCreateUpdateRemoveTabs": true,
            "allowCreateUpdateRemoveConnectors": true
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", output: 'json' } });
    assert(loggerLogSpy.calledWith({
      "allowCreateUpdateChannels": true,
      "allowDeleteChannels": true,
      "allowAddRemoveApps": true,
      "allowCreateUpdateRemoveTabs": true,
      "allowCreateUpdateRemoveConnectors": true
    }));
  });

  it('correctly handles error when listing member settings', async () => {
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

    await assert.rejects(command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900" } } as any), new CommandError('An error has occurred'));
  });
});
