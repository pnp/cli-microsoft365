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
const command: Command = require('./app-remove');

describe(commands.APP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let requests: any[];
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
    requests = [];
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the id is not a valid GUID.', async () => {
    const actual = await command.validate({
      options: { id: 'invalid' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input.', async () => {
    const actual = await command.validate({
      options: {
        id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('remove Teams app in the tenant app catalog with confirmation', async () => {
    let removeTeamsAppCalled = false;
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        removeTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, confirm: true } });
    assert(removeTeamsAppCalled);
  });

  it('remove Teams app in the tenant app catalog with confirmation (debug)', async () => {
    let removeTeamsAppCalled = false;
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        removeTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, confirm: true } });
    assert(removeTeamsAppCalled);
  });

  it('remove Teams app in the tenant app catalog without confirmation', async () => {
    let removeTeamsAppCalled = false;
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        removeTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { debug: true, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } });
    assert(removeTeamsAppCalled);
  });

  it('aborts removing Teams app when prompt not confirmed', async () => {
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

    command.action(logger, { options: { id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } });
    assert(requests.length === 0);
  });

  it('correctly handles error when removing app', async () => {
    sinon.stub(request, 'delete').rejects({
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    });
    await assert.rejects(command.action(logger, {
      options: {
        filePath: 'teamsapp.zip',
        id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, confirm: true
      }
    } as any), new CommandError('An error has occurred'));
  });
});
