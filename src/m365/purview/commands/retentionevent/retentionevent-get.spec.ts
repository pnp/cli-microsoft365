import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { accessToken } from '../../../../utils/accessToken';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
const command: Command = require('./retentionevent-get');

describe(commands.RETENTIONEVENT_GET, () => {
  const retentionEventId = 'c37d695e-d581-4ae9-82a0-9364eba4291e';
  const retentionEventGetResponse = {
    "displayName": "Employee Termination",
    "description": "This event occurs when an employee is terminated.",
    "eventTriggerDateTime": "2023-02-01T09:16:37Z",
    "lastStatusUpdateDateTime": "2023-02-01T09:21:15Z",
    "createdDateTime": "2023-02-01T09:17:40Z",
    "lastModifiedDateTime": "2023-02-01T09:17:40Z",
    "id": retentionEventId,
    "eventQueries": [
      {
        "queryType": "files",
        "query": "1234"
      },
      {
        "queryType": "messages",
        "query": "Terminate"
      }
    ],
    "eventStatus": {
      "error": null,
      "status": "success"
    },
    "eventPropagationResults": [
      {
        "serviceName": "SharePoint",
        "location": null,
        "status": "none",
        "statusInformation": null
      }
    ],
    "createdBy": {
      "user": {
        "id": null,
        "displayName": "John Doe"
      }
    },
    "lastModifiedBy": {
      "user": {
        "id": null,
        "displayName": "John Doe"
      }
    }
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => false);
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.get
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
    assert.strictEqual(command.name, commands.RETENTIONEVENT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if a correct id is entered', async () => {
    const actual = await command.validate({ options: { id: retentionEventId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves retention event by specified id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggers/retentionEvents/${retentionEventId}`) {
        return retentionEventGetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: retentionEventId, verbose: true } });
    assert(loggerLogSpy.calledWith(retentionEventGetResponse));
  });

  it('handles error when retention event by specified id is not found', async () => {
    const errorMessage = `Error: The operation couldn't be performed because object '${retentionEventId}' couldn't be found on 'FfoConfigurationSession'.`;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggers/retentionEvents/${retentionEventId}`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: retentionEventId
      }
    }), new CommandError(errorMessage));
  });

  it('throws error if something fails using application permissions', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => true);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`This command currently does not support app only permissions.`));
  });
});