import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
import { accessToken } from '../../../../utils/accessToken';
const command: Command = require('./retentioneventtype-get');

describe(commands.RETENTIONEVENTTYPE_GET, () => {
  const retentionEventTypeId = 'c37d695e-d581-4ae9-82a0-9364eba4291e';
  const retentionEventTypeGetResponse = {
    "displayName": "Test retention event type",
    "description": "Description for the retention event type",
    "createdDateTime": "2023-01-29T09:30:42Z",
    "lastModifiedDateTime": "2023-01-29T09:30:42Z",
    "id": "c37d695e-d581-4ae9-82a0-9364eba4291e",
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
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
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
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
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
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONEVENTTYPE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if a correct id is entered', async () => {
    const actual = await command.validate({ options: { id: retentionEventTypeId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves retention event type by specified id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes/${retentionEventTypeId}`) {
        return retentionEventTypeGetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: retentionEventTypeId, verbose: true } });
    assert(loggerLogSpy.calledWith(retentionEventTypeGetResponse));
  });

  it('handles error when retention event type by specified id is not found', async () => {
    const errorMessage = `Error: The operation couldn't be performed because object '${retentionEventTypeId}' couldn't be found on 'FfoConfigurationSession'.`;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes/${retentionEventTypeId}`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: retentionEventTypeId
      }
    }), new CommandError(errorMessage));
  });

  it('throws an error when we execute the command using application permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    await assert.rejects(command.action(logger, { options: { id: retentionEventTypeId } }),
      new CommandError('This command does not support application permissions.'));
  });
});