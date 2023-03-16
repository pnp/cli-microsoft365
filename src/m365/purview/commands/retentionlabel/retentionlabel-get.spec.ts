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
const command: Command = require('./retentionlabel-get');

describe(commands.RETENTIONLABEL_GET, () => {

  const retentionLabelId = '5c8af2e2-b489-4fa0-9c16-180180245ac8';
  const retentionLabelGetResponse = {
    "displayName": "TEST LABEL",
    "descriptionForAdmins": "",
    "descriptionForUsers": "",
    "isInUse": false,
    "retentionTrigger": "dateCreated",
    "behaviorDuringRetentionPeriod": "retain",
    "actionAfterRetentionPeriod": "delete",
    "createdDateTime": "2022-12-12T15:14:53Z",
    "lastModifiedDateTime": "2022-12-12T15:43:06Z",
    "labelToBeApplied": "",
    "defaultRecordBehavior": "startLocked",
    "id": retentionLabelId,
    "retentionDuration": {
      "days": 100
    },
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
    },
    "dispositionReviewStages": []
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
    assert.strictEqual(command.name, commands.RETENTIONLABEL_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves retention label specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${retentionLabelId}`) {
        return retentionLabelGetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: retentionLabelId, verbose: true } });
    assert(loggerLogSpy.calledWith(retentionLabelGetResponse));
  });

  it('handles error when retention label by id is not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${retentionLabelId}`) {
        throw `Error: The operation couldn't be performed because object '${retentionLabelId}' couldn't be found on 'FfoConfigurationSession'.`;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: retentionLabelId } }), new CommandError(`Error: The operation couldn't be performed because object '${retentionLabelId}' couldn't be found on 'FfoConfigurationSession'.`));
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if a correct id is entered', async () => {
    const actual = await command.validate({ options: { id: retentionLabelId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('throws an error when we execute the command using application permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    await assert.rejects(command.action(logger, { options: { id: retentionLabelId } }),
      new CommandError('This command does not support application permissions.'));
  });
});