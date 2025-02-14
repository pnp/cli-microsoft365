import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './retentionlabel-get.js';

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
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
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
});