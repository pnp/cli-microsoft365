import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './retentionlabel-list.js';

describe(commands.RETENTIONLABEL_LIST, () => {

  //#region Mocked responses
  const mockResponseArray = [
    {
      "displayName": "Some label",
      "descriptionForAdmins": "",
      "descriptionForUsers": null,
      "isInUse": true,
      "retentionTrigger": "dateCreated",
      "behaviorDuringRetentionPeriod": "retainAsRecord",
      "actionAfterRetentionPeriod": "delete",
      "createdDateTime": "2022-11-03T10:28:15Z",
      "lastModifiedDateTime": "2022-11-03T10:28:15Z",
      "labelToBeApplied": null,
      "defaultRecordBehavior": "startLocked",
      "id": "dc67203a-6cca-4066-b501-903401308f98",
      "retentionDuration": {
        "days": 365
      },
      "createdBy": {
        "user": {
          "id": "b52ffd35-d6fe-4b70-86d8-91cc01d76333",
          "displayName": null
        }
      },
      "lastModifiedBy": {
        "user": {
          "id": "b52ffd35-d6fe-4b70-86d8-91cc01d76333",
          "displayName": null
        }
      },
      "dispositionReviewStages": []
    }
  ];

  const mockResponse = {
    "@odata.context": "https://graph.microsoft.com/beta/$metadata#security/labels/retentionLabels",
    "@odata.count": 2,
    "value": mockResponseArray
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
    assert.strictEqual(command.name, commands.RETENTIONLABEL_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'isInUse']);
  });

  it('retrieves retention labels', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels`) {
        return mockResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(mockResponseArray));
  });

  it('handles error when retrieving retention labels', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels`) {
        throw { error: { error: { message: 'An error has occurred' } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('An error has occurred'));
  });
});