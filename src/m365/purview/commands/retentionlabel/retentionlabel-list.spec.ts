import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./retentionlabel-list');

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
  //##endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
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
    assert.strictEqual(command.name.startsWith(commands.RETENTIONLABEL_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'isInUse']);
  });

  it('retrieves retention labels', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels`) {
        return Promise.resolve(mockResponse);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false } });
    assert(loggerLogSpy.calledWith(mockResponseArray));
  });

  it('handles error when retrieving retention labels', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels`) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: false } }), new CommandError('An error has occurred'));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});