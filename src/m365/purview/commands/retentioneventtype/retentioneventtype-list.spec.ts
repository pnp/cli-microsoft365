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
const command: Command = require('./retentioneventtype-list');

describe(commands.RETENTIONEVENTTYPE_LIST, () => {

  //#region Mocked responses
  const mockResponseArray = [
    {
      "displayName": "Retention Event Type",
      "description": "",
      "createdDateTime": "2023-02-02T15:47:54Z",
      "lastModifiedDateTime": "2023-02-02T15:47:54Z",
      "id": "81fa91bd-66cd-4c6c-b0cb-71f37210dc74",
      "createdBy": {
        "user": {
          "id": "36155f4e-bdbd-4101-ba20-5e78f5fba9a9",
          "displayName": null
        }
      },
      "lastModifiedBy": {
        "user": {
          "id": "36155f4e-bdbd-4101-ba20-5e78f5fba9a9",
          "displayName": null
        }
      }
    }
  ];

  const mockResponse = {
    "@odata.context": "https://graph.microsoft.com/beta/$metadata#security/triggerTypes/retentionEventTypes",
    "@odata.count": 1,
    "value": mockResponseArray
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONEVENTTYPE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'createdDateTime']);
  });

  it('retrieves retention event types', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes`) {
        return mockResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(mockResponseArray));
  });

  it('handles error when retrieving event types', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes`) {
        throw { error: { error: { message: 'An error has occurred' } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('An error has occurred'));
  });
});
