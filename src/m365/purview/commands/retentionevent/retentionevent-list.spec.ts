import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './retentionevent-list.js';

describe(commands.RETENTIONEVENTTYPE_GET, () => {

  //#region Mocked responses
  const mockResponseArray = [
    {
      "displayName": "Retention Event",
      "description": null,
      "eventTriggerDateTime": "2023-02-03T13:51:40Z",
      "eventStatus": null,
      "lastStatusUpdateDateTime": null,
      "createdDateTime": "2023-02-03T13:51:40Z",
      "lastModifiedDateTime": "2023-02-03T13:51:40Z",
      "id": "7248cfa8-c03a-4ec1-49a4-08db05edc686",
      "eventQueries": [],
      "eventPropagationResults": [],
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
    }
  ];

  const mockResponse = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#security/triggers/retentionEvents",
    "@odata.count": 2,
    "value": mockResponseArray
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
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
    auth.service.active = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONEVENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'eventTriggerDateTime']);
  });

  it('retrieves retention events', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents`) {
        return mockResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(mockResponseArray));
  });

  it('handles error when retrieving retention events', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents`) {
        throw { error: { error: { message: 'An error has occurred' } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('An error has occurred'));
  });
});
