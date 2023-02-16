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
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
const command: Command = require('./threatassessment-list');

describe(commands.THREATASSESSMENT_LIST, () => {
  //#region Mocked Responses
  let commandInfo: CommandInfo;
  const threatAssesmentResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#informationProtection/threatAssessmentRequests",
    "value": [
      {
        "@odata.type": "#microsoft.graph.mailAssessmentRequest",
        "id": "49c5ef5b-1f65-444a-e6b9-08d772ea2059",
        "createdDateTime": "2019-11-27T03:30:18.6890937Z",
        "contentType": "mail",
        "expectedAssessment": "block",
        "category": "spam",
        "status": "pending",
        "requestSource": "administrator",
        "recipientEmail": "tifc@a830edad9050849eqtpwbjzxodq.onmicrosoft.com",
        "destinationRoutingReason": "notJunk",
        "messageUri": "https://graph.microsoft.com/v1.0/users/c52ce8db-3e4b-4181-93c4-7d6b6bffaf60/messages/AAMkADU3MWUxOTU0LWNlOTEt=",
        "createdBy": {
          "user": {
            "id": "c52ce8db-3e4b-4181-93c4-7d6b6bffaf60",
            "displayName": "Ronald Admin"
          }
        }
      },
      {
        "@odata.type": "#microsoft.graph.emailFileAssessmentRequest",
        "id": "ab2ad9b3-2213-4091-ae0c-08d76ddbcacf",
        "createdDateTime": "2019-11-20T17:05:06.4088076Z",
        "contentType": "mail",
        "expectedAssessment": "block",
        "category": "malware",
        "status": "completed",
        "requestSource": "administrator",
        "recipientEmail": "tifc@a830edad9050849EQTPWBJZXODQ.onmicrosoft.com",
        "destinationRoutingReason": "notJunk",
        "contentData": "",
        "createdBy": {
          "user": {
            "id": "c52ce8db-3e4b-4181-93c4-7d6b6bffaf60",
            "displayName": "Ronald Admin"
          }
        }
      },
      {
        "@odata.type": "#microsoft.graph.fileAssessmentRequest",
        "id": "18406a56-7209-4720-a250-08d772fccdaa",
        "createdDateTime": "2019-11-27T05:44:00.4051536Z",
        "contentType": "file",
        "expectedAssessment": "block",
        "category": "malware",
        "status": "completed",
        "requestSource": "administrator",
        "fileName": "b3d5b715-4b88-4bbb-b0ae-9a9281a3f18a.csv",
        "contentData": "",
        "createdBy": {
          "user": {
            "id": "c52ce8db-3e4b-4181-93c4-7d6b6bffaf60",
            "displayName": "Ronald Admin"
          }
        }
      },
      {
        "@odata.type": "#microsoft.graph.urlAssessmentRequest",
        "id": "723c35be-8b5a-47ae-29c0-08d76ddb7f5b",
        "createdDateTime": "2019-11-20T17:02:59.8160832Z",
        "contentType": "url",
        "expectedAssessment": "unblock",
        "category": "phishing",
        "status": "completed",
        "requestSource": "administrator",
        "url": "http://test.com",
        "createdBy": {
          "user": {
            "id": "c52ce8db-3e4b-4181-93c4-7d6b6bffaf60",
            "displayName": "Ronald Admin"
          }
        }
      }
    ]
  };

  const threatAssesmentMailResponse: any = [{
    "@odata.type": "#microsoft.graph.mailAssessmentRequest",
    "id": "49c5ef5b-1f65-444a-e6b9-08d772ea2059",
    "createdDateTime": "2019-11-27T03:30:18.6890937Z",
    "contentType": "mail",
    "expectedAssessment": "block",
    "category": "spam",
    "status": "pending",
    "requestSource": "administrator",
    "recipientEmail": "tifc@a830edad9050849eqtpwbjzxodq.onmicrosoft.com",
    "destinationRoutingReason": "notJunk",
    "messageUri": "https://graph.microsoft.com/v1.0/users/c52ce8db-3e4b-4181-93c4-7d6b6bffaf60/messages/AAMkADU3MWUxOTU0LWNlOTEt=",
    "createdBy": {
      "user": {
        "id": "c52ce8db-3e4b-4181-93c4-7d6b6bffaf60",
        "displayName": "Ronald Admin"
      }
    }
  }];

  const threatAssesmentEmailFileResponse: any = [{
    "@odata.type": "#microsoft.graph.emailFileAssessmentRequest",
    "id": "ab2ad9b3-2213-4091-ae0c-08d76ddbcacf",
    "createdDateTime": "2019-11-20T17:05:06.4088076Z",
    "contentType": "mail",
    "expectedAssessment": "block",
    "category": "malware",
    "status": "completed",
    "requestSource": "administrator",
    "recipientEmail": "tifc@a830edad9050849EQTPWBJZXODQ.onmicrosoft.com",
    "destinationRoutingReason": "notJunk",
    "contentData": "",
    "createdBy": {
      "user": {
        "id": "c52ce8db-3e4b-4181-93c4-7d6b6bffaf60",
        "displayName": "Ronald Admin"
      }
    }
  }];

  const threatAssesmentFileResponse: any = [{
    "@odata.type": "#microsoft.graph.fileAssessmentRequest",
    "id": "18406a56-7209-4720-a250-08d772fccdaa",
    "createdDateTime": "2019-11-27T05:44:00.4051536Z",
    "contentType": "file",
    "expectedAssessment": "block",
    "category": "malware",
    "status": "completed",
    "requestSource": "administrator",
    "fileName": "b3d5b715-4b88-4bbb-b0ae-9a9281a3f18a.csv",
    "contentData": "",
    "createdBy": {
      "user": {
        "id": "c52ce8db-3e4b-4181-93c4-7d6b6bffaf60",
        "displayName": "Ronald Admin"
      }
    }
  }];

  const threatAssesmentUrlResponse: any = [{
    "@odata.type": "#microsoft.graph.urlAssessmentRequest",
    "id": "723c35be-8b5a-47ae-29c0-08d76ddb7f5b",
    "createdDateTime": "2019-11-20T17:02:59.8160832Z",
    "contentType": "url",
    "expectedAssessment": "unblock",
    "category": "phishing",
    "status": "completed",
    "requestSource": "administrator",
    "url": "http://test.com",
    "createdBy": {
      "user": {
        "id": "c52ce8db-3e4b-4181-93c4-7d6b6bffaf60",
        "displayName": "Ronald Admin"
      }
    }
  }];
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    commandInfo = Cli.getCommandInfo(command);
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
    assert.strictEqual(command.name, commands.THREATASSESSMENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'contentType', 'category']);
  });

  it('fails validation if specified type is invalid ', async () => {
    const actual = await command.validate({ options: { type: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if type option correctly specified', async () => {
    const actual = await command.validate({ options: { type: 'file' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves threat assesments', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests`)) {
        return threatAssesmentResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(threatAssesmentResponse.value));
  });

  it('retrieves threat assesments with type mail', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests`)) {
        return threatAssesmentResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, type: 'mail' } });
    assert(loggerLogSpy.calledWith(threatAssesmentMailResponse));
  });

  it('retrieves threat assesments with type emailFile', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests`)) {
        return threatAssesmentResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, type: 'emailFile' } });
    assert(loggerLogSpy.calledWith(threatAssesmentEmailFileResponse));
  });

  it('retrieves threat assesments with type File', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests`)) {
        return threatAssesmentResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, type: 'file' } });
    assert(loggerLogSpy.calledWith(threatAssesmentFileResponse));
  });

  it('retrieves threat assesments with type Url', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests`)) {
        return threatAssesmentResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, type: 'url' } });
    assert(loggerLogSpy.calledWith(threatAssesmentUrlResponse));
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The threat assesments could not be retrieved'
      }
    };
    sinon.stub(request, 'get').callsFake(async () => { throw error; });

    await assert.rejects(command.action(logger, {
      options: {}
    }), new CommandError('The threat assesments could not be retrieved'));
  });
});
