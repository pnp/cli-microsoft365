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
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './threatassessment-list.js';

describe(commands.THREATASSESSMENT_LIST, () => {
  //#region Mocked Responses
  const threatAssessmentMailItem: any = {
    "type": "mail",
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
  };

  const threatAssessmentEmailFileItem: any = {
    "type": "emailFile",
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
  };

  const threatAssessmentFileItem: any = {
    "type": "file",
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
  };

  const threatAssessmentUrlItem: any = {
    "type": "url",
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
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.connection.active = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.THREATASSESSMENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'type', 'category']);
  });

  it('fails validation if specified type is invalid ', async () => {
    const actual = await command.validate({ options: { type: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if type option correctly specified', async () => {
    const actual = await command.validate({ options: { type: 'file' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves threat assessments', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests`)) {
        return {
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
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly([threatAssessmentMailItem, threatAssessmentFileItem]));
  });

  it('retrieves threat assessments with type mail', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests?$filter=contentType eq 'mail'`)) {
        return {
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
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'mail' } });
    assert(loggerLogSpy.calledOnceWithExactly([threatAssessmentMailItem]));
  });

  it('retrieves threat assessments with type emailFile', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests?$filter=contentType eq 'mail'`)) {
        return {
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
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'emailFile' } });
    assert(loggerLogSpy.calledOnceWithExactly([threatAssessmentEmailFileItem]));
  });

  it('retrieves threat assessments with type File', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests?$filter=contentType eq 'file'`)) {
        return {
          "value": [
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
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'file' } });
    assert(loggerLogSpy.calledOnceWithExactly([threatAssessmentFileItem]));
  });

  it('retrieves threat assessments with type Url', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests?$filter=contentType eq 'url'`)) {
        return {
          "value": [
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
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'url' } });
    assert(loggerLogSpy.calledOnceWithExactly([threatAssessmentUrlItem]));
  });

  it('retrieves threat assessments with unknown future value', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests?$filter=contentType eq 'url'`)) {
        return {
          "value": [
            {
              "@odata.type": "#microsoft.graph.unknownFutureValueAssessmentRequest",
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
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'url' } });
    assert(loggerLogSpy.calledOnceWithExactly([{ ...threatAssessmentUrlItem, type: 'Unknown' }]));
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The threat assessments could not be retrieved'
      }
    };
    sinon.stub(request, 'get').callsFake(async () => { throw error; });

    await assert.rejects(command.action(logger, {
      options: {}
    }), new CommandError('The threat assessments could not be retrieved'));
  });
});