import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './run-list.js';

describe(commands.RUN_LIST, () => {
  const environmentName = 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c';
  const flowName = '8c3f591b-5054-4cad-9cf1-832104ec0290';
  const status = 'Running';
  const triggerStartTime = '2023-01-21T18:19:00Z';
  const triggerEndTime = '2023-01-22T00:00:00Z';
  const flowRunListResponse = {
    value: [
      {
        "name": "08586653536760200319026785874CU62",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/396d5ec9-ae2d-4a84-967d-cd7f56cd8f30/runs/08586653536760200319026785874CU62",
        "type": "Microsoft.ProcessSimple/environments/flows/runs",
        "properties": {
          "startTime": "2018-09-06T17:00:09.9484194Z",
          "endTime": "2018-09-06T17:00:10.3406851Z",
          "status": "Succeeded",
          "correlation": {
            "clientTrackingId": "08586653536760200320026785874CU62"
          },
          "trigger": {
            "name": "When_a_file_is_created_or_modified_(properties_only)",
            "inputsLink": {
              "uri": "https://prod-59.westeurope.logic.azure.com:443/workflows/2d8d4d3c94604eeeadc68464ea5fb361/runs/08586653536760200319026785874CU62/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-06T21%3A00%3A00.0000000Z&sp=%2Fruns%2F08586653536760200319026785874CU62%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=75F6WNUyKVJXcdQJIra9jF6X_kac12GSlFHX3NY_X_U",
              "contentVersion": "98GuGIhrxUoG/lKXcXUgaA==",
              "contentSize": 515,
              "contentHash": {
                "algorithm": "md5",
                "value": "98GuGIhrxUoG/lKXcXUgaA=="
              }
            },
            "outputsLink": {
              "uri": "https://prod-59.westeurope.logic.azure.com:443/workflows/2d8d4d3c94604eeeadc68464ea5fb361/runs/08586653536760200319026785874CU62/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-06T21%3A00%3A00.0000000Z&sp=%2Fruns%2F08586653536760200319026785874CU62%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=CJrx9-PIyK8Vk_V7YdY-HV4zxcL2i6rjbXOXKPIOegk",
              "contentVersion": "KNpZY3gib8WXg6/bxuIsSA==",
              "contentSize": 3661,
              "contentHash": {
                "algorithm": "md5",
                "value": "KNpZY3gib8WXg6/bxuIsSA=="
              }
            },
            "startTime": "2018-09-06T17:00:09.4562613Z",
            "endTime": "2018-09-06T17:00:09.7844035Z",
            "scheduledTime": "2018-09-06T17:00:09.8558878Z",
            "correlation": {
              "clientTrackingId": "08586653536760200320026785874CU62"
            },
            "code": "OK",
            "status": "Succeeded"
          }
        }
      },
      {
        "name": "08586653539691313445320015404CU49",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/396d5ec9-ae2d-4a84-967d-cd7f56cd8f30/runs/08586653539691313445320015404CU49",
        "type": "Microsoft.ProcessSimple/environments/flows/runs",
        "properties": {
          "startTime": "2018-09-06T16:55:16.8922841Z",
          "endTime": "2018-09-06T16:55:17.1607417Z",
          "status": "Succeeded",
          "correlation": {
            "clientTrackingId": "08586653539691313446320015404CU29"
          },
          "trigger": {
            "name": "When_a_file_is_created_or_modified_(properties_only)",
            "inputsLink": {
              "uri": "https://prod-59.westeurope.logic.azure.com:443/workflows/2d8d4d3c94604eeeadc68464ea5fb361/runs/08586653539691313445320015404CU49/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-06T21%3A00%3A00.0000000Z&sp=%2Fruns%2F08586653539691313445320015404CU49%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=fke3vk-ABOiv-Msq-f4Pw_7ozMovk1VHmbz40P998c4",
              "contentVersion": "98GuGIhrxUoG/lKXcXUgaA==",
              "contentSize": 515,
              "contentHash": {
                "algorithm": "md5",
                "value": "98GuGIhrxUoG/lKXcXUgaA=="
              }
            },
            "outputsLink": {
              "uri": "https://prod-59.westeurope.logic.azure.com:443/workflows/2d8d4d3c94604eeeadc68464ea5fb361/runs/08586653539691313445320015404CU49/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-06T21%3A00%3A00.0000000Z&sp=%2Fruns%2F08586653539691313445320015404CU49%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=0TTEb1p5HXyLJUeMmr4iR3kyhxFStuA2ILQFQQmViqk",
              "contentVersion": "db9U8YauD8oO58o4VVtJmA==",
              "contentSize": 3680,
              "contentHash": {
                "algorithm": "md5",
                "value": "db9U8YauD8oO58o4VVtJmA=="
              }
            },
            "startTime": "2018-09-06T16:55:16.3365001Z",
            "endTime": "2018-09-06T16:55:16.6646378Z",
            "scheduledTime": "2018-09-06T16:55:15.8797016Z",
            "correlation": {
              "clientTrackingId": "08586653539691313446320015404CU29"
            },
            "code": "OK",
            "status": "Succeeded"
          }
        }
      }
    ]
  };
  const triggerInformationResponse = { 'headers': { 'Pragma': 'no-cache', 'Transfer-Encoding': 'chunked', 'Retry-After': '3600', 'Vary': 'Accept-Encoding', 'Cache-Control': 'no-store, no-cache', 'Location': 'https://flow-apim-europe-002-westeurope-01.azure-apim.net/apim/office365/shared-office365-13750b71-ac69-403f-ad09-23e056037806/v4/Mail/OnFlaggedEmail?folderPath=Inbox&importance=Any&fetchOnlyWithAttachment=false&includeAttachments=false&LastPollInformation=eyJMYXN0UmVjZWl2ZWRNYWlsVGltZSI6IjIwMjItMTEtMTVUMTA6MTM6NTUuMDYyNjQ0MyswMDowMCIsIkxhc3RDcmVhdGVkTWFpbFRpbWUiOiIyMDIzLTAzLTA0VDA5OjA1OjIxKzAwOjAwIiwiTGFzdE1lc3NhZ2VJZCI6IkFBTWtBRGd6TjJRMU5UaGlMVEkwTmpZdE5HSXhZUzA1TURkakxUZzFPV1F4Tnpnd1pHTTJaZ0JHQUFBQUFBQzZqUWZVemFjVFNJSHFNdzJ5YWNuVUJ3QmlPQzh4dlltZFQ2RzJFX2hMTUs1a0FBQUFBQUVNQUFCaU9DOHh2WW1kVDZHMkVfaExNSzVrQUFMVXF5ODFBQUE9IiwiTGFzdEludGVybmV0TWVzc2FnZUlkIjoiPERCN1BSMDNNQjUwMTg4Nzk5MTQzMjRGQzY1Njk1ODA5RkUxQUQ5QERCN1BSMDNNQjUwMTguZXVycHJkMDMucHJvZC5vdXRsb29rLmNvbT4ifQ%3d%3d', 'Set-Cookie': '', 'x-ms-request-id': 'ca65b907-e292-41ff-b76e-f3fe60e87c8c;2192534a-f34f-4920-88c9-72b1b1bcf4b4', 'Strict-Transport-Security': 'max-age=31536000; includeSubDomains', 'X-Content-Type-Options': 'nosniff', 'X-Frame-Options': 'DENY', 'Timing-Allow-Origin': '*', 'x-ms-apihub-cached-response': 'true', 'x-ms-apihub-obo': 'false', 'Date': 'Sat, 04 Mar 2023 09:05:21 GMT', 'Content-Type': 'application/json', 'Expires': '-1', 'Content-Length': '1831' }, 'body': { 'from': 'john@contoso.com', 'toRecipients': 'doe@contoso.com', 'subject': 'This is a test mail', 'body': '<html><head>\r\n<meta http-equiv=\'Content- Type\' content=\'text / html; charset=utf - 8\'></head><body><p>This is dummy content</p></body></html>', 'importance': 'normal', 'bodyPreview': 'This is dummy content', 'hasAttachments': false, 'id': 'AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E_hLMK5kAAAAAAEMAABiOC8xvYmdT6G2E_hLMK5kAALUqy81AAA=', 'internetMessageId': '<DB7PR03MB5018879914324FC65695809FE1AD9@DB7PR03MB5018.eurprd03.prod.outlook.com>', 'conversationId': 'AAQkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgAQAMqP9zsK8a1CnIYEgHclLTk=', 'receivedDateTime': '2023-03-01T15:06:57+00:00', 'isRead': false, 'attachments': [], 'isHtml': true } };

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
    assert.strictEqual(command.name, commands.RUN_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'startTime', 'status']);
  });

  it('retrieves all runs for a specific flow', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs?api-version=2016-11-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return flowRunListResponse;
        }
      }

      throw 'Invalid request ' + opts.url;
    });

    await command.action(logger, { options: { environmentName: environmentName, flowName: flowName, verbose: true } });
    assert(loggerLogSpy.calledWith(flowRunListResponse.value));
  });

  it('retrieves all runs for a specific flow as admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${environmentName}/flows/${flowName}/runs?api-version=2016-11-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return flowRunListResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentName, flowName: flowName, asAdmin: true, verbose: true } });
    assert(loggerLogSpy.calledWith(flowRunListResponse.value));
  });

  it('retrieves all runs with a specific status for a specific flow', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs?api-version=2016-11-01&$filter=status eq '${status}'`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return flowRunListResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentName, flowName: flowName, status: status, verbose: true } });
    assert(loggerLogSpy.calledWith(flowRunListResponse.value));
  });

  it('retrieves all runs between two dates for a specific flow', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs?api-version=2016-11-01&$filter=startTime ge ${triggerStartTime} and startTime lt ${triggerEndTime}`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return flowRunListResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentName, flowName: flowName, triggerStartTime: triggerStartTime, triggerEndTime: triggerEndTime, verbose: true } });
    assert(loggerLogSpy.calledWith(flowRunListResponse.value));
  });

  it('retrieves all runs for a specific flow and includes trigger information', async () => {
    const flowRunListResponseClone = { ...flowRunListResponse };
    (flowRunListResponseClone.value[0] as any).triggerInformation = triggerInformationResponse.body;
    (flowRunListResponseClone.value[1] as any).triggerInformation = triggerInformationResponse.body;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs?api-version=2016-11-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return flowRunListResponseClone;
        }
      }

      if (opts.url === flowRunListResponseClone.value[0].properties.trigger.outputsLink.uri) {
        return triggerInformationResponse;
      }

      if (opts.url === flowRunListResponseClone.value[1].properties.trigger.outputsLink.uri) {
        return triggerInformationResponse;
      }

      throw 'Invalid request ' + opts.url;
    });

    await command.action(logger, { options: { environmentName: environmentName, flowName: flowName, withTrigger: true, verbose: true, output: 'json' } });
    assert(loggerLogSpy.calledWith(flowRunListResponseClone.value));
  });

  it('correctly handles no environment found', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "EnvironmentAccessDenied",
        "message": `Access to the environment '${environmentName}' is denied.`
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName, flowName: flowName } } as any),
      new CommandError(`Access to the environment '${environmentName}' is denied.`));
  });

  it('correctly handles no runs for this flow found', async () => {
    sinon.stub(request, 'get').resolves({ value: [] });

    await command.action(logger, { options: { verbose: true, environmentName: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', flowName: '16c90c26-25e0-4800-8af9-da594e02d427' } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName, flowName: flowName } } as any),
      new CommandError('An error has occurred'));
  });

  it('fails validation if the flowName is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the status is not a valid status', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, status: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the triggerStartTime is not a valid ISO datetime', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, triggerStartTime: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the triggerEndTime is not a valid ISO datetime', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, triggerEndTime: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the output is not json and withTrigger is specified', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, output: 'text', withTrigger: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all options are passed properly', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, status: status, triggerStartTime: triggerStartTime, triggerEndTime: triggerEndTime } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
