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
const command: Command = require('./run-get');

describe(commands.RUN_GET, () => {
  const flowName = '396d5ec9-ae2d-4a84-967d-cd7f56cd8f30';
  const environmentName = 'Default-48595cc3-adce-4267-8e99-0c838923dbb9';
  const runName = '08586653536760200319026785874CU62';

  //#region Mocked Responses flow run get
  const flowResponse = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '2018-09-07T19:33:31.3640166Z', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '2018-09-07T19:23:31.3482269Z', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } } };
  const flowResponseNoEndTime = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } } };
  const triggerInformationResponse = { 'headers': { 'Pragma': 'no-cache', 'Transfer-Encoding': 'chunked', 'Retry-After': '3600', 'Vary': 'Accept-Encoding', 'Cache-Control': 'no-store, no-cache', 'Location': 'https://flow-apim-europe-002-westeurope-01.azure-apim.net/apim/office365/shared-office365-13750b71-ac69-403f-ad09-23e056037806/v4/Mail/OnFlaggedEmail?folderPath=Inbox&importance=Any&fetchOnlyWithAttachment=false&includeAttachments=false&LastPollInformation=eyJMYXN0UmVjZWl2ZWRNYWlsVGltZSI6IjIwMjItMTEtMTVUMTA6MTM6NTUuMDYyNjQ0MyswMDowMCIsIkxhc3RDcmVhdGVkTWFpbFRpbWUiOiIyMDIzLTAzLTA0VDA5OjA1OjIxKzAwOjAwIiwiTGFzdE1lc3NhZ2VJZCI6IkFBTWtBRGd6TjJRMU5UaGlMVEkwTmpZdE5HSXhZUzA1TURkakxUZzFPV1F4Tnpnd1pHTTJaZ0JHQUFBQUFBQzZqUWZVemFjVFNJSHFNdzJ5YWNuVUJ3QmlPQzh4dlltZFQ2RzJFX2hMTUs1a0FBQUFBQUVNQUFCaU9DOHh2WW1kVDZHMkVfaExNSzVrQUFMVXF5ODFBQUE9IiwiTGFzdEludGVybmV0TWVzc2FnZUlkIjoiPERCN1BSMDNNQjUwMTg4Nzk5MTQzMjRGQzY1Njk1ODA5RkUxQUQ5QERCN1BSMDNNQjUwMTguZXVycHJkMDMucHJvZC5vdXRsb29rLmNvbT4ifQ%3d%3d', 'Set-Cookie': '', 'x-ms-request-id': 'ca65b907-e292-41ff-b76e-f3fe60e87c8c;2192534a-f34f-4920-88c9-72b1b1bcf4b4', 'Strict-Transport-Security': 'max-age=31536000; includeSubDomains', 'X-Content-Type-Options': 'nosniff', 'X-Frame-Options': 'DENY', 'Timing-Allow-Origin': '*', 'x-ms-apihub-cached-response': 'true', 'x-ms-apihub-obo': 'false', 'Date': 'Sat, 04 Mar 2023 09:05:21 GMT', 'Content-Type': 'application/json', 'Expires': '-1', 'Content-Length': '1831' }, 'body': { 'from': 'john@contoso.com', 'toRecipients': 'doe@contoso.com', 'subject': 'This is a test mail', 'body': '<html><head>\r\n<meta http-equiv=\'Content- Type\' content=\'text / html; charset=utf - 8\'></head><body><p>This is dummy content</p></body></html>', 'importance': 'normal', 'bodyPreview': 'This is dummy content', 'hasAttachments': false, 'id': 'AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E_hLMK5kAAAAAAEMAABiOC8xvYmdT6G2E_hLMK5kAALUqy81AAA=', 'internetMessageId': '<DB7PR03MB5018879914324FC65695809FE1AD9@DB7PR03MB5018.eurprd03.prod.outlook.com>', 'conversationId': 'AAQkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgAQAMqP9zsK8a1CnIYEgHclLTk=', 'receivedDateTime': '2023-03-01T15:06:57+00:00', 'isRead': false, 'attachments': [], 'isHtml': true } };

  const flowResponseFormatted = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '2018-09-07T19:33:31.3640166Z', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '2018-09-07T19:23:31.3482269Z', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } }, startTime: '2018-09-07T19:23:31.3640166Z', endTime: '2018-09-07T19:33:31.3640166Z', status: 'Running', triggerName: 'manual' };
  const flowResponseFormattedNoEndTime = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } }, startTime: '2018-09-07T19:23:31.3640166Z', endTime: '', status: 'Running', triggerName: 'manual' };
  const flowResponseFormattedIncludingInformation = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '2018-09-07T19:33:31.3640166Z', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '2018-09-07T19:23:31.3482269Z', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } }, startTime: '2018-09-07T19:23:31.3640166Z', endTime: '2018-09-07T19:33:31.3640166Z', status: 'Running', triggerName: 'manual', triggerInformation: triggerInformationResponse.body };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RUN_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified run', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs/${runName}?api-version=2016-11-01`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName, verbose: true } });
    assert(loggerLogSpy.calledWith(flowResponseFormatted));
  });

  it('renders empty string for endTime, if the run specified is still running', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs/${runName}?api-version=2016-11-01`) {
        return flowResponseNoEndTime;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName } });
    assert(loggerLogSpy.calledWith(flowResponseFormattedNoEndTime));
  });

  it('retrieves information about the specified run including trigger information', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs/${runName}?api-version=2016-11-01`) {
        return flowResponse;
      }

      if (opts.url === flowResponse.properties.trigger.outputsLink.uri) {
        return triggerInformationResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName, includeTriggerInformation: true, verbose: true } });
    assert(loggerLogSpy.calledWith(flowResponseFormattedIncludingInformation));
  });

  it('correctly handles Flow not found', async () => {
    sinon.stub(request, 'get').rejects({
      'error': {
        'code': 'FlowNotFound',
        'message': `Could not find flow '${flowName}'.`
      }
    });

    await assert.rejects(command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName } } as any),
      new CommandError(`Could not find flow '${flowName}'.`));
  });

  it('fails validation if the flowName is not valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: 'invalid', name: runName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the flowName is not valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, name: runName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'startTime', 'endTime', 'status', 'triggerName']);
  });
});
