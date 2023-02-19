import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./run-list');

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
        },
        startTime: '2018-09-06T17:00:09.9484194Z',
        status: 'Succeeded'
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
        },
        startTime: '2018-09-06T16:55:16.8922841Z',
        status: 'Running'
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
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
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.RUN_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'startTime', 'status']);
  });

  it('retrieves all runs for a specific flow', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs?api-version=2016-11-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return flowRunListResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { environmentName: environmentName, flowName: flowName, verbose: true } });
    assert(loggerLogSpy.calledWith(flowRunListResponse.value));
  });

  it('retrieves all runs with a specific status for a specific flow', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs?api-version=2016-11-01&$filter=status eq '${status}'`) {
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
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs?api-version=2016-11-01&$filter=startTime ge ${triggerStartTime} and startTime lt ${triggerEndTime}`) {
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

  it('correctly handles no environment found', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw {
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": `Access to the environment '${environmentName}' is denied.`
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName, flowName: flowName } } as any),
      new CommandError(`Access to the environment '${environmentName}' is denied.`));
  });

  it('correctly handles no runs for this flow found', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      return { value: [] };
    });

    await command.action(logger, { options: { verbose: true, environmentName: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', flowName: '16c90c26-25e0-4800-8af9-da594e02d427' } });
    assert(loggerLogToStderrSpy.calledWith('No runs found'));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw {
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      };
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

  it('passes validation if all options are passed properly', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, status: status, triggerStartTime: triggerStartTime, triggerEndTime: triggerEndTime } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
