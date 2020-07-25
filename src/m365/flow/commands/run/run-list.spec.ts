import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./run-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FLOW_RUN_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FLOW_RUN_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves runs for specific flow (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/396d5ec9-ae2d-4a84-967d-cd7f56cd8f30/runs?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
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
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', flow: "396d5ec9-ae2d-4a84-967d-cd7f56cd8f30" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: '08586653536760200319026785874CU62',
            startTime: '2018-09-06T17:00:09.9484194Z',
            status: 'Succeeded'
          },
          {
            name: '08586653539691313445320015404CU49',
            startTime: '2018-09-06T16:55:16.8922841Z',
            status: 'Succeeded'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves runs for specific flow', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/396d5ec9-ae2d-4a84-967d-cd7f56cd8f30/runs?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
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
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', flow: "396d5ec9-ae2d-4a84-967d-cd7f56cd8f30" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: '08586653536760200319026785874CU62',
            startTime: '2018-09-06T17:00:09.9484194Z',
            status: 'Succeeded'
          },
          {
            name: '08586653539691313445320015404CU49',
            startTime: '2018-09-06T16:55:16.8922841Z',
            status: 'Succeeded'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs all properties when output is JSON', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/396d5ec9-ae2d-4a84-967d-cd7f56cd8f30/runs?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
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
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', flow: "396d5ec9-ae2d-4a84-967d-cd7f56cd8f30", output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
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
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no environment found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', flow: "396d5ec9-ae2d-4a84-967d-cd7f56cd8f30" } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no runs for this flow found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.resolve({ value: [] });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', flow: '16c90c26-25e0-4800-8af9-da594e02d427' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no runs for this flow found (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.resolve({ value: [] });
    });

    cmdInstance.action({ options: { debug: true, environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', flow: '16c90c26-25e0-4800-8af9-da594e02d427' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('No runs found'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying environment parameter', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--environment') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying flow parameter', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--flow') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});