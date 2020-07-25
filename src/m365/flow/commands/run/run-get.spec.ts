import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./run-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FLOW_RUN_GET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.FLOW_RUN_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified run (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/396d5ec9-ae2d-4a84-967d-cd7f56cd8f30/runs/08586653536760200319026785874CU62?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
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
                  "uri": "https://prod-59.westeurope.logic.azure.com:443/workflows/2d8d4d3c94604eeeadc68464ea5fb361/runs/08586653536760200319026785874CU62/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T22%3A00%3A00.0000000Z&sp=%2Fruns%2F08586653536760200319026785874CU62%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=K2gG1YUOzIL2XCAiW0m8UDnbF6ECKmDy5sEsdw8EXC0",
                  "contentVersion": "98GuGIhrxUoG/lKXcXUgaA==",
                  "contentSize": 515,
                  "contentHash": {
                    "algorithm": "md5",
                    "value": "98GuGIhrxUoG/lKXcXUgaA=="
                  }
                },
                "outputsLink": {
                  "uri": "https://prod-59.westeurope.logic.azure.com:443/workflows/2d8d4d3c94604eeeadc68464ea5fb361/runs/08586653536760200319026785874CU62/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T22%3A00%3A00.0000000Z&sp=%2Fruns%2F08586653536760200319026785874CU62%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=khJc2zfPe4bHnGU2BnuulbKt2c9FdYF2ZDizn3s8mF8",
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
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, flow: '396d5ec9-ae2d-4a84-967d-cd7f56cd8f30', environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', name: '08586653536760200319026785874CU62' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          name: '08586653536760200319026785874CU62',
          startTime: '2018-09-06T17:00:09.9484194Z',
          endTime: '2018-09-06T17:00:10.3406851Z',
          status: 'Succeeded',
          triggerName: 'When_a_file_is_created_or_modified_(properties_only)'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified run', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/396d5ec9-ae2d-4a84-967d-cd7f56cd8f30/runs/08586653536760200319026785874CU62?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
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
                  "uri": "https://prod-59.westeurope.logic.azure.com:443/workflows/2d8d4d3c94604eeeadc68464ea5fb361/runs/08586653536760200319026785874CU62/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T22%3A00%3A00.0000000Z&sp=%2Fruns%2F08586653536760200319026785874CU62%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=K2gG1YUOzIL2XCAiW0m8UDnbF6ECKmDy5sEsdw8EXC0",
                  "contentVersion": "98GuGIhrxUoG/lKXcXUgaA==",
                  "contentSize": 515,
                  "contentHash": {
                    "algorithm": "md5",
                    "value": "98GuGIhrxUoG/lKXcXUgaA=="
                  }
                },
                "outputsLink": {
                  "uri": "https://prod-59.westeurope.logic.azure.com:443/workflows/2d8d4d3c94604eeeadc68464ea5fb361/runs/08586653536760200319026785874CU62/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T22%3A00%3A00.0000000Z&sp=%2Fruns%2F08586653536760200319026785874CU62%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=khJc2zfPe4bHnGU2BnuulbKt2c9FdYF2ZDizn3s8mF8",
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
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, flow: '396d5ec9-ae2d-4a84-967d-cd7f56cd8f30', environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', name: '08586653536760200319026785874CU62' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          name: '08586653536760200319026785874CU62',
          startTime: '2018-09-06T17:00:09.9484194Z',
          endTime: '2018-09-06T17:00:10.3406851Z',
          status: 'Succeeded',
          triggerName: 'When_a_file_is_created_or_modified_(properties_only)'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns all properties for JSON output', (done) => {
    const runInfo: any = {
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
            "uri": "https://prod-59.westeurope.logic.azure.com:443/workflows/2d8d4d3c94604eeeadc68464ea5fb361/runs/08586653536760200319026785874CU62/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T22%3A00%3A00.0000000Z&sp=%2Fruns%2F08586653536760200319026785874CU62%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=K2gG1YUOzIL2XCAiW0m8UDnbF6ECKmDy5sEsdw8EXC0",
            "contentVersion": "98GuGIhrxUoG/lKXcXUgaA==",
            "contentSize": 515,
            "contentHash": {
              "algorithm": "md5",
              "value": "98GuGIhrxUoG/lKXcXUgaA=="
            }
          },
          "outputsLink": {
            "uri": "https://prod-59.westeurope.logic.azure.com:443/workflows/2d8d4d3c94604eeeadc68464ea5fb361/runs/08586653536760200319026785874CU62/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T22%3A00%3A00.0000000Z&sp=%2Fruns%2F08586653536760200319026785874CU62%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=khJc2zfPe4bHnGU2BnuulbKt2c9FdYF2ZDizn3s8mF8",
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
    };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/396d5ec9-ae2d-4a84-967d-cd7f56cd8f30/runs/08586653536760200319026785874CU62?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve(runInfo);
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, flow: '396d5ec9-ae2d-4a84-967d-cd7f56cd8f30', environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', name: '08586653536760200319026785874CU62', output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(runInfo));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('renders empty string for endTime, if the run specified is still running', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35?api-version=2016-11-01`) > -1) {
        return Promise.resolve({
          "name": "08586652586741142222645090602CU35",
          "id": "/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35",
          "type": "Microsoft.ProcessSimple/environments/flows/runs",
          "properties": {
            "startTime": "2018-09-07T19:23:31.3640166Z",
            "status": "Running",
            "correlation": {
              "clientTrackingId": "08586652586741142222645090602CU35",
              "clientKeywords": ["testFlow"]
            },
            "trigger": {
              "name": "manual",
              "inputsLink": {
                "uri": "https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q",
                "contentVersion": "1UI/8pYQdWDVSsijF+0l2Q==",
                "contentSize": 58,
                "contentHash": {
                  "algorithm": "md5",
                  "value": "1UI/8pYQdWDVSsijF+0l2Q=="
                }
              },
              "outputsLink": {
                "uri": "https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM",
                "contentVersion": "YgV2ecynizFKxzT8yiNtpA==",
                "contentSize": 4244,
                "contentHash": {
                  "algorithm": "md5",
                  "value": "YgV2ecynizFKxzT8yiNtpA=="
                }
              },
              "startTime": "2018-09-07T19:23:31.3482269Z",
              "endTime": "2018-09-07T19:23:31.3482269Z",
              "correlation": {
                "clientTrackingId": "08586652586741142222645090602CU35",
                "clientKeywords": ["testFlow"]
              },
              "status": "Succeeded"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, flow: 'edf73e7e-9928-4cb9-8eb2-fc263f375ada', environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', name: '08586652586741142222645090602CU35' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          name: '08586652586741142222645090602CU35',
          startTime: '2018-09-07T19:23:31.3640166Z',
          endTime: '',
          status: 'Running',
          triggerName: 'manual'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles environment not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "Access to the environment 'Default-48595cc3-adce-4267-8e99-0c838923dbbx' is denied."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, flow: '396d5ec9-ae2d-4a84-967d-cd7f56cd8f30', environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbbx', name: '08586653536760200319026785874CU62' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-48595cc3-adce-4267-8e99-0c838923dbbx' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles Flow not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "ConnectionAuthorizationFailed",
          "message": "The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '1c6ee23a-a835-44bc-a4f5-462b658efc12' under Api 'shared_logicflows'."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, flow: '1c6ee23a-a835-44bc-a4f5-462b658efc12', environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', name: '08586653536760200319026785874CU62' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '1c6ee23a-a835-44bc-a4f5-462b658efc12' under Api 'shared_logicflows'.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles run not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "BadRequest",
          "message": "The provided workflow run name is not valid."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, flow: '396d5ec9-ae2d-4a84-967d-cd7f56cd8f30', environment: 'Default-48595cc3-adce-4267-8e99-0c838923dbb9', name: 'ABC' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The provided workflow run name is not valid.`)));
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

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d' } }, (err?: any) => {
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

  it('supports specifying environment', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--environment') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying flow', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--flow') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying name', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});