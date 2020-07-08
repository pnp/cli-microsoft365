import commands from '../../commands';
import Command, { CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./status-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TENANT_STATUS_LIST, () => {
  let vorpal: Vorpal;
  let log: any[];
  let requests: any[];
  let cmdInstance: any;

  let cmdInstanceLogSpy: sinon.SinonSpy;

  let textOutput = [
    {
      Name: "Microsoft Forms",
      Status: "Normal service"
    }
  ];

  let jsonOutput = {
    "@odata.context": "https://office365servicecomms-prod.cloudapp.net/api/v1.0/contoso.sharepoint.com/$metadata#CurrentStatus",
    "value": [
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Service",
            "FeatureName": "service",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Form functionality",
            "FeatureName": "functionality",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          },
          {
            "FeatureDisplayName": "Integration",
            "FeatureName": "integration",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Forms",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-03-24T23:32:35.7309744Z",
        "Workload": "Forms",
        "WorkloadDisplayName": "Microsoft Forms"
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    requests = [];
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
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.TENANT_STATUS_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TENANT_STATUS_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('handles promise error while getting status of Microsoft 365 services', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {

      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'json',
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'json',
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services as text', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(textOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the status of Microsoft 365 services as text (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('CurrentStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(textOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

});