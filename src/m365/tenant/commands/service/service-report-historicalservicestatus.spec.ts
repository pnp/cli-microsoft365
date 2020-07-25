import commands from '../../commands';
import Command, { CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./service-report-historicalservicestatus');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TENANT_SERVICE_REPORT_HISTORICALSERVICESTATUS, () => {
  let log: any[];
  let cmdInstance: any;

  let cmdInstanceLogSpy: sinon.SinonSpy;

  let textOutput = [
    {
      WorkloadDisplayName: "Microsoft Bookings",
      StatusDisplayName: "Normal service",
      StatusTime: "2020-09-17T00:00:00Z"
    },
    {
      WorkloadDisplayName: "Microsoft Bookings",
      StatusDisplayName: "Normal service",
      StatusTime: "2020-09-16T00:00:00Z"
    },
    {
      WorkloadDisplayName: "Microsoft Bookings",
      StatusDisplayName: "Normal service",
      StatusTime: "2020-09-15T00:00:00Z"
    },
    {
      WorkloadDisplayName: "Microsoft Bookings",
      StatusDisplayName: "Normal service",
      StatusTime: "2020-09-14T00:00:00Z"
    },
    {
      WorkloadDisplayName: "Microsoft Bookings",
      StatusDisplayName: "Normal service",
      StatusTime: "2020-09-13T00:00:00Z"
    },
    {
      WorkloadDisplayName: "Microsoft Bookings",
      StatusDisplayName: "Normal service",
      StatusTime: "2020-09-12T00:00:00Z"
    },
    {
      WorkloadDisplayName: "Microsoft Bookings",
      StatusDisplayName: "Normal service",
      StatusTime: "2020-09-11T00:00:00Z"
    }
  ];

  let jsonOutput = {
    "@odata.context": "https://office365servicecomms-prod.cloudapp.net/api/v1.0/contoso.sharepoint.com/$metadata#CurrentStatus",
    "value": [
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Microsoft Bookings",
            "FeatureName": "MicrosoftBookings",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Bookings",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-17T00:00:00Z",
        "Workload": "Bookings",
        "WorkloadDisplayName": "Microsoft Bookings"
      },
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Microsoft Bookings",
            "FeatureName": "MicrosoftBookings",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Bookings",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-16T00:00:00Z",
        "Workload": "Bookings",
        "WorkloadDisplayName": "Microsoft Bookings"
      },
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Microsoft Bookings",
            "FeatureName": "MicrosoftBookings",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Bookings",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-15T00:00:00Z",
        "Workload": "Bookings",
        "WorkloadDisplayName": "Microsoft Bookings"
      },
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Microsoft Bookings",
            "FeatureName": "MicrosoftBookings",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Bookings",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-14T00:00:00Z",
        "Workload": "Bookings",
        "WorkloadDisplayName": "Microsoft Bookings"
      },
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Microsoft Bookings",
            "FeatureName": "MicrosoftBookings",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Bookings",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-13T00:00:00Z",
        "Workload": "Bookings",
        "WorkloadDisplayName": "Microsoft Bookings"
      },
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Microsoft Bookings",
            "FeatureName": "MicrosoftBookings",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Bookings",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-12T00:00:00Z",
        "Workload": "Bookings",
        "WorkloadDisplayName": "Microsoft Bookings"
      },
      {
        "FeatureStatus": [
          {
            "FeatureDisplayName": "Microsoft Bookings",
            "FeatureName": "MicrosoftBookings",
            "FeatureServiceStatus": "ServiceOperational",
            "FeatureServiceStatusDisplayName": "Normal service"
          }
        ],
        "Id": "Bookings",
        "IncidentIds": [],
        "Status": "ServiceOperational",
        "StatusDisplayName": "Normal service",
        "StatusTime": "2020-09-11T00:00:00Z",
        "Workload": "Bookings",
        "WorkloadDisplayName": "Microsoft Bookings"
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    auth.service.tenantId = '48526e9f-60c5-3000-31d7-aa1dc75ecf3c|908bel80-a04a-4422-b4a0-883d9847d110:c8e761e2-d528-34d1-8776-dc51157d619a&#xA;Tenant';
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
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.TENANT_SERVICE_REPORT_HISTORICALSERVICESTATUS), true);
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

  it('handles promise error while getting the historical service status of the Office 365 Services of the last 7 days', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
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

  it('Gets the historical service status of the Office 365 Services of the last 7 days - JSON Output', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
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

  it('Gets the historical service status of the Office 365 Services of the last 7 days - JSON Output (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
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

  it('Gets the historical service status of the Office 365 Services of the last 7 days - text Output', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
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

  it('Gets the historical service status of the Office 365 Services of the last 7 days - text Output (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
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

  it('Gets the historical service status of the Office 365 Services of the last 7 days - JSON Output With Workload', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        workload: 'Bookings',
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

  it('Gets the historical service status of the Office 365 Services of the last 7 days - JSON Output With Workload (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        workload: 'Bookings',
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

  it('Gets the historical service status of the Office 365 Services of the last 7 days - text Output With Workload', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        workload: 'Bookings',
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

  it('Gets the historical service status of the Office 365 Services of the last 7 days - text Output With Workload (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        workload: 'Bookings',
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