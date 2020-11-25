import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./service-report-historicalservicestatus');

describe(commands.TENANT_SERVICE_REPORT_HISTORICALSERVICESTATUS, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
    if (!auth.service.accessTokens[auth.defaultResource]) {
      auth.service.accessTokens[auth.defaultResource] = {
        expiresOn: 'abc',
        value: 'abc'
      };
    }
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

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['WorkloadDisplayName', 'StatusDisplayName', 'StatusTime']);
  });

  it('supports debug mode', () => {
    const options = command.options();
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

    command.action(logger, {
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

  it('Gets the historical service status of the Office 365 Services of the last 7 days', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Gets the historical service status of the Office 365 Services of the last 7 days (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Gets the historical service status of the Office 365 Services of the last 7 days With Workload', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        workload: 'Bookings',
        debug: false
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Gets the historical service status of the Office 365 Services of the last 7 days With Workload (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('HistoricalStatus') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        workload: 'Bookings',
        debug: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});