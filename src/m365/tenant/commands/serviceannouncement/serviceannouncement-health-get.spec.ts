import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./serviceannouncement-health-get');

describe(commands.SERVICEANNOUNCEMENT_HEALTH_GET, () => {
  const serviceHealthResponse =  {
    "service": "Exchange Online",
    "status": "serviceOperational",
    "id": "Exchange"
  };

  const serviceHealthResponseCSV = `service,status,id
    Exchange Online,serviceDegradation,Exchange`;

  const serviceHealthIssueResponse = [
    {
      "service": "Exchange Online",
      "status": "serviceOperational",
      "id": "Exchange",
      "issues": [
        {
          "startDateTime": "2020-11-04T00:00:00Z",
          "endDateTime": "2020-11-20T17:00:00Z",
          "lastModifiedDateTime": "2020-11-20T17:56:31.39Z",
          "title": "Admins are unable to migrate some user mailboxes from IMAP using the Exchange admin center or PowerShell",
          "id": "EX226574",
          "impactDescription": "Admins attempting to migrate some user mailboxes using the Exchange admin center or PowerShell experienced failures.",
          "classification": "Advisory",
          "origin": "Microsoft",
          "status": "ServiceRestored",
          "service": "Exchange Online",
          "feature": "Tenant Administration (Provisioning, Remote PowerShell)",
          "featureGroup": "Management and Provisioning",
          "isResolved": true,
          "details": [],
          "posts": [
            {
              "createdDateTime": "2020-11-12T07:07:38.97Z",
              "postType": "Regular",
              "description": {
                "contentType": "Text",
                "content": "Title: Exchange Online service has login issue. We'll provide an update within 30 minutes."
              }
            }
          ]
        }
      ]
    }
  ];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    (command as any).items = [];
    (command as any).planId = undefined;
    (command as any).bucketId = undefined;
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
    assert.strictEqual(command.name.startsWith(commands.SERVICEANNOUNCEMENT_HEALTH_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'status', 'service']);
  });
  
  it('passes validation when command called', (done) => {
    const actual = command.validate({
      options: {
        serviceName: "Exchange Online"
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when command called with issues', (done) => {
    const actual = command.validate({
      options: {
        serviceName: "Exchange Online",
        issues: true
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('correctly returns service health', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews/Exchange Online`) {
        return Promise.resolve(serviceHealthResponse);
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      serviceName: "Exchange Online"
    };

    command.action(logger, {options} as any, () => {
      try {
        assert(loggerLogSpy.calledWith(serviceHealthResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  
  it('correctly returns service health as csv with issues flag', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews/Exchange Online`) {
        return Promise.resolve(serviceHealthResponseCSV);
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      serviceName: "Exchange Online",
      issues: true,
      output: "csv"
    };

    command.action(logger, {options} as any, () => {
      try {
        assert(loggerLogSpy.calledWith(serviceHealthResponseCSV));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly returns service health with issues', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews/Exchange Online?$expand=issues`) {
        return Promise.resolve(serviceHealthIssueResponse);
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      serviceName: "Exchange Online",
      issues: true
    };

    command.action(logger, {options} as any, () => {
      try {
        assert(loggerLogSpy.calledWith(serviceHealthIssueResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when serviceAnnouncement endpoint fails', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews/Exchange Online`) {
        return Promise.resolve(undefined);
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      serviceName: "Exchange Online"
    };

    command.action(logger, {options} as any, (err?: any) => {
      try {
        assert.strictEqual(err.message, "Error fetching service health");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
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
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});