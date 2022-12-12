import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./serviceannouncement-health-get');

describe(commands.SERVICEANNOUNCEMENT_HEALTH_GET, () => {
  const serviceHealthResponse = {
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
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    (command as any).items = [];
    (command as any).planId = undefined;
    (command as any).bucketId = undefined;
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
      pid.getProcessName
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

  it('passes validation when command called', async () => {
    const actual = await command.validate({
      options: {
        serviceName: "Exchange Online"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when command called with issues', async () => {
    const actual = await command.validate({
      options: {
        serviceName: "Exchange Online",
        issues: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly returns service health', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews/Exchange Online`) {
        return Promise.resolve(serviceHealthResponse);
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      serviceName: "Exchange Online"
    };

    await command.action(logger, { options } as any);
    assert(loggerLogSpy.calledWith(serviceHealthResponse));
  });


  it('correctly returns service health as csv with issues flag', async () => {
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

    await command.action(logger, { options } as any);
    assert(loggerLogSpy.calledWith(serviceHealthResponseCSV));
  });

  it('correctly returns service health with issues', async () => {
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

    await command.action(logger, { options } as any);
    assert(loggerLogSpy.calledWith(serviceHealthIssueResponse));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
