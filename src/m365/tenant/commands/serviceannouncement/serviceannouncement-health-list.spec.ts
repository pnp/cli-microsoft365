import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./serviceannouncement-health-list');

describe(commands.SERVICEANNOUNCEMENT_HEALTH_LIST, () => {
  const serviceHealthResponse = [
    {
      "service": "Exchange Online",
      "status": "serviceOperational",
      "id": "Exchange"
    },
    {
      "service": "Identity Service",
      "status": "serviceOperational",
      "id": "OrgLiveID"
    },
    {
      "service": "Microsoft 365 suite",
      "status": "serviceOperational",
      "id": "OSDPPlatform"
    }
  ];

  const serviceHealthResponseCSV = `service,status,id
    Exchange Online,serviceDegradation,Exchange
    Identity Service,serviceOperational,OrgLiveID
    Microsoft 365 suite,serviceOperational,OSDPPlatform
    Skype for Business,serviceOperational,Lync
    SharePoint Online,serviceOperational,SharePoint
    Dynamics 365 Apps,serviceOperational,DynamicsCRM
    Azure Information Protection,serviceOperational,RMS
    Yammer Enterprise,serviceOperational,yammer
    Mobile Device Management for Office 365,serviceOperational,MobileDeviceManagement
    Planner,serviceOperational,Planner
    Sway,serviceOperational,SwayEnterprise
    Power BI,serviceOperational,PowerBIcom
    Microsoft Intune,extendedRecovery,Intune
    OneDrive for Business,serviceOperational,OneDriveForBusiness
    Microsoft Teams,serviceOperational,microsoftteams
    Microsoft StaffHub,serviceOperational,StaffHub
    Microsoft Bookings,serviceOperational,Bookings
    Office for the web,serviceOperational,officeonline
    Microsoft 365 Apps,serviceOperational,O365Client
    Power Apps,serviceOperational,PowerApps
    Power Apps in Microsoft 365,serviceOperational,PowerAppsM365
    Microsoft Power Automate,serviceOperational,MicrosoftFlow
    Microsoft Power Automate in Microsoft 365,serviceOperational,MicrosoftFlowM365
    Microsoft Forms,serviceOperational,Forms
    Microsoft 365 Defender,extendedRecovery,Microsoft365Defender
    Microsoft Stream,serviceOperational,Stream
    Privileged Access,serviceOperational,PAM
    Microsoft Viva,serviceOperational,Viva
    Microsoft Defender for Cloud Apps,serviceOperational,cloudappsecurity`;

  const serviceHealthIssuesResponse = [
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
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SERVICEANNOUNCEMENT_HEALTH_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'status', 'service']);
  });

  it('passes validation when command called', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when command called with issues', async () => {
    const actual = await command.validate({ options: { issues: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly returns list', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews`) {
        return Promise.resolve(
          {
            value: serviceHealthResponse
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {};

    await command.action(logger, { options } as any);
    assert(loggerLogSpy.calledWith(serviceHealthResponse));
  });

  it('correctly returns list as csv with issues flag', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews`) {
        return Promise.resolve(
          {
            value: serviceHealthResponseCSV
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      issues: true,
      output: "csv"
    };

    await command.action(logger, { options } as any);
    assert(loggerLogSpy.calledWith(serviceHealthResponseCSV));
  });

  it('correctly returns list with issues', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews?$expand=issues`) {
        return Promise.resolve(
          {
            value: serviceHealthIssuesResponse
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      issues: true
    };

    await command.action(logger, { options } as any);
    assert(loggerLogSpy.calledWith(serviceHealthIssuesResponse));
  });

  it('fails when serviceAnnouncement endpoint fails', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews`) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Error fetching service health'));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});