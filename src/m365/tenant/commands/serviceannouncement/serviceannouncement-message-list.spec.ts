import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./serviceannouncement-message-list');

describe(commands.SERVICEANNOUNCEMENT_MESSAGE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const jsonOutput = {
    "value": [
      {
        "startDateTime": "2021-02-01T19:23:04Z",
        "endDateTime": "2022-01-31T08:00:00Z",
        "lastModifiedDateTime": "2021-02-01T19:24:37.837Z",
        "title": "Service reminder: Skype for Business Online retires in 6 months",
        "id": "MC237349",
        "category": "planForChange",
        "severity": "normal",
        "tags": [
          "User impact",
          "Admin impact"
        ],
        "isMajorChange": false,
        "actionRequiredByDateTime": "2021-07-31T07:00:00Z",
        "services": [
          "Skype for Business"
        ],
        "expiryDateTime": null,
        "hasAttachments": false,
        "viewPoint": null,
        "details": [
          {
            "name": "BlogLink",
            "value": "https://techcommunity.microsoft.com/t5/microsoft-teams-blog/skype-for-business-online-will-retire-in-12-months-plan-for-a/ba-p/1554531"
          },
          {
            "name": "ExternalLink",
            "value": "https://docs.microsoft.com/microsoftteams/skype-for-business-online-retirement"
          }
        ],
        "body": {
          "contentType": "html",
          "content": "<p>Originally announced in MC219641 (July '20), as Microsoft Teams has become the core communications client for Microsoft 365, this is a reminder the Skype for Business Online service will <a href=\"https://techcommunity.microsoft.com/t5/microsoft-teams-blog/skype-for-business-online-will-retire-in-12-months-plan-for-a/ba-p/1554531\" target=\"_blank\">retire July 31, 2021</a>. At that point, access to the service will end.</p><p>Please note:</p><ul><li>Neither the Skype consumer service nor Skype for Business Server products are affected by the retirement of the Skype for Business Online service. Current Skype for Business Online customers will experience no change in service up to the retirement date.</li><li>Support for the integration of third-party party Audio Conferencing Providers (ACP) into Skype for Business Online will continue through the end of service July 31, 2021. </li></ul><p>[How does this affect me?]</p><p>While your organization will no longer be able to use Skype for Business Online once the service retires, July 31, 2021, you may continue to use Microsoft Teams as part of your existing licensing agreement.</p><p>[What do I need to do to prepare for this change?]</p><p>We encourage customers using Skype for Business Online to plan and begin the upgrade to Teams today. To help, Microsoft offers <a href=\"https://aka.ms/SkypeToTeams\" target=\"_blank\">comprehensive transition resources</a> including a proven upgrade framework, guidance for technical and organizational readiness, <a href=\"https://aka.ms/upgradeworkshops\" target=\"_blank\">free upgrade planning workshops</a>, and <a href=\"https://www.microsoft.com/en-us/FastTrack?rtc=1\" target=\"_blank\">FastTrack</a> onboarding assistance for eligible subscriptions.</p><p>Please click Additional Information to learn more.</p>"
        }
      },
      {
        "startDateTime": "2021-02-04T19:03:15Z",
        "endDateTime": "2022-06-30T07:00:00Z",
        "lastModifiedDateTime": "2021-02-04T19:03:41.873Z",
        "title": "Basic Authentication and Exchange Online – February 2021 Update",
        "id": "MC237741",
        "category": "planForChange",
        "severity": "normal",
        "tags": [
          "Admin impact",
          "Retirement"
        ],
        "isMajorChange": false,
        "actionRequiredByDateTime": null,
        "services": [
          "Exchange Online"
        ],
        "expiryDateTime": null,
        "hasAttachments": false,
        "viewPoint": null,
        "details": [
          {
            "name": "ExternalLink",
            "value": "https://techcommunity.microsoft.com/t5/exchange-team-blog/basic-authentication-and-exchange-online-february-2021-update/ba-p/2111904"
          }
        ],
        "body": {
          "contentType": "html",
          "content": "<p>In response to the unprecedented situation we are in and knowing that priorities have changed for many of our customers we are suspending until further notice the disabling of Basic Authentication for any protocols that your tenant is using. This was previously communicated in Exchange Online (MC204828 and MC208814). When we resume this program, we will provide a minimum of twelve months’ notice before we block the use of Basic Auth on any protocol being used.</p><p>We will continue with our plan to <b>disable Basic Auth for protocols that your tenant is <i>not </i>using.</b> Many customers don’t know that unneeded legacy protocols remain enabled in their tenant. We plan to disable these unused protocols to prevent potential mis-use. We will do this based on examining recorded usage of these protocols by your tenant, and we will send Message Center posts providing 30 days’ notice of any changes to your tenant. This work will begin in a few months.</p><p>The last change to the previously announced plan is that we are adding MAPI, RPC, and Offline Address Book (OAB) to the protocols included in this effort to further enhance data protection.</p><p>[How this impacts your organization:]</p><ul><li>How Will I Know When My Tenant Is Affected?<ul><li>We will publish a major change Message Center post to your tenant 30 days prior to us disabling Basic Auth for any protocols in your tenant. Major changes also trigger email notifications. We will also publish a (non-major change) Message Center post when we have made the actual change.</li></ul></li><li>What If My Tenant is Using One of These Protocols?<ul><li>If your tenant is using any of these protocols, we won’t disable them. Should you find a Message Center post to the contrary, please let us know (details on how to let us know will be in the Message Center post) and we’ll exclude you from the change. You’ll be able to do this right up until we disable these protocols for good (at a future date).</li></ul></li><li>What Happens If I Missed the Message Center Post and Need These Protocols Re-Enabled?<ul><li>We are building the capability to allow you to re-enable the protocols yourself via Support Central in the Microsoft 365 admin center. If you find yourself in this situation, you’ll be able to request help in the Microsoft 365 admin center, and we’ll allow you to re-enable these protocols until we disable them in the future.</li></ul></li><li>How Does This Change Affect Authentication Policies?<ul><li>The switch we use to disable Basic Auth for unused protocols is not available to tenant admins (with the exception of the switch for SMTP Auth). You won’t see any changes or additions to your existing authentication policies (if you have any) and our change will take precedence over any policies you might have. We understand this might be a bit confusing, so we wanted to note it here.</li></ul></li><li>Does this Change Affect Outlook?<ul><li>Outlook depends upon Exchange Web Services (EWS). Therefore, Outlook must be updated to use Modern Auth before Basic Auth for EWS is disabled. Outlook uses only one type of authentication for all connections to a mailbox, so including these protocols should not adversely affect you. If EWS has Basic Auth disabled, Outlook won’t use Basic Auth for any of the other protocols or endpoints it needs to access.</li></ul></li></ul><p>We hope this change is good news for those of you who needed more time to complete a transition from Basic Auth.</p><p>Please click Additional Information to learn more.</p>"
        }
      },
      {
        "startDateTime": "2021-03-30T22:55:52Z",
        "endDateTime": "2022-01-31T08:00:00Z",
        "lastModifiedDateTime": "2021-03-30T22:56:53.887Z",
        "title": "Teams for your personal life banners",
        "id": "MC247825",
        "category": "planForChange",
        "severity": "normal",
        "tags": [
          "New feature",
          "User impact",
          "Admin impact"
        ],
        "isMajorChange": true,
        "actionRequiredByDateTime": null,
        "services": [
          "Microsoft Teams"
        ],
        "expiryDateTime": null,
        "hasAttachments": false,
        "viewPoint": null,
        "details": [],
        "body": {
          "contentType": "html",
          "content": "<p>To inform users about the new productivity features and the ability to sign in with different accounts, we will show a banner in the activity feed for Teams mobile app users.</p><p>[Key points:]</p><ul><li>Timeline: Begin at end of April expect to complete rolling out to all customers by the end of December.</li><li>Control: Contact Support</li></ul><p>[How this will affect your organization:]</p><p>Once the change has rolled out to your tenant, users will see a banner in their activity feed saying that they can add a personal account to the Teams app on their mobile device. If you have previously disabled the ability to add additional accounts, users will not see the banner.</p><p>To manage the visibility of the banners to your users, you can submit a help ticket in the Microsoft 365 admin center and your tenant will be excluded from the banners. This will not limit your user’s ability to add a personal, work or school account to the Teams app. If you want to restrict users from adding a personal, work, or school account, instructions are available <a href=\"https://docs.microsoft.com/microsoftteams/sign-in-teams#restrict-sign-in-to-teams\" target=\"_blank\">here</a>. Restricted users will not see the banner.</p><p>[What you need to do to prepare:]</p><p>There is no action you need to take to prepare for this change, but you might consider updating your user training and notifying your help desk.</p><p>More information about the personal experience in Teams can be found <a href=\"https://www.microsoft.com/microsoft-teams/teams-for-home\" target=\"_blank\">here</a>.</p>"
        }
      },
      {
        "startDateTime": "2021-04-01T23:35:18Z",
        "endDateTime": "2022-02-28T07:00:00Z",
        "lastModifiedDateTime": "2022-01-11T19:47:40.097Z",
        "title": "(Updated) Quick Create – Easily Create Power BI Reports from Lists",
        "id": "MC248201",
        "category": "planForChange",
        "severity": "normal",
        "tags": [
          "Updated message",
          "New feature",
          "User impact",
          "Admin impact"
        ],
        "isMajorChange": true,
        "actionRequiredByDateTime": null,
        "services": [
          "SharePoint Online",
          "Power BI"
        ],
        "expiryDateTime": null,
        "hasAttachments": false,
        "viewPoint": null,
        "details": [],
        "body": {
          "contentType": "html",
          "content": "<p>Updated January 11, 2022: We have updated the rollout timeline below for Government organizations. Thank you for your patience.</p><p>We are excited to announce the arrival of a new guided authoring experience in Lists that will make it easy to quickly create business intelligence reports in Power BI using your list schema and data.</p><p>[Key points]</p><ul><li>Microsoft 365 <a href=\"https://www.microsoft.com/microsoft-365/roadmap?filters=&amp;searchterms=72175%2C\" target=\"_blank\">Roadmap ID 72175</a>.</li><li>Timing: <ul><li>Targeted Release: rolling out in early early May. - Completed</li><li>Standard Release: rolling out from early June (previously late May) to mid-July (previously early June). - Completed</li><li>Government: we will begin rolling out in early December (previously mid-September) and expect to complete by late January (previously mid-December).</li></ul></li><li>Roll-out: tenant level</li><li>Control type: admin control</li><li>Action: review and assess</li></ul><p>[How this will affect your organization]</p><p>List users will see a new menu option in <i>Integrate </i>&gt; <i>Power BI</i> &gt; <i>Visualize this list</i>, which allows users to create a new Power BI report using that list. With just one click, you'll be able to autogenerate a basic report and customize the list columns that are shown in the report. To take further advantage of Power BI’s advanced data visualization capabilities, just go into <i><a href=\"https://docs.microsoft.com/en-us/power-bi/create-reports/service-interact-with-a-report-in-editing-view\" target=\"_blank\">Edit mode</a></i>. Once a report is saved and published, it will appear in the same submenu under Integrate<i> </i>&gt; <i>Power BI</i>.</p><p><img src=\"https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RWANSC?ver=cf60\" alt=\"Power BI Submenu Entry Point\" width=\"550\"></p><ul><li>Users with a <a href=\"https://www.microsoft.com/microsoft-365/enterprise/e5\" target=\"_blank\">Microsoft 365 E5 license</a> or <a href=\"https://powerbi.microsoft.com/power-bi-pro/\" target=\"_blank\">Power BI Pro license</a> will have access to the full report authoring and viewing experience. </li><li>Users without either of the above licenses will be prompted by Power BI to sign up for a 60-day free trial of Power BI Pro when they attempt to save a new report or edit or view an existing report. To turn off self-service sign-up so that the option for a trial is not exposed to List users, click <a href=\"https://docs.microsoft.com/power-bi/admin/service-admin-disable-self-service\" target=\"_blank\">here</a>.</li><li>Users with a Power BI free license may only visualize their list data, but cannot publish nor view reports.</li></ul><p>[What you need to do to prepare]</p><p>This feature is default on, but can turned off from the <a href=\"https://docs.microsoft.com/power-bi/admin/service-admin-portal\" target=\"_blank\">Power BI Admin Portal</a> under <i>Tenant settings</i>.</p> <p>If this feature is disabled for tenants, users will continue to see the Power BI submenu in the List command bar, but any attempt to create or view a report will result in an error page.</p><p><b>Note:</b></p><p>Certain complex column types in Lists such as Person, Location, Rich Text, Multi-select Choice, and Image are not currently supported when the Power BI report is autogenerated.</p><p>Learn more:</p><ul><li><a href=\"https://docs.microsoft.com/power-bi/admin/service-admin-manage-licenses\" target=\"_blank\">View and manage Power BI user licenses</a></li><li><a href=\"https://docs.microsoft.com/power-bi/fundamentals/service-get-started\" target=\"_blank\">Get started creating in the Power BI service </a></li><li><a href=\"https://powerbi.microsoft.com/blog/quickly-create-reports-power-bi-service/\" target=\"_blank\">Quickly create reports in the Power BI service </a></li><li><a href=\"https://docs.microsoft.com/power-bi/admin/service-admin-disable-self-service\" target=\"_blank\">Enable or disable self-service sign-up and purchasing</a></li></ul>"
        }
      }
    ]
  };

  const jsonOutputMicrosoftTeams = {
    "value": [
      {
        "startDateTime": "2021-03-30T22:55:52Z",
        "endDateTime": "2022-01-31T08:00:00Z",
        "lastModifiedDateTime": "2021-03-30T22:56:53.887Z",
        "title": "Teams for your personal life banners",
        "id": "MC247825",
        "category": "planForChange",
        "severity": "normal",
        "tags": [
          "New feature",
          "User impact",
          "Admin impact"
        ],
        "isMajorChange": true,
        "actionRequiredByDateTime": null,
        "services": [
          "Microsoft Teams"
        ],
        "expiryDateTime": null,
        "hasAttachments": false,
        "viewPoint": null,
        "details": [],
        "body": {
          "contentType": "html",
          "content": "<p>To inform users about the new productivity features and the ability to sign in with different accounts, we will show a banner in the activity feed for Teams mobile app users.</p><p>[Key points:]</p><ul><li>Timeline: Begin at end of April expect to complete rolling out to all customers by the end of December.</li><li>Control: Contact Support</li></ul><p>[How this will affect your organization:]</p><p>Once the change has rolled out to your tenant, users will see a banner in their activity feed saying that they can add a personal account to the Teams app on their mobile device. If you have previously disabled the ability to add additional accounts, users will not see the banner.</p><p>To manage the visibility of the banners to your users, you can submit a help ticket in the Microsoft 365 admin center and your tenant will be excluded from the banners. This will not limit your user’s ability to add a personal, work or school account to the Teams app. If you want to restrict users from adding a personal, work, or school account, instructions are available <a href=\"https://docs.microsoft.com/microsoftteams/sign-in-teams#restrict-sign-in-to-teams\" target=\"_blank\">here</a>. Restricted users will not see the banner.</p><p>[What you need to do to prepare:]</p><p>There is no action you need to take to prepare for this change, but you might consider updating your user training and notifying your help desk.</p><p>More information about the personal experience in Teams can be found <a href=\"https://www.microsoft.com/microsoft-teams/teams-for-home\" target=\"_blank\">here</a>.</p>"
        }
      }
    ]
  };

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SERVICEANNOUNCEMENT_MESSAGE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title']);
  });

  it('handles promise error while getting service update messages available in Microsoft 365', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/admin/serviceAnnouncement/messages') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: {} } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the service update messages available in Microsoft 365', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/admin/serviceAnnouncement/messages') > -1) {
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
        assert(loggerLogSpy.calledWith(jsonOutput.value));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the service update messages available in Microsoft 365 (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/admin/serviceAnnouncement/messages') > -1) {
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
        assert(loggerLogSpy.calledWith(jsonOutput.value));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the service update messages for a particular service available in Microsoft 365', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/admin/serviceAnnouncement/messages') > -1) {
        return Promise.resolve(jsonOutputMicrosoftTeams);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        service: 'Microsoft Teams',
        debug: false
      }
    } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(jsonOutputMicrosoftTeams.value));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the service update messages for a particular service available in Microsoft 365 as text', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/admin/serviceAnnouncement/messages') > -1) {
        return Promise.resolve(jsonOutputMicrosoftTeams);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        service: 'Microsoft Teams',
        output: 'text',
        debug: false
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(jsonOutputMicrosoftTeams.value));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
