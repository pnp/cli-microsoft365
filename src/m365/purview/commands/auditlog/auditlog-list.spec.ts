import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './auditlog-list.js';

describe(commands.AUDITLOG_LIST, () => {

  const contentType = 'SharePoint';
  const now = new Date();
  const endTime = now.toISOString();
  const startTimeDate = new Date(now);
  startTimeDate.setHours(startTimeDate.getHours() - 10);
  const startTime = startTimeDate.toISOString();

  const contentTypeValue = 'Audit.SharePoint';
  const tenantId = '174290ec-373f-4d4c-89ea-9801dad0acd9';

  const listSubscriptionsResponse = [
    {
      contentType: contentTypeValue,
      status: 'enabled'
    }
  ];

  const auditLogs = [
    {
      CreationTime: "2023-01-13T00:00:00",
      Id: "c20660ac-fabd-4c04-d22d-08daf2baf8df",
      Operation: "ListItemUpdated",
      OrganizationId: tenantId,
      RecordType: 36,
      UserKey: "i:0i.t|00000003-0000-0ff1-ce00-000000000000|app@sharepoint",
      UserType: 0,
      Version: 1,
      Workload: "SharePoint",
      ClientIP: "40.114.162.47",
      ObjectId: "https://contoso.sharepoint.com/sites/project/Lists/5346c5ac-2d16-495c-9795-93577a1e4fe3/706_.000",
      UserId: "app@sharepoint"
    },
    {
      CreationTime: "2023-01-10T00:00:00",
      Id: "830f7b81-b3f2-4abe-faf3-08daf2baf8c7",
      Operation: "ListItemViewed",
      OrganizationId: tenantId,
      RecordType: 36,
      UserKey: "i:0i.t|00000003-0000-0ff1-ce00-000000000000|app@sharepoint",
      UserType: 0,
      Version: 1,
      Workload: "SharePoint",
      ClientIP: "40.114.162.47",
      ObjectId: "https://contoso.sharepoint.com/sites/project/Lists/5346c5ac-2d16-495c-9795-93577a1e4fe3/709_.000",
      UserId: "app@sharepoint"
    },
    {
      CreationTime: "2023-01-11T00:00:00",
      Id: "34b605b0-b97b-41d8-00b3-08daf2baf84c",
      Operation: "ListItemUpdated",
      OrganizationId: tenantId,
      RecordType: 36,
      UserKey: "i:0i.t|00000003-0000-0ff1-ce00-000000000000|app@sharepoint",
      UserType: 0,
      Version: 1,
      Workload: "SharePoint",
      ClientIP: "40.114.162.47",
      ObjectId: "https://contoso.sharepoint.com/sites/project/Lists/5346c5ac-2d16-495c-9795-93577a1e4fe3/725_.000",
      UserId: "app@sharepoint"
    },
    {
      CreationTime: "2023-01-12T00:00:00",
      Id: "d0c9679d-6854-4228-574f-08daf2baf7de",
      Operation: "ListItemUpdated",
      OrganizationId: tenantId,
      RecordType: 36,
      UserKey: "i:0i.t|00000003-0000-0ff1-ce00-000000000000|app@sharepoint",
      UserType: 0,
      Version: 1,
      Workload: "SharePoint",
      ClientIP: "40.114.162.47",
      ObjectId: "https://contoso.sharepoint.com/sites/project/Lists/5346c5ac-2d16-495c-9795-93577a1e4fe3/738_.000",
      UserId: "app@sharepoint"
    },
    {
      CreationTime: "2023-01-10T00:00:00",
      Id: "5df15c42-a005-4cc1-73d1-08daf2baf880",
      Operation: "ListItemUpdated",
      OrganizationId: tenantId,
      RecordType: 36,
      UserKey: "i:0i.t|00000003-0000-0ff1-ce00-000000000000|app@sharepoint",
      UserType: 0,
      Version: 1,
      Workload: "SharePoint",
      ClientIP: "40.114.162.47",
      ObjectId: "https://contoso.sharepoint.com/sites/project/Lists/5346c5ac-2d16-495c-9795-93577a1e4fe3/716_.000",
      UserId: "app@sharepoint"
    }
  ];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    if (!auth.connection.accessTokens[auth.defaultResource]) {
      auth.connection.accessTokens[auth.defaultResource] = {
        expiresOn: 'abc',
        accessToken: 'abc'
      };
    }
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
    sinon.stub(accessToken, 'getTenantIdFromAccessToken').returns(tenantId);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      accessToken.getTenantIdFromAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.AUDITLOG_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if contentType has an invalid value', async () => {
    const actual = await command.validate({ options: { contentType: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if startTime is not a valid date', async () => {
    const actual = await command.validate({ options: { contentType: contentType, startTime: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if endTime is not a valid date', async () => {
    const actual = await command.validate({ options: { contentType: contentType, endTime: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if startTime is more than 7 days in the past', async () => {
    const startTime = new Date();
    startTime.setDate(startTime.getDate() - 7);
    startTime.setHours(startTime.getHours() - 2);
    const actual = await command.validate({ options: { contentType: contentType, startTime: startTime.toISOString() } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if endTime is in the future', async () => {
    const endTime = new Date();
    endTime.setHours(endTime.getHours() + 1);
    const actual = await command.validate({ options: { contentType: contentType, endTime: endTime.toISOString() } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if endTime is before startTime', async () => {
    const startTime = new Date();
    const endTime = new Date(startTime);
    endTime.setTime(endTime.getTime() - 1);
    const actual = await command.validate({ options: { contentType: contentType, startTime: startTime.toISOString(), endTime: endTime.toISOString() } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if only contentType is provided', async () => {
    const actual = await command.validate({ options: { contentType: contentType } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if startTime and endTime are provided', async () => {
    const actual = await command.validate({ options: { contentType: contentType, startTime: startTime, endTime: endTime } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('handles error when unable to start new subscription', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/subscriptions/list`) {
        return [];
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/subscriptions/start?contentType=DLP.All`) {
        return {
          contentType: contentTypeValue,
          status: 'disabled'
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, {
      options: {
        contentType: 'DLP'
      }
    }), new CommandError(`Unable to start subscription 'DLP.All'`));
  });

  it('starts subscription if there was no subscription active', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/subscriptions/list`) {
        return [];
      }

      if (opts.url === `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/subscriptions/content?contentType=${contentTypeValue}&startTime=${startTime}&endTime=${endTime}`) {
        return { headers: {}, data: [] };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/subscriptions/start?contentType=${contentTypeValue}`) {
        return listSubscriptionsResponse[0];
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        contentType: contentType,
        startTime: startTime,
        endTime: endTime
      }
    });

    assert(postStub.called);
  });

  it('retrieves audit logs correctly', async () => {
    const contentUriApiUrl = `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/subscriptions/content?contentType=${contentTypeValue}&startTime=${startTime}&endTime=${endTime}`;

    const contentUris = [
      `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/audit/20230110010221444060394$20230110033214410058910`,
      `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/audit/20230110033214617061387$20230110033216340073677`,
      `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/audit/20230110033216408077102$20230110033218101079570`
    ];

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/subscriptions/list`) {
        return [
          {
            contentType: contentTypeValue,
            status: 'enabled'
          }
        ];
      }

      if (opts.url === contentUriApiUrl) {
        return {
          headers: { nextpageuri: contentUriApiUrl + '&page=2' },
          data: [
            {
              contentUri: contentUris[0],
              contentId: "20230110010221444060394$20230110033214410058910",
              contentType: contentTypeValue
            },
            {
              contentUri: contentUris[1],
              contentId: "20230110033214617061387$20230110033216340073677",
              contentType: contentTypeValue
            }
          ]
        };
      }

      if (opts.url === contentUriApiUrl + '&page=2') {
        return {
          headers: {},
          data: [
            {
              contentUri: contentUris[2],
              contentId: "20230110033216408077102$20230110033218101079570",
              contentType: contentTypeValue
            }
          ]
        };
      }

      if (opts.url === contentUris[0]) {
        return auditLogs.slice(0, 2);
      }

      if (opts.url === contentUris[1]) {
        return auditLogs.slice(2, 4);
      }

      if (opts.url === contentUris[2]) {
        return auditLogs.slice(4, 6);
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        contentType: contentType,
        startTime: startTime,
        endTime: endTime,
        verbose: true
      }
    });

    assert(loggerLogSpy.calledWith(auditLogs.sort((a, b) => a.CreationTime < b.CreationTime ? -1 : a.CreationTime > b.CreationTime ? 1 : 0)));
  });
});
