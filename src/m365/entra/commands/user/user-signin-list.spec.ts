import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './user-signin-list.js';

describe(commands.USER_SIGNIN_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const jsonOutput = {
    "value": [
      {
        "id": "66ea54eb-6301-4ee5-be62-ff5a759b0100",
        "createdDateTime": "2020-03-13T19:15:41.6195833Z",
        "userDisplayName": "Test Contoso",
        "userPrincipalName": "testaccount1@contoso.com",
        "userId": "26be570a-ae82-4189-b4e2-a37c6808512d",
        "appId": "de8bc8b5-d9f9-48b1-a8ad-b748da725064",
        "appDisplayName": "Graph explorer",
        "ipAddress": "131.107.159.37",
        "clientAppUsed": "Browser",
        "correlationId": "d79f5bee-5860-4832-928f-3133e22ae912",
        "conditionalAccessStatus": "notApplied",
        "isInteractive": true,
        "riskDetail": "none",
        "riskLevelAggregated": "none",
        "riskLevelDuringSignIn": "none",
        "riskState": "none",
        "riskEventTypes": [],
        "resourceDisplayName": "Microsoft Graph",
        "resourceId": "00000003-0000-0000-c000-000000000000",
        "status": {
          "errorCode": 0,
          "failureReason": null,
          "additionalDetails": null
        },
        "deviceDetail": {
          "deviceId": "",
          "displayName": null,
          "operatingSystem": "Windows 10",
          "browser": "Edge 80.0.361",
          "isCompliant": null,
          "isManaged": null,
          "trustType": null
        },
        "location": {
          "city": "Redmond",
          "state": "Washington",
          "countryOrRegion": "US",
          "geoCoordinates": {
            "altitude": null,
            "latitude": 47.68050003051758,
            "longitude": -122.12094116210938
          }
        },
        "appliedConditionalAccessPolicies": [
          {
            "id": "de7e60eb-ed89-4d73-8205-2227def6b7c9",
            "displayName": "SharePoint limited access for guest workers",
            "enforcedGrantControls": [],
            "enforcedSessionControls": [],
            "result": "notEnabled"
          },
          {
            "id": "6701123a-b4c6-48af-8565-565c8bf7cabc",
            "displayName": "Medium signin risk block",
            "enforcedGrantControls": [],
            "enforcedSessionControls": [],
            "result": "notEnabled"
          }
        ]
      },
      {
        "id": "66ea54eb-6301-4ee5-be62-ff5a759b0100",
        "createdDateTime": "2020-03-13T19:15:41.6195833Z",
        "userDisplayName": "Test Contoso",
        "userPrincipalName": "testaccount1@contoso.com",
        "userId": "26be570a-ae82-4189-b4e2-a37c6808512d",
        "appId": "de8bc8b5-d9f9-48b1-a8ad-b748da725064",
        "appDisplayName": "Graph explorer",
        "ipAddress": "131.107.159.37",
        "clientAppUsed": "Browser",
        "correlationId": "d79f5bee-5860-4832-928f-3133e22ae912",
        "conditionalAccessStatus": "notApplied",
        "isInteractive": true,
        "riskDetail": "none",
        "riskLevelAggregated": "none",
        "riskLevelDuringSignIn": "none",
        "riskState": "none",
        "riskEventTypes": [],
        "resourceDisplayName": "Microsoft Graph",
        "resourceId": "00000003-0000-0000-c000-000000000000",
        "status": {
          "errorCode": 0,
          "failureReason": null,
          "additionalDetails": null
        },
        "deviceDetail": {
          "deviceId": "",
          "displayName": null,
          "operatingSystem": "Windows 10",
          "browser": "Edge 80.0.361",
          "isCompliant": null,
          "isManaged": null,
          "trustType": null
        },
        "location": {
          "city": "Redmond",
          "state": "Washington",
          "countryOrRegion": "US",
          "geoCoordinates": {
            "altitude": null,
            "latitude": 47.68050003051758,
            "longitude": -122.12094116210938
          }
        },
        "appliedConditionalAccessPolicies": [
          {
            "id": "de7e60eb-ed89-4d73-8205-2227def6b7c9",
            "displayName": "SharePoint limited access for guest workers",
            "enforcedGrantControls": [],
            "enforcedSessionControls": [],
            "result": "notEnabled"
          },
          {
            "id": "6701123a-b4c6-48af-8565-565c8bf7cabc",
            "displayName": "Medium signin risk block",
            "enforcedGrantControls": [],
            "enforcedSessionControls": [],
            "result": "notEnabled"
          }
        ]
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_SIGNIN_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'userPrincipalName', 'appId', 'appDisplayName', 'createdDateTime']);
  });

  it('lists all signins in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/auditLogs/signIns`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });
  it('lists all signins by userName in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/auditLogs/signIns?$filter=userPrincipalName eq 'testaccount1%40contoso.com'`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userName: 'testaccount1@contoso.com' } });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('lists all signins by userId in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/auditLogs/signIns?$filter=userId eq '737002f2-9582-4068-b706-044e09481897'`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userId: '737002f2-9582-4068-b706-044e09481897' } });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('lists all signins by appId in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/auditLogs/signIns?$filter=appId eq 'de8bc8b5-d9f9-48b1-a8ad-b748da725064'`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, appId: 'de8bc8b5-d9f9-48b1-a8ad-b748da725064' } });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('lists all signins by appDisplayName in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/auditLogs/signIns?$filter=appDisplayName eq 'Graph%20explorer'`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, appDisplayName: 'Graph explorer' } });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('lists all signins by userName and appId in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/auditLogs/signIns?$filter=userPrincipalName eq 'testaccount1%40contoso.com' and appId eq 'de8bc8b5-d9f9-48b1-a8ad-b748da725064'`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userName: 'testaccount1@contoso.com', appId: 'de8bc8b5-d9f9-48b1-a8ad-b748da725064' } });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('lists all signins by userName and appDisplayName in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/auditLogs/signIns?$filter=userPrincipalName eq 'testaccount1%40contoso.com' and appDisplayName eq 'Graph%20explorer'`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userName: 'testaccount1@contoso.com', appDisplayName: 'Graph explorer' } });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: { userName: 'testaccount1@contoso.com', appDisplayName: 'Graph explorer' } }), new CommandError(errorMessage));
  });

  it('fails validation if userId and userName specified', async () => {
    const actual = await command.validate({ options: { userId: 'de8bc8b5-d9f9-48b1-a8ad-b748da725064', userName: 'Graph explorer' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'not-c49b-4fd4-8223-28f0ac3a6402' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'de8bc8b5-d9f9-48b1-a8ad-b748da725064' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when userName has an invalid value', async () => {
    const actual = await command.validate({ options: { userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and appDisplayName specified', async () => {
    const actual = await command.validate({ options: { appId: 'de8bc8b5-d9f9-48b1-a8ad-b748da725064', appDisplayName: 'Graph explorer' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appId: 'not-c49b-4fd4-8223-28f0ac3a6402' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the appId is a valid GUID', async () => {
    const actual = await command.validate({ options: { appId: 'de8bc8b5-d9f9-48b1-a8ad-b748da725064' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
