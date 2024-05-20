import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './site-admin-remove.js';
import { spo } from '../../../../utils/spo.js';
import { CommandError } from '../../../../Command.js';

describe(commands.SITE_ADMIN_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const listOfAdminsFromAdminSource = [
    {
      email: 'user1Email@email.com',
      loginName: 'i:0#.f|membership|user1loginName@email.com',
      name: 'user1DisplayName',
      userPrincipalName: 'user1loginName'
    },
    {
      email: 'user2Email@email.com',
      loginName: 'i:0#.f|membership|user2loginName@email.com',
      name: 'user2DisplayName',
      userPrincipalName: 'user2loginName'
    }
  ];
  const rootUrl = 'https://contoso.sharepoint.com';
  const adminUrl = 'https://contoso-admin.sharepoint.com';
  const siteUrl = 'https://contoso.sharepoint.com/sites/site';
  const siteId = '00000000-0000-0000-0000-000000000010';
  const primaryAdminId = '00000000-0000-0000-0000-000000000000';
  const primaryAdminUPN = 'userPrimaryAdminEmail@email.com';
  const primaryAdminLoginName = 'i:0#.f|membership|userPrimaryAdminEmail@email.com';
  const adminToRemoveId = '10000000-1000-0000-0000-000000000000';
  const adminToRemoveUPN = 'user1loginName@email.com';
  const groupId = '00000000-1000-0000-0000-000000000000';
  const groupName = 'TestGroup';


  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('fails validation if siteUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo', userId: adminToRemoveId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { siteUrl: siteUrl, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { siteUrl: siteUrl, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { siteUrl: siteUrl, groupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the siteUrl is a valid SharePoint URL and userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { siteUrl: siteUrl, userId: adminToRemoveId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the siteUrl is a valid SharePoint URL and userName is a valid UPN', async () => {
    const actual = await command.validate({ options: { siteUrl: siteUrl, userName: adminToRemoveUPN } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts for confirmation before removing site admin if --force option is not passed', async () => {
    const promptStub = sinon.stub(cli, 'promptForConfirmation').resolves(true);
    let result = false;
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites(%27${siteId}%27)?$select=OwnerLoginName`) {
        return JSON.stringify({ OwnerLoginName: primaryAdminLoginName });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users/${adminToRemoveId}`) {
        return { userPrincipalName: adminToRemoveUPN };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId=%27${siteId}%27`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators` && opts.data.secondaryAdministratorsFieldsData) {
        result = opts.data.secondaryAdministratorsFieldsData.siteId === siteId &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames.length === 1 &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames[0] === 'i:0#.f|membership|user2loginName@email.com';
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, asAdmin: true } });
    assert(promptStub.calledOnce);
    assert(result);
  });

  it('aborts removing site admin when confirmation is not given', async () => {
    const promptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { siteUrl: siteUrl, userName: adminToRemoveUPN } });
    assert(promptStub.calledOnce);
  });

  it('removes a user from site collection admins by userId as admin', async () => {
    let result = false;
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites(%27${siteId}%27)?$select=OwnerLoginName`) {
        return JSON.stringify({ OwnerLoginName: primaryAdminLoginName });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users/${adminToRemoveId}`) {
        return { userPrincipalName: adminToRemoveUPN };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId=%27${siteId}%27`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators` && opts.data.secondaryAdministratorsFieldsData) {
        result = opts.data.secondaryAdministratorsFieldsData.siteId === siteId &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames.length === 1 &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames[0] === 'i:0#.f|membership|user2loginName@email.com';
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, force: true, asAdmin: true } });
    assert(result);
  });

  it('removes a user from site collection admins by userName as admin', async () => {
    let result = false;
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites(%27${siteId}%27)?$select=OwnerLoginName`) {
        return JSON.stringify({ OwnerLoginName: primaryAdminLoginName });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users('user1loginName%40email.com')`) {
        return { userPrincipalName: adminToRemoveUPN };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId=%27${siteId}%27`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators` && opts.data.secondaryAdministratorsFieldsData) {
        result = opts.data.secondaryAdministratorsFieldsData.siteId === siteId &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames.length === 1 &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames[0] === 'i:0#.f|membership|user2loginName@email.com';
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userName: adminToRemoveUPN, force: true, asAdmin: true } });
    assert(result);
  });

  it('correctly handles an error if trying to remove primary site collection administrator as admin by userId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites(%27${siteId}%27)?$select=OwnerLoginName`) {
        return JSON.stringify({ OwnerLoginName: primaryAdminLoginName });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users/${primaryAdminId}`) {
        return { userPrincipalName: primaryAdminUPN };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userId: primaryAdminId, force: true, asAdmin: true } }), new CommandError('You cannot remove the primary site collection administrator.'));
  });

  it('correctly handles an error if trying to remove primary site collection administrator as admin by userName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites(%27${siteId}%27)?$select=OwnerLoginName`) {
        return JSON.stringify({ OwnerLoginName: primaryAdminLoginName });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users('userPrimaryAdminEmail%40email.com')`) {
        return { userPrincipalName: primaryAdminUPN };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userName: primaryAdminUPN, force: true, asAdmin: true } }), new CommandError('You cannot remove the primary site collection administrator.'));
  });

  it('removes a group from site collection admin by groupId as admin - for M365 Group', async () => {
    let result = false;
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites(%27${siteId}%27)?$select=OwnerLoginName`) {
        return JSON.stringify({ OwnerLoginName: primaryAdminLoginName });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return {
          mail: 'mail',
          id: groupId
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId=%27${siteId}%27`) {
        return JSON.stringify({
          value: [
            ...listOfAdminsFromAdminSource,
            {
              email: '',
              loginName: `c:0o.c|federateddirectoryclaimprovider|${groupId}`,
              name: groupName,
              userPrincipalName: ''
            }
          ]
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators` && opts.data.secondaryAdministratorsFieldsData) {
        result = opts.data.secondaryAdministratorsFieldsData.siteId === siteId &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames.length === 2 &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames[0] === 'i:0#.f|membership|user1loginName@email.com' &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames[1] === 'i:0#.f|membership|user2loginName@email.com';
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupId: groupId, force: true, asAdmin: true } });
    assert(result);
  });

  it('removes a group from site collection admin by groupId as admin - for Security Group', async () => {
    let result = false;
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites(%27${siteId}%27)?$select=OwnerLoginName`) {
        return JSON.stringify({ OwnerLoginName: primaryAdminLoginName });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return {
          mail: undefined,
          id: groupId
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId=%27${siteId}%27`) {
        return JSON.stringify({
          value: [
            ...listOfAdminsFromAdminSource,
            {
              email: '',
              loginName: `c:0t.c|tenant|${groupId}`,
              name: groupName,
              userPrincipalName: ''
            }
          ]
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators` && opts.data.secondaryAdministratorsFieldsData) {
        result = opts.data.secondaryAdministratorsFieldsData.siteId === siteId &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames.length === 2 &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames[0] === 'i:0#.f|membership|user1loginName@email.com' &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames[1] === 'i:0#.f|membership|user2loginName@email.com';
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupId: groupId, force: true, asAdmin: true } });
    assert(result);
  });

  it('removes a group from site collection admin by groupName as admin', async () => {
    let result = false;
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites(%27${siteId}%27)?$select=OwnerLoginName`) {
        return JSON.stringify({ OwnerLoginName: primaryAdminLoginName });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${groupName}'`) {
        return {
          value: [{
            mail: undefined,
            id: groupId
          }]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId=%27${siteId}%27`) {
        return JSON.stringify({
          value: [
            ...listOfAdminsFromAdminSource,
            {
              email: '',
              loginName: `c:0t.c|tenant|${groupId}`,
              name: groupName,
              userPrincipalName: ''
            }
          ]
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators` && opts.data.secondaryAdministratorsFieldsData) {
        result = opts.data.secondaryAdministratorsFieldsData.siteId === siteId &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames.length === 2 &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames[0] === 'i:0#.f|membership|user1loginName@email.com' &&
          opts.data.secondaryAdministratorsFieldsData.secondaryAdministratorLoginNames[1] === 'i:0#.f|membership|user2loginName@email.com';
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupName: groupName, force: true, asAdmin: true } });
    assert(result);
  });

  it('removes a user from site collection admins by userId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${adminToRemoveId}`) {
        return { userPrincipalName: adminToRemoveUPN };
      }

      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('i%3A0%23.f%7Cmembership%7Cuser1loginName%40email.com')` && opts.data.IsSiteAdmin === false) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, force: true } });
  });

  it('removes a user from site collection admins by userName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('user1loginName%40email.com')`) {
        return { userPrincipalName: adminToRemoveUPN };
      }

      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('i%3A0%23.f%7Cmembership%7Cuser1loginName%40email.com')` && opts.data.IsSiteAdmin === false) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userName: adminToRemoveUPN, force: true } });
  });

  it('correctly handles an error if trying to remove primary site collection administrator by userId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${primaryAdminId}`) {
        return { userPrincipalName: primaryAdminUPN };
      }

      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('i%3A0%23.f%7Cmembership%7Cuser1loginName%40email.com')` && opts.data.IsSiteAdmin === false) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userId: primaryAdminId, force: true } }), new CommandError('You cannot remove the primary site collection administrator.'));
  });

  it('removes a group from site collection admin by groupId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return {
          mail: 'mail',
          id: groupId
        };
      }

      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('c%3A0o.c%7Cfederateddirectoryclaimprovider%7C${groupId}')` && opts.data.IsSiteAdmin === false) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupId: groupId, force: true } });
  });

  it('removes a group from site collection admin by groupName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${groupName}'`) {
        return {
          value: [{
            mail: undefined,
            id: groupId
          }]
        };
      }

      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('c%3A0t.c%7Ctenant%7C00000000-1000-0000-0000-000000000000')` && opts.data.IsSiteAdmin === false) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupName: groupName, force: true } });
  });

  it('correctly handles incorrect site Id guid in admin mode', async () => {
    const incorrectSiteId = 'foo';
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users/${adminToRemoveId}`) {
        return { userPrincipalName: adminToRemoveUPN };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId=%27${incorrectSiteId}%27`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }
      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${incorrectSiteId})/` }
            ]
          }
        ]
        );
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, asAdmin: true, force: true } }), new CommandError('Failed to obtain site Id from the provided site URL.'));
  });

  it('correctly handles error when site id is not found for specified site URL in admin mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users/${adminToRemoveId}`) {
        return { userPrincipalName: adminToRemoveUPN };
      }
      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": {
              "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "d2d0389e-a040-4000-b24b-d16b0546a03c", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
            }, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }
        ]
        );
      }
      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, asAdmin: true, force: true } }), new CommandError('Unknown Error'));
  });

  it('correctly handles error when site is not found for specified site URL admin mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users/${adminToRemoveId}`) {
        return { userPrincipalName: adminToRemoveUPN };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": []
          }
        ]
        );
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, asAdmin: true, force: true } }), new CommandError(`Failed to obtain site Id from the provided site URL.`));
  });

  it('correctly handles error when user is not found userId admin mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users/${adminToRemoveId}`) {
        return null;
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": []
          }
        ]
        );
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, asAdmin: true, force: true } }), new CommandError(`User not found.`));
  });
});
