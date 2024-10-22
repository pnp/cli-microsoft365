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
import command from './site-admin-add.js';
import { spo } from '../../../../utils/spo.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

describe(commands.SITE_ADMIN_ADD, () => {
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
  const adminToAddId = '10000000-1000-0000-0000-000000000000';
  const adminToAddUPN = 'user3loginName@email.com';
  const primaryAdminLoginName = 'i:0#.f|membership|userPrimaryAdminEmail@email.com';
  const groupId = '00000000-1000-0000-0000-000000000000';
  const groupName = 'TestGroupName';

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
      request.patch,
      entraGroup.getGroupById,
      entraGroup.getGroupByDisplayName,
      entraUser.getUpnByUserId,
      spo.getSiteAdminPropertiesByUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('fails validation if siteUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo', userId: adminToAddId } }, commandInfo);
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
    const actual = await command.validate({ options: { siteUrl: siteUrl, userId: adminToAddId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the siteUrl is a valid SharePoint URL and userName is a valid UPN', async () => {
    const actual = await command.validate({ options: { siteUrl: siteUrl, userName: adminToAddUPN } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds a user to site collection admins by userId as admin', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').resolves(adminToAddUPN);
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userId: adminToAddId, asAdmin: true, verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId, secondaryAdministratorLoginNames:
          ['i:0#.f|membership|user1loginName@email.com', 'i:0#.f|membership|user2loginName@email.com', `i:0#.f|membership|${adminToAddUPN}`]
      }
    });
  });

  it('adds a user as primary site collection admins by userName as admin', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users('user3loginName%40email.com')`) {
        return { userPrincipalName: adminToAddUPN };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')`) {
        return;
      }

      throw `Invalid PATCH request: ${JSON.stringify(opts, null, 2)}`;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userName: adminToAddUPN, asAdmin: true, primary: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId, secondaryAdministratorLoginNames:
          ['i:0#.f|membership|user1loginName@email.com', 'i:0#.f|membership|user2loginName@email.com', `i:0#.f|membership|${adminToAddUPN}`]
      }
    });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      Owner: `i:0#.f|membership|${adminToAddUPN}`,
      SetOwnerWithoutUpdatingSecondaryAdmin: true
    });
  });

  it('adds a group to site collection admin by groupId as admin - for M365 Group', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(entraGroup, 'getGroupById').resolves({
      mail: 'mail',
      id: groupId
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupId: groupId, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId, secondaryAdministratorLoginNames:
          ['i:0#.f|membership|user1loginName@email.com', 'i:0#.f|membership|user2loginName@email.com', `c:0o.c|federateddirectoryclaimprovider|${groupId}`]
      }
    });
  });

  it('adds a group to site collection admin by groupId as admin - for Security Group', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(entraGroup, 'getGroupById').resolves({
      mail: undefined,
      id: groupId
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupId: groupId, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId, secondaryAdministratorLoginNames:
          ['i:0#.f|membership|user1loginName@email.com', 'i:0#.f|membership|user2loginName@email.com', `c:0t.c|tenant|${groupId}`]
      }
    });
  });

  it('adds a group to site collection admin by groupName as admin', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(entraGroup, 'getGroupByDisplayName').resolves(
      {
        mail: undefined,
        id: groupId
      }
    );

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupName: groupName, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId, secondaryAdministratorLoginNames:
          ['i:0#.f|membership|user1loginName@email.com', 'i:0#.f|membership|user2loginName@email.com', `c:0t.c|tenant|${groupId}`]
      }
    });
  });

  it('adds a group as primary site collection admins by userName as admin', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(entraGroup, 'getGroupByDisplayName').resolves(
      {
        mail: undefined,
        id: groupId
      }
    );

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')`) {
        return;
      }

      throw `Invalid PATCH request: ${JSON.stringify(opts, null, 2)}`;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupName: groupName, asAdmin: true, primary: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId, secondaryAdministratorLoginNames:
          ['i:0#.f|membership|user1loginName@email.com', 'i:0#.f|membership|user2loginName@email.com', `c:0t.c|tenant|${groupId}`]
      }
    });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      Owner: `c:0t.c|tenant|${groupId}`,
      SetOwnerWithoutUpdatingSecondaryAdmin: true
    });
  });

  it('adds a user to site collection admins by userId', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').resolves(adminToAddUPN);

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('i%3A0%23.f%7Cmembership%7Cuser3loginName%40email.com')`) {
        return;
      }

      if (opts.url === `${siteUrl}/_api/web/ensureuser` && opts.data.logonName === `i:0#.f|membership|${adminToAddUPN}`) {
        return { LoginName: `i:0#.f|membership|${adminToAddUPN}` };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userId: adminToAddId, verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { IsSiteAdmin: true });
  });

  it('adds a user to site collection admins by userName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('user3loginName%40email.com')`) {
        return { userPrincipalName: adminToAddUPN };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('i%3A0%23.f%7Cmembership%7Cuser3loginName%40email.com')`) {
        return;
      }

      if (opts.url === `${siteUrl}/_api/web/ensureuser` && opts.data.logonName === `i:0#.f|membership|${adminToAddUPN}`) {
        return { LoginName: `i:0#.f|membership|${adminToAddUPN}` };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userName: adminToAddUPN } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { IsSiteAdmin: true });
  });

  it('adds a group to site collection admin by groupId', async () => {
    sinon.stub(entraGroup, 'getGroupById').resolves({
      mail: 'mail',
      id: groupId
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('c%3A0o.c%7Cfederateddirectoryclaimprovider%7C${groupId}')`) {
        return;
      }

      if (opts.url === `${siteUrl}/_api/web/ensureuser` && opts.data.logonName === `c:0o.c|federateddirectoryclaimprovider|${groupId}`) {
        return { LoginName: `c:0o.c|federateddirectoryclaimprovider|${groupId}` };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupId: groupId } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { IsSiteAdmin: true });
  });

  it('adds a group to site collection admin by groupName', async () => {
    sinon.stub(entraGroup, 'getGroupByDisplayName').resolves(
      {
        mail: undefined,
        id: groupId
      }
    );

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('c%3A0t.c%7Ctenant%7C00000000-1000-0000-0000-000000000000')`) {
        return;
      }

      if (opts.url === `${siteUrl}/_api/web/ensureuser` && opts.data.logonName === `c:0t.c|tenant|${groupId}`) {
        return { LoginName: `c:0t.c|tenant|${groupId}` };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupName: groupName } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { IsSiteAdmin: true });
  });

  it('adds a user as primary site collection admins by userId', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').resolves(adminToAddUPN);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/site?$select=Id`) {
        return { Id: siteId };
      }

      if (opts.url === `${siteUrl}/_api/site/owner?$select=LoginName`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      const userId = 5;
      if (opts.url === `${siteUrl}/_api/web/siteusers('i%3A0%23.f%7Cmembership%7Cuser3loginName%40email.com')`) {
        return;
      }

      if (opts.url === `${siteUrl}/_api/web/siteusers('i%3A0%23.f%7Cmembership%7CuserPrimaryAdminEmail%40email.com')`) {
        return;
      }

      if (opts.url === `${siteUrl}/_api/web/ensureuser` && opts.data.logonName === `i:0#.f|membership|${adminToAddUPN}`) {
        return { LoginName: `i:0#.f|membership|${adminToAddUPN}`, Id: userId };
      }

      if (opts.url === `${siteUrl}/_vti_bin/client.svc/ProcessQuery` &&
        opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><SetProperty Id="10" ObjectPathId="2" Name="Owner"><Parameter ObjectPathId="3" /></SetProperty></Actions><ObjectPaths><Property Id="2" ParentId="0" Name="Site" /><Identity Id="3" Name="6d452ba1-40a8-8000-e00d-46e1adaa12bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:u:${userId}" /><StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
      ) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userId: adminToAddId, primary: true } });
    assert.deepStrictEqual(postStub.getCall(postStub.callCount - 1).args[0].data, { IsSiteAdmin: true });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { IsSiteAdmin: true });
  });

  it('correctly handles error when site id is not found for specified site URL in admin mode', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').rejects(new Error(`Cannot get site ${siteUrl}`));
    sinon.stub(entraUser, 'getUpnByUserId').resolves(adminToAddUPN);

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userId: adminToAddId, asAdmin: true } }),
      new CommandError(`Cannot get site ${siteUrl}`));
  });

  it('correctly handles error when user is not found userId admin mode', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').resolves(undefined);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userId: adminToAddId, asAdmin: true } }), new CommandError(`User not found.`));
  });
});