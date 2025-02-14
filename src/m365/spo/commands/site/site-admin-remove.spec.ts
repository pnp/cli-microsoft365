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
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

describe(commands.SITE_ADMIN_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    sinon.stub(telemetry, 'trackEvent').resolves();
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      entraGroup.getGroupById,
      entraGroup.getGroupByDisplayName,
      entraUser.getUpnByUserId,
      cli.getSettingWithDefaultValue,
      cli.promptForConfirmation,
      spo.getSiteAdminPropertiesByUrl
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

  it('prompts for confirmation before removing site admin if --force option is not passed for user option', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    const promptStub = sinon.stub(cli, 'promptForConfirmation').callsFake(async opts => {
      if (opts.message === `Are you sure you want to remove specified user from the site administrators list ${siteUrl}?`) {
        return true;
      }

      throw 'Invalid request: ' + opts.message;
    });

    sinon.stub(entraUser, 'getUpnByUserId').resolves(adminToRemoveUPN);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return { OwnerLoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return {
          value: listOfAdminsFromAdminSource
        };
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, asAdmin: true } });
    assert(promptStub.calledOnce);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId,
        secondaryAdministratorLoginNames: [
          'i:0#.f|membership|user2loginName@email.com'
        ]
      }
    });
  });

  it('prompts for confirmation before removing site admin if --force option is not passed for group option', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    const promptStub = sinon.stub(cli, 'promptForConfirmation').callsFake(async opts => {
      if (opts.message === `Are you sure you want to remove specified group from the site administrators list ${siteUrl}?`) {
        return true;
      }

      throw 'Invalid request: ' + opts.message;
    });

    sinon.stub(entraGroup, 'getGroupById').resolves({ id: groupId, mail: 'mail' });
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return { OwnerLoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return {
          value: [
            ...listOfAdminsFromAdminSource,
            {
              email: '',
              loginName: `c:0o.c|federateddirectoryclaimprovider|${groupId}`,
              name: groupName,
              userPrincipalName: ''
            }
          ]
        };
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupId: groupId, asAdmin: true } });
    assert(promptStub.calledOnce);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId,
        secondaryAdministratorLoginNames: [
          'i:0#.f|membership|user1loginName@email.com',
          'i:0#.f|membership|user2loginName@email.com'
        ]
      }
    });
  });

  it('aborts removing site admin when confirmation is not given', async () => {
    const promptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { siteUrl: siteUrl, userName: adminToRemoveUPN } });
    assert(promptStub.calledOnce);
  });

  it('removes a user from site collection admins by userId as admin', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(entraUser, 'getUpnByUserId').resolves(adminToRemoveUPN);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return { OwnerLoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return {
          value: listOfAdminsFromAdminSource
        };
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, force: true, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId,
        secondaryAdministratorLoginNames: [
          'i:0#.f|membership|user2loginName@email.com'
        ]
      }
    });
  });

  it('removes a user from site collection admins by userId as admin with verbose parameter', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(entraUser, 'getUpnByUserId').resolves(adminToRemoveUPN);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return { OwnerLoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return {
          value: listOfAdminsFromAdminSource
        };
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators` && opts.data.secondaryAdministratorsFieldsData) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, force: true, asAdmin: true, verbose: true } });
    assert(loggerLogToStderrSpy.calledWith(`Removing site administrator as an administrator...`));
  });

  it('removes a user from site collection admins by userName as admin', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return { OwnerLoginName: primaryAdminLoginName };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users('user1loginName%40email.com')`) {
        return { userPrincipalName: adminToRemoveUPN };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return {
          value: listOfAdminsFromAdminSource
        };
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userName: adminToRemoveUPN, force: true, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId,
        secondaryAdministratorLoginNames: [
          'i:0#.f|membership|user2loginName@email.com'
        ]
      }
    });
  });

  it('correctly handles an error if trying to remove primary site collection administrator as admin by userId', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(entraUser, 'getUpnByUserId').resolves(primaryAdminUPN);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return { OwnerLoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userId: primaryAdminId, force: true, asAdmin: true } }), new CommandError('You cannot remove the primary site collection administrator.'));
  });

  it('correctly handles an error if trying to remove primary site collection administrator as admin by userName', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return { OwnerLoginName: primaryAdminLoginName };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users('userPrimaryAdminEmail%40email.com')`) {
        return { userPrincipalName: primaryAdminUPN };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userName: primaryAdminUPN, force: true, asAdmin: true } }), new CommandError('You cannot remove the primary site collection administrator.'));
  });

  it('removes a group from site collection admin by groupId as admin - for M365 Group', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(entraGroup, 'getGroupById').resolves({ id: groupId, mail: 'mail' });
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return { OwnerLoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return {
          value: [
            ...listOfAdminsFromAdminSource,
            {
              email: '',
              loginName: `c:0o.c|federateddirectoryclaimprovider|${groupId}`,
              name: groupName,
              userPrincipalName: ''
            }
          ]
        };
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupId: groupId, force: true, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId,
        secondaryAdministratorLoginNames: [
          'i:0#.f|membership|user1loginName@email.com',
          'i:0#.f|membership|user2loginName@email.com'
        ]
      }
    });
  });

  it('removes a group from site collection admin by groupId as admin - for Security Group', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(entraGroup, 'getGroupById').resolves({ id: groupId, mail: undefined });
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return { OwnerLoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return {
          value: [
            ...listOfAdminsFromAdminSource,
            {
              email: '',
              loginName: `c:0t.c|tenant|${groupId}`,
              name: groupName,
              userPrincipalName: ''
            }
          ]
        };
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupId: groupId, force: true, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId,
        secondaryAdministratorLoginNames: [
          'i:0#.f|membership|user1loginName@email.com',
          'i:0#.f|membership|user2loginName@email.com'
        ]
      }
    });
  });

  it('removes a group from site collection admin by groupName as admin', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    sinon.stub(entraGroup, 'getGroupByDisplayName').resolves({ id: groupId, mail: undefined });
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return { OwnerLoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return {
          value: [
            ...listOfAdminsFromAdminSource,
            {
              email: '',
              loginName: `c:0t.c|tenant|${groupId}`,
              name: groupName,
              userPrincipalName: ''
            }
          ]
        };
      }

      if (opts.url === `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupName: groupName, force: true, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      secondaryAdministratorsFieldsData: {
        siteId: siteId,
        secondaryAdministratorLoginNames: [
          'i:0#.f|membership|user1loginName@email.com',
          'i:0#.f|membership|user2loginName@email.com'
        ]
      }
    });
  });

  it('removes a user from site collection admins by userId', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').resolves(adminToRemoveUPN);
    sinon.stub(request, 'get').callsFake(async opts => {
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

  it('removes a user from site collection admins by userName with verbose parameter', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('user1loginName%40email.com')`) {
        return { userPrincipalName: adminToRemoveUPN };
      }

      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('i%3A0%23.f%7Cmembership%7Cuser1loginName%40email.com')`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, userName: adminToRemoveUPN, force: true, verbose: true } });
    assert(loggerLogToStderrSpy.calledWith(`Removing site administrator...`));
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      IsSiteAdmin: false
    });
  });

  it('correctly handles an error if trying to remove primary site collection administrator by userId', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').resolves(primaryAdminUPN);
    sinon.stub(request, 'get').callsFake(async opts => {
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
    sinon.stub(entraGroup, 'getGroupById').resolves({ id: groupId, mail: 'mail' });
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('c%3A0o.c%7Cfederateddirectoryclaimprovider%7C${groupId}')`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupId: groupId, force: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      IsSiteAdmin: false
    });
  });

  it('removes a group from site collection admin by groupName', async () => {
    sinon.stub(entraGroup, 'getGroupByDisplayName').resolves({ id: groupId, mail: undefined });
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers('c%3A0t.c%7Ctenant%7C00000000-1000-0000-0000-000000000000')`) {
        return;
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, groupName: groupName, force: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      IsSiteAdmin: false
    });
  });

  it('correctly handles incorrect site Id guid in admin mode', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').resolves(adminToRemoveUPN);
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').rejects(new Error(`Cannot get site ${siteUrl}`));

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, asAdmin: true, force: true } }), new CommandError(`Cannot get site ${siteUrl}`));
  });

  it('correctly handles error when user is not found userId admin mode', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').resolves(undefined);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, userId: adminToRemoveId, asAdmin: true, force: true } }), new CommandError(`User not found.`));
  });
});
