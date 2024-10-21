import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './tenant-site-membership-list.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';
import { CommandError } from '../../../../Command.js';

describe(commands.TENANT_SITE_MEMBERSHIP_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const ownerMembershipList = [
    {
      email: 'owner1Email@email.com',
      loginName: 'i:0#.f|membership|owner1loginName@email.com',
      name: 'owner1DisplayName',
      userPrincipalName: 'owner1loginName'
    },
    {
      email: 'owner2Email@email.com',
      loginName: 'i:0#.f|membership|owner2loginName@email.com',
      name: 'owner2DisplayName',
      userPrincipalName: 'owner2loginName'
    }
  ];
  const membersMembershipList = [
    {
      email: 'member1Email@email.com',
      loginName: 'i:0#.f|membership|member1loginName@email.com',
      name: 'member1DisplayName',
      userPrincipalName: 'member1loginName'
    },
    {
      email: 'member2Email@email.com',
      loginName: 'i:0#.f|membership|member2loginName@email.com',
      name: 'member2DisplayName',
      userPrincipalName: 'member2loginName'
    }
  ];
  const visitorsMembershipList = [
    {
      email: 'visitor1Email@email.com',
      loginName: 'i:0#.f|membership|visitor1loginName@email.com',
      name: 'visitor1DisplayName',
      userPrincipalName: 'visitor1loginName'
    },
    {
      email: 'visitor2Email@email.com',
      loginName: 'i:0#.f|membership|visitor2loginName@email.com',
      name: 'visitor2DisplayName',
      userPrincipalName: 'visitor2loginName'
    }
  ];
  const ownerMembershipListCSVOutput = [
    {
      email: 'owner1Email@email.com',
      loginName: 'i:0#.f|membership|owner1loginName@email.com',
      name: 'owner1DisplayName',
      userPrincipalName: 'owner1loginName',
      associatedGroupType: 'Owner'
    },
    {
      email: 'owner2Email@email.com',
      loginName: 'i:0#.f|membership|owner2loginName@email.com',
      name: 'owner2DisplayName',
      userPrincipalName: 'owner2loginName',
      associatedGroupType: 'Owner'
    }
  ];
  const membersMembershipListCSVOutput = [
    {
      email: 'member1Email@email.com',
      loginName: 'i:0#.f|membership|member1loginName@email.com',
      name: 'member1DisplayName',
      userPrincipalName: 'member1loginName',
      associatedGroupType: 'Member'
    },
    {
      email: 'member2Email@email.com',
      loginName: 'i:0#.f|membership|member2loginName@email.com',
      name: 'member2DisplayName',
      userPrincipalName: 'member2loginName',
      associatedGroupType: 'Member'
    }
  ];
  const visitorsMembershipListCSVOutput = [
    {
      email: 'visitor1Email@email.com',
      loginName: 'i:0#.f|membership|visitor1loginName@email.com',
      name: 'visitor1DisplayName',
      userPrincipalName: 'visitor1loginName',
      associatedGroupType: 'Visitor'
    },
    {
      email: 'visitor2Email@email.com',
      loginName: 'i:0#.f|membership|visitor2loginName@email.com',
      name: 'visitor2DisplayName',
      userPrincipalName: 'visitor2loginName',
      associatedGroupType: 'Visitor'
    }
  ];
  const adminUrl = 'https://contoso-admin.sharepoint.com';
  const siteUrl = 'https://contoso.sharepoint.com/sites/site';
  const siteId = '00000000-0000-0000-0000-000000000010';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getSpoAdminUrl').resolves(adminUrl);
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SITE_MEMBERSHIP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation if the role option is a valid role', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', role: 'Owner' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the siteUrl option is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the role option is not a valid role', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', role: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('lists all site membership groups', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[0,1,2]`) {
        return { value: [{ userGroup: ownerMembershipList }, { userGroup: membersMembershipList }, { userGroup: visitorsMembershipList }] };
      };

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, output: 'json' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], {
      AssociatedOwnerGroup: ownerMembershipList,
      AssociatedMemberGroup: membersMembershipList,
      AssociatedVisitorGroup: visitorsMembershipList
    });
  });

  it('lists all site membership groups - just Owners group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[0]`) {
        return { value: [{ userGroup: ownerMembershipList }] };
      };

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, role: "Owner", output: 'json' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], {
      AssociatedOwnerGroup: ownerMembershipList
    });
  });

  it('lists all site membership groups - just Members group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[1]`) {
        return { value: [{ userGroup: membersMembershipList }] };
      };

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, role: "Member", output: 'json' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], {
      AssociatedMemberGroup: membersMembershipList
    });
  });

  it('lists all site membership groups - just Visitors group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[2]`) {
        return { value: [{ userGroup: visitorsMembershipList }] };
      };

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, role: "Visitor", output: 'json' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], {
      AssociatedVisitorGroup: visitorsMembershipList
    });
  });

  it('lists all site membership groups - csv output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[0,1,2]`) {
        return { value: [{ userGroup: ownerMembershipList }, { userGroup: membersMembershipList }, { userGroup: visitorsMembershipList }] };
      };

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, output: 'csv' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], [
      ...ownerMembershipListCSVOutput,
      ...membersMembershipListCSVOutput,
      ...visitorsMembershipListCSVOutput
    ]);
  });

  it('lists all site membership groups - text output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[0,1,2]`) {
        return { value: [{ userGroup: ownerMembershipList }, { userGroup: membersMembershipList }, { userGroup: visitorsMembershipList }] };
      };

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, output: 'text' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], [
      ...ownerMembershipListCSVOutput,
      ...membersMembershipListCSVOutput,
      ...visitorsMembershipListCSVOutput
    ]);
  });

  it('lists all site membership groups - markdown output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[0,1,2]`) {
        return { value: [{ userGroup: ownerMembershipList }, { userGroup: membersMembershipList }, { userGroup: visitorsMembershipList }] };
      };

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, output: 'md' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], [
      ...ownerMembershipListCSVOutput,
      ...membersMembershipListCSVOutput,
      ...visitorsMembershipListCSVOutput
    ]);
  });

  it('lists all site membership groups - text output when outputs are empty', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[0,1,2]`) {
        return { value: [{ userGroup: [] }, { userGroup: [] }, { userGroup: [] }] };
      };

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, output: 'text' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], []);
  });

  it('lists all site membership groups - just Owners group - csv output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[0]`) {
        return { value: [{ userGroup: ownerMembershipList }] };
      };

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, role: "Owner", output: 'csv' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], [
      ...ownerMembershipListCSVOutput
    ]);
  });

  it('lists all site membership groups - just Members group - csv output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[1]`) {
        return { value: [{ userGroup: membersMembershipList }] };
      };

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, role: "Member", output: 'csv' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], [
      ...membersMembershipListCSVOutput
    ]);
  });

  it('lists all site membership groups - just Visitors group - csv output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[2]`) {
        return { value: [{ userGroup: visitorsMembershipList }] };
      };

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, role: "Visitor", output: 'csv' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], [
      ...visitorsMembershipListCSVOutput
    ]);
  });

  it('correctly handles error when site is not found for specified site URL', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: "-1, Microsoft.Online.SharePoint.Common.SpoNoSiteException", message: { lang: "en-US", value: `Cannot get site ${siteUrl}.` }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, verbose: true } }), new CommandError(`Cannot get site ${siteUrl}.`));
  });
});