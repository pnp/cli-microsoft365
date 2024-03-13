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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './site-admin-list.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.SITE_ADMIN_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const primaryAdminLoginName = 'user1loginName';
  const listOfAdminsResultRegularMode = [
    {
      Id: 1,
      LoginName: 'user1loginName',
      Title: 'user1DisplayName',
      PrincipalType: 1,
      PrincipalTypeString: 'User',
      IsPrimaryAdmin: true,
      Email: 'user1Email@email.com'
    },
    {
      Id: 2,
      LoginName: 'user2loginName',
      Title: 'user2DisplayName',
      PrincipalType: 1,
      PrincipalTypeString: 'User',
      IsPrimaryAdmin: false,
      Email: 'user2Email@email.com'
    }
  ];

  const listOfAdminsResultAsAdmin = [
    {
      Id: null,
      LoginName: 'user1loginName',
      Title: 'user1DisplayName',
      PrincipalType: null,
      PrincipalTypeString: null,
      IsPrimaryAdmin: true,
      Email: 'user1Email@email.com'
    },
    {
      Id: null,
      LoginName: 'user2loginName',
      Title: 'user2DisplayName',
      PrincipalType: null,
      PrincipalTypeString: null,
      IsPrimaryAdmin: false,
      Email: 'user2Email@email.com'
    }
  ];

  const listOfAdminsFromSiteSource = [
    {
      Email: 'user1Email@email.com',
      Id: 1,
      IsSiteAdmin: true,
      LoginName: 'user1loginName',
      PrincipalType: 1,
      Title: 'user1DisplayName'
    },
    {
      Email: 'user2Email@email.com',
      Id: 2,
      IsSiteAdmin: true,
      LoginName: 'user2loginName',
      PrincipalType: 1,
      Title: 'user2DisplayName'
    }
  ];

  const listOfAdminsFromAdminSource = [
    {
      email: 'user1Email@email.com',
      loginName: 'user1loginName',
      name: 'user1DisplayName',
      userPrincipalName: 'user1loginName'
    },
    {
      email: 'user2Email@email.com',
      loginName: 'user2loginName',
      name: 'user2DisplayName',
      userPrincipalName: 'user2loginName'
    }
  ];
  const rootUrl = 'https://contoso.sharepoint.com';
  const siteUrl = 'https://contoso.sharepoint.com/sites/site';
  const adminUrl = 'https://contoso-admin.sharepoint.com';
  const siteId = '00000000-0000-0000-0000-000000000000';

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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_ADMIN_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets site collection admins in regular mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers?$filter=IsSiteAdmin eq true`) {
        return { value: listOfAdminsFromSiteSource };
      }

      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl } });
    assert(loggerLogSpy.calledOnceWithExactly(listOfAdminsResultRegularMode));
  });

  it('gets site collection admins in admin mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return JSON.stringify({ OwnerLoginName: primaryAdminLoginName });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/site?$select=id`) {
        return { id: `contoso.sharepoint.com,${siteId},fb0a066f-c10f-4734-94d1-f896de4aa484` };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, asAdmin: true } });
    assert(loggerLogSpy.calledOnceWithExactly(listOfAdminsResultAsAdmin));
  });

  it('correctly handles empty list of site collection admins from API in regular mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers?$filter=IsSiteAdmin eq true`) {
        return { value: [] };
      }

      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return null;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl } });
    assert(loggerLogSpy.calledOnceWithExactly([]));
  });

  it('correctly handles errors from API in regular mode', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl } }), new CommandError('An error has occurred'));
  });

  it('correctly handles empty list of site collection admins from API in admin mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return JSON.stringify({ OwnerLoginName: '' });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/site?$select=id`) {
        return { id: `contoso.sharepoint.com,${siteId},fb0a066f-c10f-4734-94d1-f896de4aa484` };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: []
        });
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, asAdmin: true } });
    assert(loggerLogSpy.calledOnceWithExactly([]));
  });

  it('handles error when primary admin API returns error in regular mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers?$filter=IsSiteAdmin eq true`) {
        return { value: listOfAdminsFromSiteSource };
      }

      if (opts.url === `${siteUrl}/_api/site/owner`) {
        throw "Invalid request";
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl } }), new CommandError('Invalid request'));
  });

  it('handles error when primary admin API returns error in admin mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        throw "Invalid request";
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/site?$select=id`) {
        return { id: `contoso.sharepoint.com,${siteId},fb0a066f-c10f-4734-94d1-f896de4aa484` };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }
      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, asAdmin: true } }), new CommandError('Invalid request'));
  });

  it('handles error when returned siteId is incorrect in admin mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        throw "Invalid request";
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/site?$select=id`) {
        return { id: 'Incorrect ID' };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }
      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, asAdmin: true } }), new CommandError(`Site with URL ${siteUrl} not found`));
  });

  it('passes validation when only correct siteUrl option specified', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/site'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when correct siteUrl and correct asAdmin options specified', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/site',
        asAdmin: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the url option not specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }
      return defaultValue;
    });

    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.notDeepEqual(actual, true);
  });

  it('fails validation when the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('get additional log when verbose parameter is set in regular mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/web/siteusers?$filter=IsSiteAdmin eq true`) {
        return { value: listOfAdminsFromSiteSource };
      }

      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, verbose: true } });
    assert(loggerLogToStderrSpy.firstCall.firstArg === 'Retrieving site administrators...');
  });

  it('get additional log when verbose parameter is set in admin mode', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`) {
        return { res: { webUrl: rootUrl } };
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }

      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return JSON.stringify({ OwnerLoginName: primaryAdminLoginName });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/site?$select=id`) {
        return { id: `contoso.sharepoint.com,${siteId},fb0a066f-c10f-4734-94d1-f896de4aa484` };
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`) {
        return JSON.stringify({
          value: listOfAdminsFromAdminSource
        });
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: siteUrl, asAdmin: true, verbose: true } });
    assert(loggerLogToStderrSpy.firstCall.firstArg === 'Retrieving site administrators as an administrator...');
  });
});
