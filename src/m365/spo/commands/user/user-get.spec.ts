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
import command from './user-get.js';
import { settingsNames } from '../../../../settingsNames.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.USER_GET, () => {
  const validUserName = 'john.doe_hotmail.com#ext#@contoso.onmicrosoft.com';
  const validEmail = 'john.doe@contoso.onmicrosoft.com';
  const validEntraGroupId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
  const validEntraGroupName = 'Finance';
  const validEntraSecurityGroupName = 'EntraGroupTest';
  const validLoginName = `i:0#.f|membership|${validUserName}`;
  const validWebUrl = 'https://contoso.sharepoint.com/subsite';

  const userResponse = {
    "Id": 10,
    "IsHiddenInUI": false,
    "LoginName": validLoginName,
    "Title": "John Doe",
    "PrincipalType": 1,
    "Email": validEmail,
    "Expiration": "",
    "IsEmailAuthenticationGuestUser": false,
    "IsShareByEmailGuestUser": false,
    "IsSiteAdmin": false,
    "UserId": { "NameId": "10010001b0c19a2", "NameIdIssuer": "urn:federation:microsoftonline" },
    "UserPrincipalName": validUserName
  };

  const groupM365Response = {
    value: [{
      "id": "2056d2f6-3257-4253-8cfc-b73393e414e5",
      "deletedDateTime": null,
      "classification": null,
      "createdDateTime": "2017-11-29T03:27:05Z",
      "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
      "displayName": "Finance",
      "groupTypes": [
        "Unified"
      ],
      "mail": "finance@contoso.onmicrosoft.com",
      "mailEnabled": true,
      "mailNickname": "finance",
      "onPremisesLastSyncDateTime": null,
      "onPremisesProvisioningErrors": [],
      "onPremisesSecurityIdentifier": null,
      "onPremisesSyncEnabled": null,
      "preferredDataLocation": null,
      "proxyAddresses": [
        "SMTP:finance@contoso.onmicrosoft.com"
      ],
      "renewedDateTime": "2017-11-29T03:27:05Z",
      "securityEnabled": false,
      "visibility": "Public"
    }]
  };

  const groupSecurityResponse = {
    value: [{
      "id": "2056d2f6-3257-4253-8cfc-b73393e414e5",
      "deletedDateTime": null,
      "classification": null,
      "createdDateTime": "2024-01-27T16:02:56Z",
      "creationOptions": [],
      "description": "Entra Group Test",
      "displayName": "EntraGroupTest",
      "expirationDateTime": null,
      "groupTypes": [],
      "isAssignableToRole": true,
      "mail": null,
      "mailEnabled": false,
      "mailNickname": "f45205a2-d",
      "membershipRule": null,
      "membershipRuleProcessingState": null,
      "onPremisesDomainName": null,
      "onPremisesLastSyncDateTime": null,
      "onPremisesNetBiosName": null,
      "onPremisesSamAccountName": null,
      "onPremisesSecurityIdentifier": null,
      "onPremisesSyncEnabled": null,
      "preferredDataLocation": null,
      "preferredLanguage": null,
      "proxyAddresses": [],
      "renewedDateTime": "2024-01-27T16:02:56Z",
      "resourceBehaviorOptions": [],
      "resourceProvisioningOptions": [],
      "securityEnabled": true,
      "securityIdentifier": "S-1-12-1-1968173404-1154184881-1694549896-3083850660",
      "theme": null,
      "visibility": "Private",
      "onPremisesProvisioningErrors": [],
      "serviceProvisioningErrors": []
    }]
  };

  let log: any[];
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
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves user by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/siteusers/GetById('10')`) {
        return userResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: validWebUrl,
        id: 10
      }
    });

    assert(loggerLogSpy.calledWith(userResponse));
  });

  it('retrieves user by email', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/siteusers/GetByEmail('${formatting.encodeQueryParameter(validEmail)}')`) {
        return userResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: validWebUrl,
        email: validEmail
      }
    });

    assert(loggerLogSpy.calledWith(userResponse));
  });

  it('retrieves user by loginName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/siteusers/GetByLoginName('${formatting.encodeQueryParameter(validLoginName)}')`) {
        return userResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: validWebUrl,
        loginName: validLoginName
      }
    });

    assert(loggerLogSpy.calledWith(userResponse));
  });

  it('retrieves user by userName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/siteusers?$filter=UserPrincipalName eq ('${formatting.encodeQueryParameter(validUserName)}')`) {
        return {
          "value": [userResponse]
        };
      }

      if (opts.url === `${validWebUrl}/_api/web/siteusers/GetById('10')`) {
        return userResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: validWebUrl,
        userName: validUserName
      }
    });

    assert(loggerLogSpy.calledWith(userResponse));
  });

  it('retrieves m365 group by entraGroupId for mail enabled group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validEntraGroupId}`) {
        return groupM365Response.value[0];
      }

      if (opts.url === `${validWebUrl}/_api/web/siteusers/GetByEmail('finance%40contoso.onmicrosoft.com')`) {
        return {
          "Id": 45,
          "IsHiddenInUI": false,
          "LoginName": "c:0o.c|federateddirectoryclaimprovider|2056d2f6-3257-4253-8cfc-b73393e414e5",
          "Title": "Finance",
          "PrincipalType": 4,
          "Email": "finance@contoso.onmicrosoft.com",
          "Expiration": "",
          "IsEmailAuthenticationGuestUser": false,
          "IsShareByEmailGuestUser": false,
          "IsSiteAdmin": false,
          "UserId": null,
          "UserPrincipalName": null
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: validWebUrl,
        entraGroupId: validEntraGroupId
      }
    });

    assert(loggerLogSpy.calledWith({
      "Id": 45,
      "IsHiddenInUI": false,
      "LoginName": "c:0o.c|federateddirectoryclaimprovider|2056d2f6-3257-4253-8cfc-b73393e414e5",
      "Title": "Finance",
      "PrincipalType": 4,
      "Email": "finance@contoso.onmicrosoft.com",
      "Expiration": "",
      "IsEmailAuthenticationGuestUser": false,
      "IsShareByEmailGuestUser": false,
      "IsSiteAdmin": false,
      "UserId": null,
      "UserPrincipalName": null
    }));
  });

  it('retrieves security group by entraGroupName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${validEntraSecurityGroupName}'`) {
        return groupSecurityResponse;
      }

      if (opts.url === `${validWebUrl}/_api/web/siteusers/GetByLoginName('c:0t.c|tenant|${validEntraGroupId}')`) {
        return {
          "Id": 31,
          "IsHiddenInUI": false,
          "LoginName": "c:0t.c|tenant|2056d2f6-3257-4253-8cfc-b73393e414e5",
          "Title": "EntraGroupTest",
          "PrincipalType": 4,
          "Email": "",
          "Expiration": "",
          "IsEmailAuthenticationGuestUser": false,
          "IsShareByEmailGuestUser": false,
          "IsSiteAdmin": false,
          "UserId": null,
          "UserPrincipalName": null
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        webUrl: validWebUrl,
        entraGroupName: validEntraSecurityGroupName
      }
    } as any);

    assert(loggerLogSpy.calledWith({
      "Id": 31,
      "IsHiddenInUI": false,
      "LoginName": "c:0t.c|tenant|2056d2f6-3257-4253-8cfc-b73393e414e5",
      "Title": "EntraGroupTest",
      "PrincipalType": 4,
      "Email": "",
      "Expiration": "",
      "IsEmailAuthenticationGuestUser": false,
      "IsShareByEmailGuestUser": false,
      "IsSiteAdmin": false,
      "UserId": null,
      "UserPrincipalName": null
    }));
  });

  it('retrieves current logged-in user', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/currentuser') {
        return {
          "Id": 6,
          "IsHiddenInUI": false,
          "LoginName": "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
          "Title": "John Doe",
          "PrincipalType": 1,
          "Email": "john.doe@mytenant.onmicrosoft.com",
          "Expiration": "",
          "IsEmailAuthenticationGuestUser": false,
          "IsShareByEmailGuestUser": false,
          "IsSiteAdmin": false,
          "UserId": { "NameId": "10010001b0c19a2", "NameIdIssuer": "urn:federation:microsoftonline" },
          "UserPrincipalName": "john.doe@mytenant.onmicrosoft.com"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com'
      }
    });

    assert(loggerLogSpy.calledWith({
      Id: 6,
      IsHiddenInUI: false,
      LoginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
      Title: "John Doe",
      PrincipalType: 1,
      Email: "john.doe@mytenant.onmicrosoft.com",
      Expiration: "",
      IsEmailAuthenticationGuestUser: false,
      IsShareByEmailGuestUser: false,
      IsSiteAdmin: false,
      UserId: { NameId: "10010001b0c19a2", NameIdIssuer: "urn:federation:microsoftonline" },
      UserPrincipalName: "john.doe@mytenant.onmicrosoft.com"
    }));
  });

  it('handles generic error when user not found when username is passed', async () => {
    const err = `User not found: ${validUserName}`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/siteusers?$filter=UserPrincipalName eq ('${formatting.encodeQueryParameter(validUserName)}')`) {
        return { "value": [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        userName: validUserName
      }
    }), new CommandError(err));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if entraGroupId is not a valid id', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, entraGroupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a valid number', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid user principal name', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if email is not a valid user principal name', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, email: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id, email, loginName, userName, entraGroupId, and entraGroupName options are passed (multiple options)', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: validWebUrl, id: 1, email: validEmail, loginName: validLoginName, userName: validUserName, entraGroupId: validEntraGroupId, entraGroupName: validEntraGroupName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation url is valid and id is passed', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, id: 1 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and email is passed', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, email: validEmail } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and loginName is passed', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, loginName: validLoginName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and userName is passed', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, userName: validUserName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and entraGroupName is passed', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, entraGroupName: validEntraGroupName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and entraGroupId is passed', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, entraGroupId: validEntraGroupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});