import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { formatting } from '../../../../utils/formatting.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './user-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.USER_REMOVE, () => {
  const validUserName = 'john.deo_hotmail.com#ext#@contoso.onmicrosoft.com';
  const validEmail = 'john.deo@contoso.onmicrosoft.com';
  const validEntraGroupId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
  const validEntraM365GroupName = 'Finance';
  const validEntraSecurityGroupName = 'EntraGroupTest';
  const validLoginName = `i:0#.f|membership|${validUserName}`;
  const validWebUrl = 'https://contoso.sharepoint.com/subsite';
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

  let log: any[];
  let requests: any[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
    requests = [];
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      cli.promptForConfirmation,
      spo.getUserByEmail,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id or loginName or userName or email or entraGroupName or entraGroupId options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }
      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        webUrl: validWebUrl
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if more than one of the options userName or email or entraGroupName or entraGroupId are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        webUrl: validWebUrl,
        id: 10,
        loginName: validLoginName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'foo',
        id: 10
      }
    }, commandInfo);
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
    const actual = await command.validate({ options: { webUrl: validWebUrl, entraGroupName: validEntraM365GroupName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and entraGroupId is passed', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, entraGroupId: validEntraGroupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });


  it('should prompt before removing user using id from web when confirmation argument not passed ', async () => {
    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/subsite',
        id: 10
      }
    });

    assert(promptIssued);
  });

  it('should prompt before removing user using login name from web when confirmation argument not passed ', async () => {
    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/subsite',
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    });

    assert(promptIssued);
  });

  it('removes user by id successfully without prompting with confirmation argument', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removebyid(10)`) {
        return true;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        id: 10,
        force: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `${validWebUrl}/_api/web/siteusers/removebyid(10)` &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes user by login name successfully without prompting with confirmation argument', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removeByLoginName('i%3A0%23.f%7Cmembership%7Cparker%40tenant.onmicrosoft.com')`) {
        return true;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        loginName: "i:0#.f|membership|parker@tenant.onmicrosoft.com",
        force: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `${validWebUrl}/_api/web/siteusers/removeByLoginName('i%3A0%23.f%7Cmembership%7Cparker%40tenant.onmicrosoft.com')` &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes user by id successfully from web when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removebyid(10)`) {
        return true;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        id: 10
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `${validWebUrl}/_api/web/siteusers/removebyid(10)` &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes user by login name successfully from web when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removeByLoginName('${formatting.encodeQueryParameter(validLoginName)}')`) {
        return true;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        loginName: validLoginName
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `${validWebUrl}/_api/web/siteusers/removeByLoginName('${formatting.encodeQueryParameter(validLoginName)}')` &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes user from web successfully without prompting with confirmation argument (verbose)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removebyid(10)`) {
        return true;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: validWebUrl,
        id: 10,
        force: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `${validWebUrl}/_api/web/siteusers/removebyid(10)` &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes user from web successfully without prompting with confirmation argument (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removebyid(10)`) {
        return true;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        id: 10,
        force: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `${validWebUrl}/_api/web/siteusers/removebyid(10)` &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes user by email successfully without prompting with confirmation argument', async () => {
    let removeRequestIssued = false;
    sinon.stub(spo, 'getUserByEmail').resolves(userResponse);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removebyid(10)`) {
        removeRequestIssued = true;
        return Promise.resolve();
      }
      throw `Invalid request`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        email: validEmail,
        force: true
      }
    });
    assert(removeRequestIssued);
  });

  it('removes user by username successfully without prompting with confirmation argument', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers?$filter=UserPrincipalName eq ('${formatting.encodeQueryParameter(validUserName)}')`) {
        return {
          "value": [userResponse]
        };
      }
      throw `Invalid request`;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removebyid(10)`) {
        return true;
      }
      throw `Invalid request`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        userName: validUserName,
        force: true
      }
    });
    assert(true);
  });

  it('removes user by entraGroupId successfully without prompting with confirmation argument', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validEntraGroupId}`) {
        return groupM365Response.value[0];
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removeByLoginName('c:0o.c|federateddirectoryclaimprovider|${validEntraGroupId}')`) {
        return true;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        entraGroupId: validEntraGroupId,
        force: true
      }
    });
    assert(true);
  });

  it('removes m365 group by entraGroupName successfully without prompting with confirmation argument', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${validEntraM365GroupName}'`) {
        return groupM365Response;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removeByLoginName('c:0o.c|federateddirectoryclaimprovider|${validEntraGroupId}')`) {
        return true;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        entraGroupName: validEntraM365GroupName,
        force: true
      }
    });
    assert(true);
  });

  it('removes security group by entraGroupName successfully without prompting with confirmation argument', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${validEntraSecurityGroupName}'`) {
        return groupSecurityResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removeByLoginName('c:0t.c|tenant|${validEntraGroupId}')`) {
        return true;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        entraGroupName: validEntraSecurityGroupName,
        force: true
      }
    } as any);
  });

  it('handles error when removing user using from web', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers/removebyid(10)`) {
        throw 'An error has occurred';
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        id: 10,
        force: true
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('handles generic error when user not found when username is passed without prompting with confirmation argument', async () => {
    const err = `User not found: ${validUserName}`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${validWebUrl}/_api/web/siteusers?$filter=UserPrincipalName eq ('${formatting.encodeQueryParameter(validUserName)}')`) {
        return { "value": [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        userName: validUserName,
        force: true
      }
    }), new CommandError(err));
  });
});

