import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './user-ensure.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

describe(commands.USER_ENSURE, () => {
  const validUserName = 'john@contoso.com';
  const validEntraId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
  const validWebUrl = 'https://contoso.sharepoint.com';
  const validEntraGroupId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
  const validEntraGroupName = 'Finance';
  const validEntraSecurityGroupName = 'EntraGroupTest';
  const validLoginName = `i:0#.f|membership|${validUserName}`;
  const ensuredUserResponse = {
    Id: 35,
    IsHiddenInUI: false,
    LoginName: `i:0#.f|membership|${validUserName}`,
    Title: 'John Doe',
    PrincipalType: 1,
    Email: 'john@contoso.com',
    Expiration: '',
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: false,
    UserId: {
      NameId: '1003200274f51d2d',
      NameIdIssuer: 'urn:federation:microsoftonline'
    },
    UserPrincipalName: validUserName
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

  const ensuredGroupResponse = {
    Id: 35,
    IsHiddenInUI: false,
    LoginName: `c:0o.c|federateddirectoryclaimprovider|${validEntraGroupId}`,
    Title: validEntraGroupName,
    PrincipalType: 4,
    Email: 'finance@contoso.com',
    Expiration: '',
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: false,
    UserId: null,
    UserPrincipalName: null
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

  const ensuredSecurityGroupResponse = {
    logonName: 'c:0t.c|tenant|2056d2f6-3257-4253-8cfc-b73393e414e5',
    Id: 35,
    IsHiddenInUI: false,
    LoginName: `c:0t.c|tenant||${validEntraGroupId}`,
    Title: validEntraGroupName,
    PrincipalType: 4,
    Email: null,
    Expiration: '',
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: false,
    UserId: null,
    UserPrincipalName: null
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      entraUser.getUpnByUserId,
      entraGroup.getGroupById,
      entraGroup.getGroupByDisplayName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_ENSURE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('ensures user in a specific web by userPrincipalName', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureuser`) {
        return ensuredUserResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: validWebUrl, userName: validUserName } });
    assert(loggerLogSpy.calledWith(ensuredUserResponse));
  });

  it('ensures user in a specific web by entraId', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').callsFake(async () => {
      return validUserName;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureuser`) {
        return ensuredUserResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: validWebUrl, entraId: validEntraId } });
    assert(loggerLogSpy.calledWith(ensuredUserResponse));
  });

  it('ensures user in a specific web by loginName', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureuser`) {
        return ensuredUserResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: validWebUrl, loginName: validLoginName } });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, { logonName: 'i:0#.f|membership|john@contoso.com' });
  });

  it('ensures user in a specific web by entraGroupId', async () => {
    sinon.stub(entraGroup, 'getGroupById').resolves(groupM365Response.value[0]);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureuser`) {
        return ensuredGroupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: validWebUrl, entraGroupId: validEntraGroupId } });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, { logonName: 'c:0o.c|federateddirectoryclaimprovider|2056d2f6-3257-4253-8cfc-b73393e414e5' });
  });

  it('ensures security group in a specific web by entraGroupName', async () => {
    sinon.stub(entraGroup, 'getGroupByDisplayName').resolves(groupSecurityResponse.value[0]);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureuser`) {
        return ensuredSecurityGroupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: validWebUrl, entraGroupName: validEntraSecurityGroupName } });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, { logonName: 'c:0t.c|tenant|2056d2f6-3257-4253-8cfc-b73393e414e5' });
  });

  it('ensures group in a specific web by entraGroupName', async () => {
    sinon.stub(entraGroup, 'getGroupByDisplayName').resolves(groupM365Response.value[0]);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureuser`) {
        return ensuredGroupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: validWebUrl, entraGroupName: validEntraGroupName } });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, { logonName: 'c:0o.c|federateddirectoryclaimprovider|2056d2f6-3257-4253-8cfc-b73393e414e5' });
  });

  it('throws error message when no user was found with a specific id', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').callsFake(async (id) => {
      throw {
        "error": {
          "error": {
            "code": "Request_ResourceNotFound",
            "message": `Resource '${id}' does not exist or one of its queried reference-property objects are not present.`,
            "innerError": {
              "date": "2023-02-17T22:44:21",
              "request-id": "25519ac1-8f24-46a7-90b0-19baace49a7a",
              "client-request-id": "25519ac1-8f24-46a7-90b0-19baace49a7a"
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, webUrl: validWebUrl, entraId: validEntraId } }), new CommandError(`Resource '${validEntraId}' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('throws error message when no user was found with a specific user name', async () => {
    const error = {
      'error': {
        'odata.error': {
          'code': '-2146232832, Microsoft.SharePoint.SPException',
          'message': {
            'lang': 'en-US',
            'value': `The specified user ${validUserName} could not be found.`
          }
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureuser`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, webUrl: validWebUrl, userName: validUserName } }), new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if webUrl is not a valid url', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', entraId: validEntraId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if entraId is not a valid id', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, entraId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid user principal name', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if entraGroupId is not a valid id', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, entraGroupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url is valid and entraId is a valid id', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, entraId: validEntraId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and userName is a valid user principal name', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, userName: validUserName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and loginName is passed', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, loginName: validLoginName } }, commandInfo);
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
