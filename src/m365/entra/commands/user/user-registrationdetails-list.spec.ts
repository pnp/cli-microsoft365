import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import command from './user-registrationdetails-list.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.USER_REGISTRATIONDETAILS_LIST, () => {
  const registrationDetails = [
    {
      "id": "61b0c52f-a902-4769-9a09-c6628335b00a",
      "userPrincipalName": "AdeleV@contoso.onmicrosoft.com",
      "userDisplayName": "Adele Vance",
      "userType": "member",
      "isAdmin": false,
      "isSsprRegistered": false,
      "isSsprEnabled": false,
      "isSsprCapable": false,
      "isMfaRegistered": false,
      "isMfaCapable": false,
      "isPasswordlessCapable": false,
      "methodsRegistered": [],
      "isSystemPreferredAuthenticationMethodEnabled": false,
      "systemPreferredAuthenticationMethods": [],
      "userPreferredMethodForSecondaryAuthentication": "none",
      "lastUpdatedDateTime": "2024-01-11T11:38:04.5006379Z"
    },
    {
      "id": "f9e0ee63-73dc-48a9-aa97-e5159ec11705",
      "userPrincipalName": "JohannaL@contoso.onmicrosoft.com",
      "userDisplayName": "Johanna Lorenz",
      "userType": "member",
      "isAdmin": false,
      "isSsprRegistered": false,
      "isSsprEnabled": false,
      "isSsprCapable": false,
      "isMfaRegistered": true,
      "isMfaCapable": true,
      "isPasswordlessCapable": false,
      "methodsRegistered": [
        "microsoftAuthenticatorPush",
        "softwareOneTimePasscode"
      ],
      "isSystemPreferredAuthenticationMethodEnabled": false,
      "systemPreferredAuthenticationMethods": [],
      "userPreferredMethodForSecondaryAuthentication": "push",
      "lastUpdatedDateTime": "2024-01-11T11:38:04.5053823Z"
    },
    {
      "id": "abcd1234-e024-4bc6-8e98-123458962525",
      "userPrincipalName": "JohnDoe@contoso.onmicrosoft.com",
      "userDisplayName": "John Doe",
      "userType": "member",
      "isAdmin": true,
      "isSsprRegistered": true,
      "isSsprEnabled": true,
      "isSsprCapable": true,
      "isMfaRegistered": true,
      "isMfaCapable": true,
      "isPasswordlessCapable": false,
      "methodsRegistered": [
        "email",
        "mobilePhone",
        "microsoftAuthenticatorPush",
        "softwareOneTimePasscode"
      ],
      "isSystemPreferredAuthenticationMethodEnabled": false,
      "systemPreferredAuthenticationMethods": [],
      "userPreferredMethodForSecondaryAuthentication": "push",
      "lastUpdatedDateTime": "2024-01-11T11:38:04.5040399Z"
    }
  ];

  let log: string[];
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
      cli.getSettingWithDefaultValue,
      entraUser.getUpnByUserId
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_REGISTRATIONDETAILS_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['userPrincipalName', 'methodsRegistered', 'lastUpdatedDateTime']);
  });

  it('fails validation if userType contains invalid value', async () => {
    const actual = await command.validate({ options: { userType: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userPreferredMethodForSecondaryAuthentication contains invalid value', async () => {
    const actual = await command.validate({ options: { userPreferredMethodForSecondaryAuthentication: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if systemPreferredAuthenticationMethods contains invalid value', async () => {
    const actual = await command.validate({ options: { systemPreferredAuthenticationMethods: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if methodsRegistered contains invalid value', async () => {
    const actual = await command.validate({ options: { methodsRegistered: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userIds contains invalid GUID', async () => {
    const userIds = ['7167b488-1ffb-43f1-9547-35969469bada', 'foo'];
    const actual = await command.validate({ options: { userIds: userIds.join(',') } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userPrincipalNames contains invalid user principal name', async () => {
    const userPrincipalNames = ['john.doe@contoso.com', 'foo'];
    const actual = await command.validate({ options: { userPrincipalNames: userPrincipalNames.join(',') } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all optional parameters are valid', async () => {
    const userIds = ['7167b488-1ffb-43f1-9547-35969469bada', '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302'];
    const userPrincipalNames = ['john.doe@contoso.com', 'adele.vance@contoso.com'];
    const actual = await command.validate({
      options:
      {
        userType: 'guest',
        userPreferredMethodForSecondaryAuthentication: 'push',
        systemPreferredAuthenticationMethods: 'push',
        methodsRegistered: 'microsoftAuthenticatorPush',
        userIds: userIds.join(','),
        userPrincipalNames: userPrincipalNames.join(',')
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should get a list of user registration details', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails`) {
        return {
          value: registrationDetails
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });

    assert(loggerLogSpy.calledWith(registrationDetails));
  });

  it('should get a list of user registration details with selected properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?$select=id,userPrincipalName,methodsRegistered`) {
        return {
          value: registrationDetails
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { properties: 'id,userPrincipalName,methodsRegistered' } });

    assert(loggerLogSpy.calledWith(registrationDetails));
  });

  it('should get a filtered list of user registration details when bool options are set to true', async () => {
    const filter = 'isAdmin eq true and isMfaCapable eq true and isMfaRegistered eq true and isPasswordlessCapable eq true and isSelfServicePasswordResetCapable eq true ' +
      'and isSelfServicePasswordResetEnabled eq true and isSelfServicePasswordResetRegistered eq true and isSystemPreferredAuthenticationMethodEnabled eq true ' +
      `and (methodsRegistered/any(m:m eq 'fido2') or methodsRegistered/any(m:m eq 'microsoftAuthenticatorPush')) ` +
      `and (systemPreferredAuthenticationMethods/any(m:m eq 'oath') or systemPreferredAuthenticationMethods/any(m:m eq 'voiceMobile') or systemPreferredAuthenticationMethods/any(m:m eq 'push')) ` +
      `and (userPrincipalName eq '${formatting.encodeQueryParameter('joe.guest_external#EXT#@contoso.com')}' or userPrincipalName eq '${formatting.encodeQueryParameter('PradeepG@contoso.com')}' or userPrincipalName eq '${formatting.encodeQueryParameter('john.doe@contoso.com')}' or userPrincipalName eq '${formatting.encodeQueryParameter('adele.vance@contoso.com')}') ` +
      `and (userPreferredMethodForSecondaryAuthentication eq 'oath' or userPreferredMethodForSecondaryAuthentication eq 'voiceMobile' or userPreferredMethodForSecondaryAuthentication eq 'push') ` +
      `and userType eq 'member'`;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?$filter=${filter}`) {
        return {
          value: registrationDetails
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(entraUser, 'getUpnByUserId').callsFake(async (opts) => {
      if (opts === '7167b488-1ffb-43f1-9547-35969469bada') {
        return 'joe.guest_external#EXT#@contoso.com';
      }
      else if (opts === '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302') {
        return 'PradeepG@contoso.com';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        isAdmin: true,
        userType: 'member',
        userPreferredMethodForSecondaryAuthentication: 'oath, voiceMobile,push',
        systemPreferredAuthenticationMethods: 'oath, voiceMobile,push',
        isSelfServicePasswordResetCapable: true,
        isSelfServicePasswordResetEnabled: true,
        isSelfServicePasswordResetRegistered: true,
        isMfaCapable: true,
        isMfaRegistered: true,
        isPasswordlessCapable: true,
        isSystemPreferredAuthenticationMethodEnabled: true,
        methodsRegistered: 'fido2, microsoftAuthenticatorPush',
        userIds: '7167b488-1ffb-43f1-9547-35969469bada, 6dcd4ce0-4f89-11d3-9a0c-0305e82c3302',
        userPrincipalNames: 'john.doe@contoso.com, adele.vance@contoso.com'
      }
    });

    assert(loggerLogSpy.calledWith(registrationDetails));
  });

  it('should get a filtered list of user registration details when bool options are set to false', async () => {
    const filter = 'isAdmin eq false and isMfaCapable eq false and isMfaRegistered eq false and isPasswordlessCapable eq false and isSelfServicePasswordResetCapable eq false and ' +
      'isSelfServicePasswordResetEnabled eq false and isSelfServicePasswordResetRegistered eq false and isSystemPreferredAuthenticationMethodEnabled eq false';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?$filter=${filter}`) {
        return {
          value: registrationDetails
        };
      }

      throw opts.url;
    });

    await command.action(logger, {
      options: {
        isAdmin: false,
        isSelfServicePasswordResetCapable: false,
        isSelfServicePasswordResetEnabled: false,
        isSelfServicePasswordResetRegistered: false,
        isMfaCapable: false,
        isMfaRegistered: false,
        isPasswordlessCapable: false,
        isSystemPreferredAuthenticationMethodEnabled: false
      }
    });

    assert(loggerLogSpy.calledWith(registrationDetails));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: {} }),
      new CommandError('Invalid request'));
  });
});