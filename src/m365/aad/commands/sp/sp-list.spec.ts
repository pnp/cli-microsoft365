import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./sp-list');

describe(commands.SP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const displayName = "My custom service principal";
  const tag = "WindowsAzureActiveDirectoryIntegratedApp";
  const servicePrincipalResponse: any = {
    value: [
      {
        id: "226859cc-86f0-40d3-b308-f43b3a729b6e",
        deletedDateTime: null,
        accountEnabled: true,
        alternativeNames: [],
        appDisplayName: "My custom service principal",
        appDescription: null,
        appId: "a62ef842-f9ef-49cf-9119-31b85ea58445",
        applicationTemplateId: null,
        appOwnerOrganizationId: "fd71909b-55e5-44d2-9f78-dc432421d527",
        appRoleAssignmentRequired: false,
        createdDateTime: "2022-11-28T20:32:11Z",
        description: null,
        disabledByMicrosoftStatus: null,
        displayName: "My custom service principal",
        homepage: null,
        loginUrl: null,
        logoutUrl: null,
        notes: null,
        notificationEmailAddresses: [],
        preferredSingleSignOnMode: null,
        preferredTokenSigningKeyThumbprint: null,
        replyUrls: [
          "urn:ietf:wg:oauth:2.0:oob",
          "https://localhost",
          "http://localhost",
          "http://localhost:8400"
        ],
        servicePrincipalNames: [
          "https://contoso.onmicrosoft.com/907a8cea-411a-461a-bb30-261e52febcca",
          "907a8cea-411a-461a-bb30-261e52febcca"
        ],
        servicePrincipalType: "Application",
        signInAudience: "AzureADMultipleOrgs",
        tags: [
          "WindowsAzureActiveDirectoryIntegratedApp"
        ],
        tokenEncryptionKeyId: null,
        samlSingleSignOnSettings: null,
        addIns: [],
        appRoles: [],
        info: {
          logoUrl: null,
          marketingUrl: null,
          privacyStatementUrl: null,
          supportUrl: null,
          termsOfServiceUrl: null
        },
        keyCredentials: [],
        oauth2PermissionScopes: [
          {
            adminConsentDescription: "Allow the application to access My custom service principal on behalf of the signed -in user.",
            adminConsentDisplayName: "Access My custom service principal",
            id: "907a8cea-411a-461a-bb30-261e52febcca",
            isEnabled: true,
            type: "User",
            userConsentDescription: "Allow the application to access My custom service principal on your behalf.",
            userConsentDisplayName: "Access My custom service principal",
            value: "user_impersonation"
          }
        ],
        passwordCredentials: [],
        resourceSpecificApplicationPermissions: [],
        verifiedPublisher: {
          displayName: null,
          verifiedPublisherId: null,
          addedDateTime: null
        }
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['appId', 'displayName', 'tag']);
  });

  it('list all service principals', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals`) {
        return servicePrincipalResponse;
      }

      return 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(servicePrincipalResponse.value));
  });

  it('list all service principals with the given display name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=(displayName eq '${displayName}')`) {
        return servicePrincipalResponse;
      }

      return 'Invalid request';
    });

    await command.action(logger, { options: { displayName: displayName } });
    assert(loggerLogSpy.calledWith(servicePrincipalResponse.value));
  });

  it('list all service principals with the given tag', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=(tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp'))`) {
        return servicePrincipalResponse;
      }

      return 'Invalid request';
    });

    await command.action(logger, { options: { tag: tag } });
    assert(loggerLogSpy.calledWith(servicePrincipalResponse.value));
  });

  it('correctly handles API OData error', async () => {
    const error = {
      'odata.error': {
        message: {
          value: "The service principals could not be retrieved"
        }
      }
    };

    sinon.stub(request, 'get').callsFake(async () => {
      throw error;
    });

    await assert.rejects(command.action(logger, { options: {} } as any), error['odata.error'].message.value);
  });
});
