import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './enterpriseapp-list.js';

describe(commands.ENTERPRISEAPP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const displayName = "My custom enterprise application";
  const tag = "WindowsAzureActiveDirectoryIntegratedApp";
  const servicePrincipalResponse: any = {
    value: [
      {
        id: "226859cc-86f0-40d3-b308-f43b3a729b6e",
        deletedDateTime: null,
        accountEnabled: true,
        alternativeNames: [],
        appDisplayName: "My custom enterprise application",
        appDescription: null,
        appId: "a62ef842-f9ef-49cf-9119-31b85ea58445",
        applicationTemplateId: null,
        appOwnerOrganizationId: "fd71909b-55e5-44d2-9f78-dc432421d527",
        appRoleAssignmentRequired: false,
        createdDateTime: "2022-11-28T20:32:11Z",
        description: null,
        disabledByMicrosoftStatus: null,
        displayName: "My custom enterprise application",
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
            adminConsentDescription: "Allow the application to access My custom enterprise application on behalf of the signed -in user.",
            adminConsentDisplayName: "Access My custom enterprise application",
            id: "907a8cea-411a-461a-bb30-261e52febcca",
            isEnabled: true,
            type: "User",
            userConsentDescription: "Allow the application to access My custom enterprise application on your behalf.",
            userConsentDisplayName: "Access My custom enterprise application",
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
    auth.connection.active = true;
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENTERPRISEAPP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['appId', 'displayName', 'tag']);
  });

  it('list all enterprise applications', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals`) {
        return servicePrincipalResponse;
      }

      return 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(servicePrincipalResponse.value));
  });

  it('list all enterprise applications with the given display name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=(displayName eq '${displayName}')`) {
        return servicePrincipalResponse;
      }

      return 'Invalid request';
    });

    await command.action(logger, { options: { displayName: displayName } });
    assert(loggerLogSpy.calledWith(servicePrincipalResponse.value));
  });

  it('list all enterprise applications with the given tag', async () => {
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
          value: "The enterprise applications could not be retrieved"
        }
      }
    };

    sinon.stub(request, 'get').callsFake(async () => {
      throw error;
    });

    await assert.rejects(command.action(logger, { options: {} } as any), error['odata.error'].message.value);
  });
});
