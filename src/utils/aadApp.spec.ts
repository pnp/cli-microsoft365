import * as assert from 'assert';
import * as sinon from 'sinon';
import request from "../request.js";
import { formatting } from './formatting.js';
import { sinonUtil } from "./sinonUtil.js";
import { aadApp } from './aadApp.js';
import { Logger } from '../cli/Logger.js';

const validAppId = '00000000-0000-0000-0000-000000000000';
const appResponse = {
  id: "d80f247f-487f-409a-8d5f-a44c89179bd2",
  deletedDateTime: null,
  appId: "4a0c5dd7-73cd-47fc-9f38-80937e379445",
  applicationTemplateId: null,
  disabledByMicrosoftStatus: null,
  createdDateTime: "2022-10-31T11:57:57Z",
  displayName: "Custom PnP CLI for Microsoft 365",
  description: null,
  groupMembershipClaims: null,
  identifierUris: [],
  isDeviceOnlyAuthSupported: null,
  isFallbackPublicClient: true,
  notes: null,
  publisherDomain: "contoso.onmicrosoft.com",
  serviceManagementReference: null,
  signInAudience: "AzureADMyOrg",
  tags: [],
  tokenEncryptionKeyId: null,
  samlMetadataUrl: null,
  defaultRedirectUri: null,
  certification: null,
  optionalClaims: null,
  servicePrincipalLockConfiguration: null,
  requestSignatureVerification: null,
  addIns: [],
  api: {
    acceptMappedClaims: null,
    knownClientApplications: [],
    requestedAccessTokenVersion: null,
    oauth2PermissionScopes: [],
    preAuthorizedApplications: []
  },
  appRoles: [],
  info: {
    logoUrl: null,
    marketingUrl: null,
    privacyStatementUrl: null,
    supportUrl: null,
    termsOfServiceUrl: null
  },
  keyCredentials: [],
  parentalControlSettings: {
    countriesBlockedForMinors: [],
    legalAgeGroupRule: "Allow"
  },
  passwordCredentials: [
    {
      customKeyIdentifier: null,
      displayName: "123456",
      endDateTime: "2025-03-29T09:42:45.761Z",
      hint: "OHi",
      keyId: "4e2a37fc-a05a-461b-99e2-f46ab379b77e",
      secretText: null,
      startDateTime: "2023-03-30T08:42:45.761Z"
    },
    {
      customKeyIdentifier: null,
      displayName: "7891011",
      endDateTime: "2023-09-25T11:27:11.663Z",
      hint: "1R9",
      keyId: "ec49d61b-ce18-4667-a156-52565215f99e",
      secretText: null,
      startDateTime: "2023-03-29T11:27:11.663Z"
    },
    {
      customKeyIdentifier: null,
      displayName: "121314",
      endDateTime: "2023-05-01T12:55:26.564Z",
      hint: "v~L",
      keyId: "2a909c8b-59ba-4025-9a9e-80b3bb930954",
      secretText: null,
      startDateTime: "2022-10-31T13:55:26.564Z"
    }
  ],
  publicClient: {
    redirectUris: [
      "https://login.microsoftonline.com/common/oauth2/nativeclient"
    ]
  },
  requiredResourceAccess: [
    {
      resourceAppId: "00000003-0000-0000-c000-000000000000",
      resourceAccess: [
        {
          id: "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
          type: "Scope"
        },
        {
          id: "332a536c-c7ef-4017-ab91-336970924f0d",
          type: "Role"
        },
        {
          id: "19da66cb-0fb0-4390-b071-ebc76a349482",
          type: "Role"
        },
        {
          id: "45bbb07e-7321-4fd7-a8f6-3ff27e6a81c8",
          type: "Role"
        },
        {
          id: "a2611786-80b3-417e-adaa-707d4261a5f0",
          type: "Role"
        }
      ]
    }
  ],
  verifiedPublisher: {
    displayName: null,
    verifiedPublisherId: null,
    addedDateTime: null
  },
  web: {
    homePageUrl: null,
    logoutUrl: null,
    redirectUris: [],
    implicitGrantSettings: {
      enableAccessTokenIssuance: false,
      enableIdTokenIssuance: false
    },
    redirectUriSettings: []
  },
  spa: {
    redirectUris: []
  }
};

describe('utils/aadApp', () => {
  let logger: Logger;
  let log: string[];

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
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

  it('correctly get a single app by id.', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '${formatting.encodeQueryParameter(validAppId)}'`) {
        return { value: [appResponse] };
      }

      return 'Invalid Request';
    });

    const actual = await aadApp.getAppById(validAppId, logger, true);
    assert.strictEqual(actual, appResponse);
  });

  it('throws error message when no app was found with id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '${formatting.encodeQueryParameter(validAppId)}'`) {
        return { value: [] };
      }

      return 'Invalid Request';
    });

    await assert.rejects(aadApp.getAppById(validAppId), `An error occured`);
  });
}); 