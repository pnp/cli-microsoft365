import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './id-get.js';

describe(commands.ID_GET, () => {
  let log: any[];
  let loggerLogSpy: sinon.SinonSpy;
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    if (!auth.connection.accessTokens[auth.defaultResource]) {
      auth.connection.accessTokens[auth.defaultResource] = {
        expiresOn: '123',
        accessToken: 'abc'
      };
    }
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
    assert.strictEqual(command.name, commands.ID_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets logged in Microsoft 365 tenant ID if no domain name is passed', async () => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').callsFake(() => {
      return 'admin@contoso.onmicrosoft.com';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://login.windows.net/contoso.onmicrosoft.com/.well-known/openid-configuration`) {
        return {
          "token_endpoint": "https://login.windows.net/31537af4-6d77-4bb9-a681-d2394888ea26/oauth2/token",
          "token_endpoint_auth_methods_supported": [
            "client_secret_post",
            "private_key_jwt",
            "client_secret_basic"
          ],
          "jwks_uri": "https://login.windows.net/common/discovery/keys",
          "response_modes_supported": [
            "query",
            "fragment",
            "form_post"
          ],
          "subject_types_supported": [
            "pairwise"
          ],
          "id_token_signing_alg_values_supported": [
            "RS256"
          ],
          "response_types_supported": [
            "code",
            "id_token",
            "code id_token",
            "token id_token",
            "token"
          ],
          "scopes_supported": [
            "openid"
          ],
          "issuer": "https://sts.windows.net/31537af4-6d77-4bb9-a681-d2394888ea26/",
          "microsoft_multi_refresh_token": true,
          "authorization_endpoint": "https://login.windows.net/31537af4-6d77-4bb9-a681-d2394888ea26/oauth2/authorize",
          "http_logout_supported": true,
          "frontchannel_logout_supported": true,
          "end_session_endpoint": "https://login.windows.net/31537af4-6d77-4bb9-a681-d2394888ea26/oauth2/logout",
          "claims_supported": [
            "sub",
            "iss",
            "cloud_instance_name",
            "cloud_instance_host_name",
            "cloud_graph_host_name",
            "msgraph_host",
            "aud",
            "exp",
            "iat",
            "auth_time",
            "acr",
            "amr",
            "nonce",
            "email",
            "given_name",
            "family_name",
            "nickname"
          ],
          "check_session_iframe": "https://login.windows.net/31537af4-6d77-4bb9-a681-d2394888ea26/oauth2/checksession",
          "userinfo_endpoint": "https://login.windows.net/31537af4-6d77-4bb9-a681-d2394888ea26/openid/userinfo",
          "tenant_region_scope": "AS",
          "cloud_instance_name": "microsoftonline.com",
          "cloud_graph_host_name": "graph.windows.net",
          "msgraph_host": "graph.microsoft.com",
          "rbac_url": "https://pas.windows.net"
        };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith('31537af4-6d77-4bb9-a681-d2394888ea26'));
    sinonUtil.restore(accessToken.getUserNameFromAccessToken);
  });

  it('gets Microsoft 365 tenant ID with correct domain name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://login.windows.net/contoso.com/.well-known/openid-configuration`) {
        return {
          "authorization_endpoint": "https://login.windows.net/6babcaad-604b-40ac-a9d7-9fd97c0b779f/oauth2/authorize",
          "token_endpoint": "https://login.windows.net/6babcaad-604b-40ac-a9d7-9fd97c0b779f/oauth2/token",
          "token_endpoint_auth_methods_supported": [
            "client_secret_post",
            "private_key_jwt",
            "client_secret_basic"
          ],
          "jwks_uri": "https://login.windows.net/common/discovery/keys",
          "response_modes_supported": [
            "query",
            "fragment",
            "form_post"
          ],
          "subject_types_supported": [
            "pairwise"
          ],
          "id_token_signing_alg_values_supported": [
            "RS256"
          ],
          "http_logout_supported": true,
          "frontchannel_logout_supported": true,
          "end_session_endpoint": "https://login.windows.net/6babcaad-604b-40ac-a9d7-9fd97c0b779f/oauth2/logout",
          "response_types_supported": [
            "code",
            "id_token",
            "code id_token",
            "token id_token",
            "token"
          ],
          "scopes_supported": [
            "openid"
          ],
          "issuer": "https://sts.windows.net/6babcaad-604b-40ac-a9d7-9fd97c0b779f/",
          "claims_supported": [
            "sub",
            "iss",
            "cloud_instance_name",
            "cloud_instance_host_name",
            "cloud_graph_host_name",
            "msgraph_host",
            "aud",
            "exp",
            "iat",
            "auth_time",
            "acr",
            "amr",
            "nonce",
            "email",
            "given_name",
            "family_name",
            "nickname"
          ],
          "microsoft_multi_refresh_token": true,
          "check_session_iframe": "https://login.windows.net/6babcaad-604b-40ac-a9d7-9fd97c0b779f/oauth2/checksession",
          "userinfo_endpoint": "https://login.windows.net/6babcaad-604b-40ac-a9d7-9fd97c0b779f/openid/userinfo",
          "tenant_region_scope": "NA",
          "cloud_instance_name": "microsoftonline.com",
          "cloud_graph_host_name": "graph.windows.net",
          "msgraph_host": "graph.microsoft.com",
          "rbac_url": "https://pas.windows.net"
        };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { domainName: 'contoso.com' } });
    assert(loggerLogSpy.calledWith('6babcaad-604b-40ac-a9d7-9fd97c0b779f'));
  });

  it('returns errors when trying to retrieve ID for a non-existant tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://login.windows.net/xyz.com/.well-known/openid-configuration`) {
        return {
          "error": "invalid_tenant",
          "error_description": "AADSTS90002: Tenant 'xyz.com' not found. This may happen if there are no active subscriptions for the tenant. Check with your subscription administrator.\r\nTrace ID: 8c0e5644-738f-460f-900c-edb4c918b100\r\nCorrelation ID: 69a7237f-1f84-4b88-aae7-8f7fd46d685a\r\nTimestamp: 2019-06-15 15:41:39Z",
          "error_codes": [
            90002
          ],
          "timestamp": "2019-06-15 15:41:39Z",
          "trace_id": "8c0e5644-738f-460f-900c-edb4c918b100",
          "correlation_id": "69a7237f-1f84-4b88-aae7-8f7fd46d685a"
        };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, { options: { domainName: 'xyz.com' } } as any), new CommandError("AADSTS90002: Tenant 'xyz.com' not found. This may happen if there are no active subscriptions for the tenant. Check with your subscription administrator.\r\nTrace ID: 8c0e5644-738f-460f-900c-edb4c918b100\r\nCorrelation ID: 69a7237f-1f84-4b88-aae7-8f7fd46d685a\r\nTimestamp: 2019-06-15 15:41:39Z"));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { domainName: 'xyz.com' } } as any), new CommandError('An error has occurred'));
  });
});
