import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
const command: Command = require('./id-get');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import request from '../../../../request';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';

describe(commands.TENANT_ID_GET, () => {
  let log: any[];
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let cmdInstance: any;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: any) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_ID_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets logged in Microsoft 365 tenant ID if no domain name is passed', (done) => {
    sinon.stub(Utils, 'getUserNameFromAccessToken').callsFake(() => {
      return 'admin@contoso.onmicrosoft.com';
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://login.windows.net/contoso.onmicrosoft.com/.well-known/openid-configuration`) {
        return Promise.resolve(
          {
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
          }
        );
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({ options: {} }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('31537af4-6d77-4bb9-a681-d2394888ea26'));
        Utils.restore(Utils.getUserNameFromAccessToken);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets Microsoft 365 tenant ID with correct domain name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://login.windows.net/contoso.com/.well-known/openid-configuration`) {
        return Promise.resolve(
          {
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
          }
        );
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({ options: { debug: false, domainName: 'contoso.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('6babcaad-604b-40ac-a9d7-9fd97c0b779f'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns errors when trying to retrieve ID for a non-existant tenant', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://login.windows.net/xyz.com/.well-known/openid-configuration`) {
        return Promise.resolve(
          {
            "error": "invalid_tenant",
            "error_description": "AADSTS90002: Tenant 'xyz.com' not found. This may happen if there are no active subscriptions for the tenant. Check with your subscription administrator.\r\nTrace ID: 8c0e5644-738f-460f-900c-edb4c918b100\r\nCorrelation ID: 69a7237f-1f84-4b88-aae7-8f7fd46d685a\r\nTimestamp: 2019-06-15 15:41:39Z",
            "error_codes": [
              90002
            ],
            "timestamp": "2019-06-15 15:41:39Z",
            "trace_id": "8c0e5644-738f-460f-900c-edb4c918b100",
            "correlation_id": "69a7237f-1f84-4b88-aae7-8f7fd46d685a"
          }
        );
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({ options: { debug: false, domainName: 'xyz.com' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("AADSTS90002: Tenant 'xyz.com' not found. This may happen if there are no active subscriptions for the tenant. Check with your subscription administrator.\r\nTrace ID: 8c0e5644-738f-460f-900c-edb4c918b100\r\nCorrelation ID: 69a7237f-1f84-4b88-aae7-8f7fd46d685a\r\nTimestamp: 2019-06-15 15:41:39Z")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    cmdInstance.action({ options: { debug: false, domainName: 'xyz.com' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});