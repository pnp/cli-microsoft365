import commands from '../commands';
import Command, { CommandValidate } from '../../../Command';
import * as sinon from 'sinon';
const command: Command = require('./tenant-id-get');
import * as assert from 'assert';
import Utils from '../../../Utils';
import request from '../../../request';
import appInsights from '../../../appInsights';

describe(commands.TENANT_ID_GET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: any) => {
        log.push(msg);
      }
    };
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore(vorpal.find);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.TENANT_ID_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.TENANT_ID_GET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if domainName is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if domainName is undefined', () => {
    const actual = (command.validate() as CommandValidate)({ options: {domainName: undefined} });
    assert.notEqual(actual, true);
  });

  it('fails validation if domainName is blank', () => {
    const actual = (command.validate() as CommandValidate)({ options: {domainName: ''} });
    assert.notEqual(actual, true);
  });

  it('passes validation on valid domainName', () => {
    const actual = (command.validate() as CommandValidate)({ options: { domainName: 'contoso.com' } });
    assert.equal(actual, true);
  });

  it('gets Microsoft Azure or Office 365 tenant ID', (done) => {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, domainName: 'contoso.com' } }, () => {
      try {
        assert(1 === 1);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TENANT_ID_GET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});