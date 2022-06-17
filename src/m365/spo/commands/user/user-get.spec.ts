import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./user-get');

describe(commands.USER_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.USER_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves user by id with output option json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/siteusers/GetById') > -1) {
        return Promise.resolve(
          {
            "value": [{
              "Id": 6,
              "IsHiddenInUI": false,
              "LoginName": "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
              "Title": "John Doe",
              "PrincipalType": 1,
              "Email": "john.deo@mytenant.onmicrosoft.com",
              "Expiration": "",
              "IsEmailAuthenticationGuestUser": false,
              "IsShareByEmailGuestUser": false,
              "IsSiteAdmin": false,
              "UserId": { "NameId": "10010001b0c19a2", "NameIdIssuer": "urn:federation:microsoftonline" },
              "UserPrincipalName": "john.deo@mytenant.onmicrosoft.com"
            }]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 1
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          value: [{
            Id: 6,
            IsHiddenInUI: false,
            LoginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
            Title: "John Doe",
            PrincipalType: 1,
            Email: "john.deo@mytenant.onmicrosoft.com",
            Expiration: "",
            IsEmailAuthenticationGuestUser: false,
            IsShareByEmailGuestUser: false,
            IsSiteAdmin: false,
            UserId: { NameId: "10010001b0c19a2", NameIdIssuer: "urn:federation:microsoftonline" },
            UserPrincipalName: "john.deo@mytenant.onmicrosoft.com"
          }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('retrieves user by email with output option json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/siteusers/GetByEmail') > -1) {
        return Promise.resolve(
          {
            "value": [{
              "Id": 6,
              "IsHiddenInUI": false,
              "LoginName": "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
              "Title": "John Doe",
              "PrincipalType": 1,
              "Email": "john.deo@mytenant.onmicrosoft.com",
              "Expiration": "",
              "IsEmailAuthenticationGuestUser": false,
              "IsShareByEmailGuestUser": false,
              "IsSiteAdmin": false,
              "UserId": { "NameId": "10010001b0c19a2", "NameIdIssuer": "urn:federation:microsoftonline" },
              "UserPrincipalName": "john.deo@mytenant.onmicrosoft.com"
            }]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        email: "john.deo@mytenant.onmicrosoft.com"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          value: [{
            Id: 6,
            IsHiddenInUI: false,
            LoginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
            Title: "John Doe",
            PrincipalType: 1,
            Email: "john.deo@mytenant.onmicrosoft.com",
            Expiration: "",
            IsEmailAuthenticationGuestUser: false,
            IsShareByEmailGuestUser: false,
            IsSiteAdmin: false,
            UserId: { NameId: "10010001b0c19a2", NameIdIssuer: "urn:federation:microsoftonline" },
            UserPrincipalName: "john.deo@mytenant.onmicrosoft.com"
          }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves user by loginName with output option json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/siteusers/GetByLoginName') > -1) {
        return Promise.resolve(
          {
            "value": [{
              "Id": 6,
              "IsHiddenInUI": false,
              "LoginName": "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
              "Title": "John Doe",
              "PrincipalType": 1,
              "Email": "john.deo@mytenant.onmicrosoft.com",
              "Expiration": "",
              "IsEmailAuthenticationGuestUser": false,
              "IsShareByEmailGuestUser": false,
              "IsSiteAdmin": false,
              "UserId": { "NameId": "10010001b0c19a2", "NameIdIssuer": "urn:federation:microsoftonline" },
              "UserPrincipalName": "john.deo@mytenant.onmicrosoft.com"
            }]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          value: [{
            Id: 6,
            IsHiddenInUI: false,
            LoginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
            Title: "John Doe",
            PrincipalType: 1,
            Email: "john.deo@mytenant.onmicrosoft.com",
            Expiration: "",
            IsEmailAuthenticationGuestUser: false,
            IsShareByEmailGuestUser: false,
            IsSiteAdmin: false,
            UserId: { NameId: "10010001b0c19a2", NameIdIssuer: "urn:federation:microsoftonline" },
            UserPrincipalName: "john.deo@mytenant.onmicrosoft.com"
          }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('fails validation if id or email or loginName options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id, email and loginName options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, email: "jonh.deo@mytenant.com", loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id and email both are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, email: "jonh.deo@mytenant.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id and loginName options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if email and loginName options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', email: "jonh.deo@mytenant.com", loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified id is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation url is valid and id is passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and email is passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', email: "jonh.deo@mytenant.com" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and loginName is passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
    assert.strictEqual(actual, true);
  });
}); 