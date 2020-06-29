import commands from '../../commands';
import Command, { CommandValidate, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./user-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.USER_GET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.USER_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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

    cmdInstance.action({
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 1
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

    cmdInstance.action({
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        email: "john.deo@mytenant.onmicrosoft.com"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

    cmdInstance.action({
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', id: 1 } });
    assert.notEqual(actual, true);
  });

  it('fails validation if id or email or loginName options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if id, email and loginName options are passed (multiple options)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, email: "jonh.deo@mytenant.com", loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } });
    assert.notEqual(actual, true);
  });

  it('fails validation if id and email both are passed (multiple options)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, email: "jonh.deo@mytenant.com" } });
    assert.notEqual(actual, true);
  });

  it('fails validation if id and loginName options are passed (multiple options)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } });
    assert.notEqual(actual, true);
  });

  it('fails validation if email and loginName options are passed (multiple options)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', email: "jonh.deo@mytenant.com", loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } });
    assert.notEqual(actual, true);
  });

  it('fails validation if specified id is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'a' } });
    assert.notEqual(actual, true);
  });

  it('passes validation url is valid and id is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1 } });
    assert.equal(actual, true);
  });

  it('passes validation if the url is valid and email is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', email: "jonh.deo@mytenant.com" } });
    assert.equal(actual, true);
  });

  it('passes validation if the url is valid and loginName is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } });
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.USER_GET));
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