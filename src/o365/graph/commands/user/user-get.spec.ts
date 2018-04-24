import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./user-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.USER_GET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.service = new Service();
    telemetry = null;
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.USER_GET), true);
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
        assert.equal(telemetry.name, commands.USER_GET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to Microsoft Graph', (done) => {
    auth.service = new Service();
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to the Microsoft Graph first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves user using id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/68be84bf-a585-4776-80b3-30aa5207aa21`) {
        return Promise.resolve({"id":"68be84bf-a585-4776-80b3-30aa5207aa21","businessPhones":["+1 425 555 0100"],"displayName":"Aarif Sherzai","givenName":"Aarif","jobTitle":"Administrative","mail":null,"mobilePhone":"+1 425 555 0100","officeLocation":null,"preferredLanguage":null,"surname":"Sherzai","userPrincipalName":"AarifS@contoso.onmicrosoft.com"});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '68be84bf-a585-4776-80b3-30aa5207aa21' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({"id":"68be84bf-a585-4776-80b3-30aa5207aa21","businessPhones":["+1 425 555 0100"],"displayName":"Aarif Sherzai","givenName":"Aarif","jobTitle":"Administrative","mail":null,"mobilePhone":"+1 425 555 0100","officeLocation":null,"preferredLanguage":null,"surname":"Sherzai","userPrincipalName":"AarifS@contoso.onmicrosoft.com"}));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves user using id (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/68be84bf-a585-4776-80b3-30aa5207aa21`) {
        return Promise.resolve({"id":"68be84bf-a585-4776-80b3-30aa5207aa21","businessPhones":["+1 425 555 0100"],"displayName":"Aarif Sherzai","givenName":"Aarif","jobTitle":"Administrative","mail":null,"mobilePhone":"+1 425 555 0100","officeLocation":null,"preferredLanguage":null,"surname":"Sherzai","userPrincipalName":"AarifS@contoso.onmicrosoft.com"});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, id: '68be84bf-a585-4776-80b3-30aa5207aa21' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({"id":"68be84bf-a585-4776-80b3-30aa5207aa21","businessPhones":["+1 425 555 0100"],"displayName":"Aarif Sherzai","givenName":"Aarif","jobTitle":"Administrative","mail":null,"mobilePhone":"+1 425 555 0100","officeLocation":null,"preferredLanguage":null,"surname":"Sherzai","userPrincipalName":"AarifS@contoso.onmicrosoft.com"}));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves user using user name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/AarifS%40contoso.onmicrosoft.com`) {
        return Promise.resolve({"id":"68be84bf-a585-4776-80b3-30aa5207aa21","businessPhones":["+1 425 555 0100"],"displayName":"Aarif Sherzai","givenName":"Aarif","jobTitle":"Administrative","mail":null,"mobilePhone":"+1 425 555 0100","officeLocation":null,"preferredLanguage":null,"surname":"Sherzai","userPrincipalName":"AarifS@contoso.onmicrosoft.com"});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, userName: 'AarifS@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({"id":"68be84bf-a585-4776-80b3-30aa5207aa21","businessPhones":["+1 425 555 0100"],"displayName":"Aarif Sherzai","givenName":"Aarif","jobTitle":"Administrative","mail":null,"mobilePhone":"+1 425 555 0100","officeLocation":null,"preferredLanguage":null,"surname":"Sherzai","userPrincipalName":"AarifS@contoso.onmicrosoft.com"}));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves only the specified properties', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/AarifS%40contoso.onmicrosoft.com?$select=id,mail`) {
        return Promise.resolve({"id":"68be84bf-a585-4776-80b3-30aa5207aa21","mail":null});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, userName: 'AarifS@contoso.onmicrosoft.com', properties: 'id,mail' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({"id":"68be84bf-a585-4776-80b3-30aa5207aa21","mail":null}));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles user not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "Request_ResourceNotFound",
          "message": "Resource '68be84bf-a585-4776-80b3-30aa5207aa22' does not exist or one of its queried reference-property objects are not present.",
          "innerError": {
            "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
            "date": "2018-04-24T18:56:48"
          }
        }
      });
    });

    auth.service = new Service('https://graph.windows.net');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '68be84bf-a585-4776-80b3-30aa5207aa22' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError(`Resource '68be84bf-a585-4776-80b3-30aa5207aa22' does not exist or one of its queried reference-property objects are not present.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if neither the id nor the userName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { } });
    assert.notEqual(actual, true);
  });

  it('fails validation if both the id and the userName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22', userName: 'AarifS@contoso.onmicrosoft.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22' } });
    assert.equal(actual, true);
  });

  it('passes validation if the userName is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { userName: 'AarifS@contoso.onmicrosoft.com' } });
    assert.equal(actual, true);
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

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});