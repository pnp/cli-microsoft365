import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./user-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.USER_GET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
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
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
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
    assert.strictEqual(command.name.startsWith(commands.USER_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves user using id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/68be84bf-a585-4776-80b3-30aa5207aa21`) {
        return Promise.resolve({"id":"68be84bf-a585-4776-80b3-30aa5207aa21","businessPhones":["+1 425 555 0100"],"displayName":"Aarif Sherzai","givenName":"Aarif","jobTitle":"Administrative","mail":null,"mobilePhone":"+1 425 555 0100","officeLocation":null,"preferredLanguage":null,"surname":"Sherzai","userPrincipalName":"AarifS@contoso.onmicrosoft.com"});
      }

      return Promise.reject('Invalid request');
    });

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

    cmdInstance.action({ options: { debug: false, id: '68be84bf-a585-4776-80b3-30aa5207aa22' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Resource '68be84bf-a585-4776-80b3-30aa5207aa22' does not exist or one of its queried reference-property objects are not present.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if neither the id nor the userName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both the id and the userName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22', userName: 'AarifS@contoso.onmicrosoft.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if the userName is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { userName: 'AarifS@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
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