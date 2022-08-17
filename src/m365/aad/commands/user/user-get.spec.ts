import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken, sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./user-get');

describe(commands.USER_GET, () => {
  const userId = "68be84bf-a585-4776-80b3-30aa5207aa21";
  const userName = "AarifS@contoso.onmicrosoft.com";
  const resultValue = { "id": "68be84bf-a585-4776-80b3-30aa5207aa21", "businessPhones": ["+1 425 555 0100"], "displayName": "Aarif Sherzai", "givenName": "Aarif", "jobTitle": "Administrative", "mail": null, "mobilePhone": "+1 425 555 0100", "officeLocation": null, "preferredLanguage": null, "surname": "Sherzai", "userPrincipalName": "AarifS@contoso.onmicrosoft.com" };
    
  let log: string[];
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
    (command as any).items = [];
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

  it('retrieves user using id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: userId } }, () => {
      try {
        assert(loggerLogSpy.calledWith(resultValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves user using @userid token', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(accessToken, 'getUserIdFromAccessToken').callsFake(() => { return userId; });    
    auth.service.connected = true;
    if (!auth.service.accessTokens[auth.defaultResource]) {
      auth.service.accessTokens[auth.defaultResource] = {
        expiresOn: '123',
        accessToken: 'abc'
      };
    }

    command.action(logger, { options: { debug: false, id: '@meid' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(resultValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves user using id (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: userId } }, () => {
      try {
        assert(loggerLogSpy.calledWith(resultValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves user using user name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${encodeURIComponent(userName)}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, userName: userName } }, () => {
      try {
        assert(loggerLogSpy.calledWith(resultValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves user using @meusername token', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${encodeURIComponent(userName)}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(accessToken, 'getUserNameFromAccessToken').callsFake(() => { return userName; });    
    auth.service.connected = true;
    if (!auth.service.accessTokens[auth.defaultResource]) {
      auth.service.accessTokens[auth.defaultResource] = {
        expiresOn: '123',
        accessToken: 'abc'
      };
    }

    command.action(logger, { options: { debug: false, userName: '@meusername' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(resultValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves user using email', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=mail eq '${encodeURIComponent(userName)}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, email: userName } }, () => {
      try {
        assert(loggerLogSpy.calledWith(resultValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves only the specified properties', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${encodeURIComponent(userName)}'&$select=id,mail`) {
        return Promise.resolve({ value: [{ "id": "userId", "mail": null }] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, userName: userName, properties: 'id,mail' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({ "id": "userId", "mail": null }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles user not found', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
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

    command.action(logger, { options: { debug: false, id: '68be84bf-a585-4776-80b3-30aa5207aa22' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Resource '68be84bf-a585-4776-80b3-30aa5207aa22' does not exist or one of its queried reference-property objects are not present.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get user when user with provided id does not exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`The specified user with id ${userId} does not exist`);
    });

    command.action(logger, { options: { debug: false, id: userId } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified user with id ${userId} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get user when user with provided user name does not exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${encodeURIComponent(userName)}'`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`The specified user with user name ${userName} does not exist`);
    });

    command.action(logger, { options: { debug: false, userName: userName } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified user with user name ${userName} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get user when user with provided email does not exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=mail eq '${encodeURIComponent(userName)}'`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`The specified user with email ${userName} does not exist`);
    });

    command.action(logger, { options: { debug: false, email: userName } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified user with email ${userName} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when multiple users with the specified email found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if ((opts.url as string).indexOf('https://graph.microsoft.com/v1.0/users?$filter') > -1) {
        return Promise.resolve({
          value: [
            resultValue,
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', userPrincipalName: 'DebraB@contoso.onmicrosoft.com' }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        email: userName
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Multiple users with email ${userName} found. Please disambiguate (user names): ${userName}, DebraB@contoso.onmicrosoft.com or (ids): ${userId}, 9b1b1e42-794b-4c71-93ac-5ed92488b67f`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if id or email or userName options are not passed', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id, email, and userName options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", email: "john.doe@contoso.onmicrosoft.com", userName: "i:0#.f|membership|john.doe@contoso.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and email options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", email: "john.doe@contoso.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and userName options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", userName: "john.doe@contoso.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both email and userName options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { email: "jonh.deo@contoso.com", userName: "john.doe@contoso.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the userName is specified', async () => {
    const actual = await command.validate({ options: { userName: 'john.doe@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the email is specified', async () => {
    const actual = await command.validate({ options: { email: 'john.doe@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});