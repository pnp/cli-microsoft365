import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./user-get');

describe(commands.USER_GET, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves user by id with output option json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/siteusers/GetById') > -1) {
        return {
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
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 1
      }
    });
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
  });


  it('retrieves user by email with output option json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/siteusers/GetByEmail') > -1) {
        return {
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
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        email: "john.deo@mytenant.onmicrosoft.com"
      }
    });
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
  });

  it('retrieves user by userName with output option json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/siteusers/GetByLoginName') > -1) {
        return {
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
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        userName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    });
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
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        userName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    } as any), new CommandError('An error has occurred'));
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

  it('fails validation if id or email or userName options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id, email and userName options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, email: "jonh.deo@mytenant.com", userName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id and email both are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, email: "jonh.deo@mytenant.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id and userName options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, userName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if email and userName options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', email: "jonh.deo@mytenant.com", userName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
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

  it('passes validation if the url is valid and userName is passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', userName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
    assert.strictEqual(actual, true);
  });
}); 
