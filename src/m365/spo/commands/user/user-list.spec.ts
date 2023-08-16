import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './user-list.js';

describe(commands.USER_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
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
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'Title', 'LoginName']);
  });

  it('retrieves lists of site users with output option json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/siteusers') > -1) {
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
            "IsSiteAdmin": true,
            "UserId": { "NameId": "10010001b0c19a2", "NameIdIssuer": "urn:federation:microsoftonline" },
            "UserPrincipalName": "john.doe@mytenant.onmicrosoft.com"
          },
          {
            "Id": 7,
            "IsHiddenInUI": false,
            "LoginName": "i:0#.f|membership|abc@mytenant.onmicrosoft.com",
            "Title": "FName Lname",
            "PrincipalType": 1,
            "Email": "abc@mytenant.onmicrosoft.com",
            "Expiration": "",
            "IsEmailAuthenticationGuestUser": false,
            "IsShareByEmailGuestUser": false,
            "IsSiteAdmin": false,
            "UserId": { "NameId": "1003201096515567", "NameIdIssuer": "urn:federation:microsoftonline" },
            "UserPrincipalName": "abc@mytenant.onmicrosoft.com"
          }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledWith([{
      Id: 6,
      IsHiddenInUI: false,
      LoginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
      Title: "John Doe",
      PrincipalType: 1,
      Email: "john.deo@mytenant.onmicrosoft.com",
      Expiration: "",
      IsEmailAuthenticationGuestUser: false,
      IsShareByEmailGuestUser: false,
      IsSiteAdmin: true,
      UserId: { NameId: "10010001b0c19a2", NameIdIssuer: "urn:federation:microsoftonline" },
      UserPrincipalName: "john.doe@mytenant.onmicrosoft.com"
    },
    {
      Id: 7,
      IsHiddenInUI: false,
      LoginName: "i:0#.f|membership|abc@mytenant.onmicrosoft.com",
      Title: "FName Lname",
      PrincipalType: 1,
      Email: "abc@mytenant.onmicrosoft.com",
      Expiration: "",
      IsEmailAuthenticationGuestUser: false,
      IsShareByEmailGuestUser: false,
      IsSiteAdmin: false,
      UserId: { NameId: "1003201096515567", NameIdIssuer: "urn:federation:microsoftonline" },
      UserPrincipalName: "abc@mytenant.onmicrosoft.com"
    }]));
  });

  it('retrieves lists of site users', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/siteusers') > -1) {
        return {
          "value": [{
            "Id": 6,
            "Title": "John Doe",
            "Email": "john.deo@mytenant.onmicrosoft.com",
            "LoginName": "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
          },
          {
            "Id": 7,
            "Title": "FName Lname",
            "Email": "abc@mytenant.onmicrosoft.com",
            "LoginName": "i:0#.f|membership|abc@mytenant.onmicrosoft.com"
          }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledWith([{
      "Id": 6,
      "Title": "John Doe",
      "Email": "john.deo@mytenant.onmicrosoft.com",
      "LoginName": "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
    },
    {
      "Id": 7,
      "Title": "FName Lname",
      "Email": "abc@mytenant.onmicrosoft.com",
      "LoginName": "i:0#.f|membership|abc@mytenant.onmicrosoft.com"
    }]));
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

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com' } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url is valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
