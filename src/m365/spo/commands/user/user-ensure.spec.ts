import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './user-ensure.js';

describe(commands.USER_ENSURE, () => {
  const validUserName = 'john@contoso.com';
  const validEntraId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
  const validWebUrl = 'https://contoso.sharepoint.com';
  const ensuredUserResponse = {
    Id: 35,
    IsHiddenInUI: false,
    LoginName: `i:0#.f|membership|${validUserName}`,
    Title: 'John Doe',
    PrincipalType: 1,
    Email: 'john@contoso.com',
    Expiration: '',
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: false,
    UserId: {
      NameId: '1003200274f51d2d',
      NameIdIssuer: 'urn:federation:microsoftonline'
    },
    UserPrincipalName: validUserName
  };

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
    commandInfo = cli.getCommandInfo(command);
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
      request.post,
      entraUser.getUpnByUserId
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_ENSURE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('ensures user for a specific web by userPrincipalName', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureuser`) {
        return ensuredUserResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: validWebUrl, userName: validUserName } });
    assert(loggerLogSpy.calledWith(ensuredUserResponse));
  });

  it('ensures user for a specific web by entraId', async () => {
    sinon.stub(aadUser, 'getUpnByUserId').callsFake(async () => {
      return validUserName;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureuser`) {
        return ensuredUserResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: validWebUrl, entraId: validEntraId } });
    assert(loggerLogSpy.calledWith(ensuredUserResponse));
  });

  it('ensures user for a specific web by aadId', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').callsFake(async () => {
      return validUserName;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureuser`) {
        return ensuredUserResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: validWebUrl, aadId: validEntraId } });
    assert(loggerLogSpy.calledWith(ensuredUserResponse));
  });

  it('throws error message when no user was found with a specific id', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').callsFake(async (id) => {
      throw {
        "error": {
          "error": {
            "code": "Request_ResourceNotFound",
            "message": `Resource '${id}' does not exist or one of its queried reference-property objects are not present.`,
            "innerError": {
              "date": "2023-02-17T22:44:21",
              "request-id": "25519ac1-8f24-46a7-90b0-19baace49a7a",
              "client-request-id": "25519ac1-8f24-46a7-90b0-19baace49a7a"
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, webUrl: validWebUrl, entraId: validEntraId } }), new CommandError(`Resource '${validEntraId}' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('throws error message when no user was found with a specific user name', async () => {
    const error = {
      'error': {
        'odata.error': {
          'code': '-2146232832, Microsoft.SharePoint.SPException',
          'message': {
            'lang': 'en-US',
            'value': `The specified user ${validUserName} could not be found.`
          }
        }
      }
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureuser`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, webUrl: validWebUrl, userName: validUserName } }), new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if webUrl is not a valid url', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', entraId: validEntraId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if entraId is not a valid id', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, entraId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if aadId is not a valid id', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, aadId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid user principal name', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url is valid and entraId is a valid id', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, entraId: validEntraId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and aadId is a valid id', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, aadId: validEntraId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and userName is a valid user principal name', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, userName: validUserName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
}); 
