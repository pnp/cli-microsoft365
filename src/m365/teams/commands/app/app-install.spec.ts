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
const command: Command = require('./app-install');

describe(commands.APP_INSTALL, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APP_INSTALL), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when neither teamId, userId nor userName are specified', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when teamId and userId are specified', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        teamId: '00000000-0000-0000-0000-000000000000',
        userId: '00000000-0000-0000-0000-000000000000'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when teamId and userName are specified', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        teamId: '00000000-0000-0000-0000-000000000000',
        userName: 'steve@contoso.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when userId and userName are specified', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        userId: '00000000-0000-0000-0000-000000000000',
        userName: 'steve@contoso.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid',
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        id: 'not-c49b-4fd4-8223-28f0ac3a6402',
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        userId: 'not-c49b-4fd4-8223-28f0ac3a6402',
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id and teamId are correct', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c6',
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the id and userId are correct', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c6',
        userId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the id and userName are correct', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c6',
        userName: 'steve@contoso.com'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds app from the catalog to a Microsoft Team', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/c527a470-a882-481c-981c-ee6efaba85c7/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/4440558e-8c73-4597-abc7-3644a64c4bce"}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        teamId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    });
    assert.strictEqual(log.length, 0);
  });

  it('installs app from the catalog the user specified with userId', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=id eq 'c527a470-a882-481c-981c-ee6efaba85c7'`) {
        return Promise.resolve({
          "value": [
            {
              "businessPhones": [
                "425-555-0100"
              ],
              "displayName": "MOD Administrator",
              "givenName": "MOD",
              "jobTitle": null,
              "mail": "admin@contoso.OnMicrosoft.com",
              "mobilePhone": "425-555-0101",
              "officeLocation": null,
              "preferredLanguage": "en-US",
              "surname": "Administrator",
              "userPrincipalName": "admin@contoso.onmicrosoft.com",
              "id": "c527a470-a882-481c-981c-ee6efaba85c7"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/4440558e-8c73-4597-abc7-3644a64c4bce"}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    });
    assert.strictEqual(log.length, 0);
  });

  it('installs app from the catalog the user specified with userId (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=id eq 'c527a470-a882-481c-981c-ee6efaba85c7'`) {
        return Promise.resolve({
          "value": [
            {
              "businessPhones": [
                "425-555-0100"
              ],
              "displayName": "MOD Administrator",
              "givenName": "MOD",
              "jobTitle": null,
              "mail": "admin@contoso.OnMicrosoft.com",
              "mobilePhone": "425-555-0101",
              "officeLocation": null,
              "preferredLanguage": "en-US",
              "surname": "Administrator",
              "userPrincipalName": "admin@contoso.onmicrosoft.com",
              "id": "c527a470-a882-481c-981c-ee6efaba85c7"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/4440558e-8c73-4597-abc7-3644a64c4bce"}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce',
        debug: true
      }
    });
  });

  it('installs app from the catalog the user specified with userName', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/steve%40contoso.com/teamwork/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/4440558e-8c73-4597-abc7-3644a64c4bce"}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        userName: 'steve@contoso.com',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    });
    assert.strictEqual(log.length, 0);
  });

  it('correctly handles error while installing Teams app', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, {
      options: {
        teamId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it(`correctly handles error when trying to install an app for a user that doesn't exist (invalid user name)`, async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "NotFound",
          "message": "Failed to find user with id 'steve@contoso.com' in the tenant",
          "innerError": {
            "date": "2022-02-14T12:14:15",
            "request-id": "1d6fc213-9f35-4cb3-b496-3d8b10aebdfa",
            "client-request-id": "1d6fc213-9f35-4cb3-b496-3d8b10aebdfa"
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: { userName: 'steve@contoso.com', id: '4440558e-8c73-4597-abc7-3644a64c4bce' } } as any), new CommandError("Failed to find user with id 'steve@contoso.com' in the tenant"));
  });

  it(`correctly handles error when trying to install an app for a user that doesn't exist (invalid user ID)`, async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=id eq 'c527a470-a882-481c-981c-ee6efaba85c7'`) {
        return Promise.reject({
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource 'c527a470-a882-481c-981c-ee6efaba85c7' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "date": "2022-02-14T13:27:37",
              "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
              "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake(() => Promise.reject('Invalid request'));

    await assert.rejects(command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    } as any), new CommandError("User with ID c527a470-a882-481c-981c-ee6efaba85c7 not found. Original error: Resource 'c527a470-a882-481c-981c-ee6efaba85c7' does not exist or one of its queried reference-property objects are not present."));
  });

  it(`correctly handles error when trying to install an app for a user that doesn't exist (invalid user ID; debug)`, async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=id eq 'c527a470-a882-481c-981c-ee6efaba85c7'`) {
        return Promise.reject({
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource 'c527a470-a882-481c-981c-ee6efaba85c7' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "date": "2022-02-14T13:27:37",
              "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
              "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake(() => Promise.reject('Invalid request'));

    await assert.rejects(command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce',
        debug: true
      }
    } as any), new CommandError("User with ID c527a470-a882-481c-981c-ee6efaba85c7 not found. Original error: Resource 'c527a470-a882-481c-981c-ee6efaba85c7' does not exist or one of its queried reference-property objects are not present."));
  });
});
