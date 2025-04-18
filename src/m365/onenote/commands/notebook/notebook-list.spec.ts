import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './notebook-list.js';

describe(commands.NOTEBOOK_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.NOTEBOOK_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['createdDateTime', 'displayName', 'id']);
  });

  it('fails validation if webUrl is not a valid webUrl', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if no option specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('lists Microsoft OneNote notebooks for the currently logged in user (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/onenote/notebooks`) {
        return {
          "value": [
            {
              "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
              "createdDateTime": "2021-11-15T10:27:22Z",
              "displayName": "Meeting Notes",
              "lastModifiedDateTime": "2021-11-15T10:27:22Z"
            },
            {
              "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
              "createdDateTime": "2020-01-13T17:52:03Z",
              "displayName": "My Notebook",
              "lastModifiedDateTime": "2020-01-13T17:52:03Z"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
        "createdDateTime": "2021-11-15T10:27:22Z",
        "displayName": "Meeting Notes",
        "lastModifiedDateTime": "2021-11-15T10:27:22Z"
      },
      {
        "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
        "createdDateTime": "2020-01-13T17:52:03Z",
        "displayName": "My Notebook",
        "lastModifiedDateTime": "2020-01-13T17:52:03Z"
      }
    ]));
  });

  it('lists Microsoft OneNote notebooks for user by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/2609af39-7775-4f94-a3dc-0dd67657e900/onenote/notebooks`) {
        return {
          "value": [
            {
              "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
              "createdDateTime": "2021-11-15T10:27:22Z",
              "displayName": "Meeting Notes",
              "lastModifiedDateTime": "2021-11-15T10:27:22Z"
            },
            {
              "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
              "createdDateTime": "2020-01-13T17:52:03Z",
              "displayName": "My Notebook",
              "lastModifiedDateTime": "2020-01-13T17:52:03Z"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: '2609af39-7775-4f94-a3dc-0dd67657e900' } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
        "createdDateTime": "2021-11-15T10:27:22Z",
        "displayName": "Meeting Notes",
        "lastModifiedDateTime": "2021-11-15T10:27:22Z"
      },
      {
        "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
        "createdDateTime": "2020-01-13T17:52:03Z",
        "displayName": "My Notebook",
        "lastModifiedDateTime": "2020-01-13T17:52:03Z"
      }
    ]));
  });

  it('lists Microsoft OneNote notebooks in group by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks`) {
        return {
          "value": [
            {
              "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
              "createdDateTime": "2021-11-15T10:27:22Z",
              "displayName": "Meeting Notes",
              "lastModifiedDateTime": "2021-11-15T10:27:22Z"
            },
            {
              "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
              "createdDateTime": "2020-01-13T17:52:03Z",
              "displayName": "My Notebook",
              "lastModifiedDateTime": "2020-01-13T17:52:03Z"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4' } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
        "createdDateTime": "2021-11-15T10:27:22Z",
        "displayName": "Meeting Notes",
        "lastModifiedDateTime": "2021-11-15T10:27:22Z"
      },
      {
        "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
        "createdDateTime": "2020-01-13T17:52:03Z",
        "displayName": "My Notebook",
        "lastModifiedDateTime": "2020-01-13T17:52:03Z"
      }
    ]));
  });

  it('handles error when retrieving Microsoft OneNote notebooks in group by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { groupName: 'MyGroup' } } as any), new CommandError('An error has occurred'));
  });

  it('lists Microsoft OneNote notebooks in group by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return {
          "value": [
            {
              "id": "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4",
              "description": "MyGroup",
              "displayName": "MyGroup"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks`) {
        return {
          "value": [
            {
              "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
              "createdDateTime": "2021-11-15T10:27:22Z",
              "displayName": "Meeting Notes",
              "lastModifiedDateTime": "2021-11-15T10:27:22Z"
            },
            {
              "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
              "createdDateTime": "2020-01-13T17:52:03Z",
              "displayName": "My Notebook",
              "lastModifiedDateTime": "2020-01-13T17:52:03Z"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: 'MyGroup' } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
        "createdDateTime": "2021-11-15T10:27:22Z",
        "displayName": "Meeting Notes",
        "lastModifiedDateTime": "2021-11-15T10:27:22Z"
      },
      {
        "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
        "createdDateTime": "2020-01-13T17:52:03Z",
        "displayName": "My Notebook",
        "lastModifiedDateTime": "2020-01-13T17:52:03Z"
      }
    ]));
  });

  it('handles error when retrieving Microsoft OneNote notebooks for site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/sites/`) > -1) {
        return Promise.reject('An error has occurred');
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/testsite' } } as any), new CommandError('An error has occurred'));
  });

  it('lists Microsoft OneNote notebooks for site', async () => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return {
            "id": "contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2",
            "name": "testsite",
            "webUrl": "https://contoso.sharepoint.com/sites/testsite",
            "displayName": "testsite"
          };
        }
        throw 'Invalid request';
      });

    getRequestStub.onCall(1)
      .callsFake(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2/onenote/notebooks`) {
          return {
            "value": [
              {
                "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
                "createdDateTime": "2021-11-15T10:27:22Z",
                "displayName": "Meeting Notes",
                "lastModifiedDateTime": "2021-11-15T10:27:22Z"
              },
              {
                "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
                "createdDateTime": "2020-01-13T17:52:03Z",
                "displayName": "My Notebook",
                "lastModifiedDateTime": "2020-01-13T17:52:03Z"
              }
            ]
          };
        }

        throw 'Invalid request';
      });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/testsite', debug: true } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
        "createdDateTime": "2021-11-15T10:27:22Z",
        "displayName": "Meeting Notes",
        "lastModifiedDateTime": "2021-11-15T10:27:22Z"
      },
      {
        "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
        "createdDateTime": "2020-01-13T17:52:03Z",
        "displayName": "My Notebook",
        "lastModifiedDateTime": "2020-01-13T17:52:03Z"
      }
    ]));
  });

  it('lists Microsoft OneNote notebooks for user by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/user1@contoso.onmicrosoft.com/onenote/notebooks`) {
        return {
          "value": [
            {
              "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
              "createdDateTime": "2021-11-15T10:27:22Z",
              "displayName": "Meeting Notes",
              "lastModifiedDateTime": "2021-11-15T10:27:22Z"
            },
            {
              "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
              "createdDateTime": "2020-01-13T17:52:03Z",
              "displayName": "My Notebook",
              "lastModifiedDateTime": "2020-01-13T17:52:03Z"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: 'user1@contoso.onmicrosoft.com' } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "1-99a44a87-c92f-495a-8295-3ab308387821",
        "createdDateTime": "2021-11-15T10:27:22Z",
        "displayName": "Meeting Notes",
        "lastModifiedDateTime": "2021-11-15T10:27:22Z"
      },
      {
        "id": "1-1c1fbd21-1d48-4057-bfb1-ce41b4f7d624",
        "createdDateTime": "2020-01-13T17:52:03Z",
        "displayName": "My Notebook",
        "lastModifiedDateTime": "2020-01-13T17:52:03Z"
      }
    ]));
  });
});
