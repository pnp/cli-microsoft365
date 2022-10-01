import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./notebook-list');

describe(commands.NOTEBOOK_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.NOTEBOOK_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both userId and userName options are passed', async () => {
    const actual = await command.validate({
      options:
      {
        userId: '2609af39-7775-4f94-a3dc-0dd67657e900',
        userName: 'Documents'
      }
    }, commandInfo);
    assert.strictEqual(actual, 'Specify either userId or userName, but not both');
  });

  it('fails validation if both groupId and groupName options are passed', async () => {
    const actual = await command.validate({
      options:
      {
        groupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4',
        groupName: 'MyGroup'
      }
    }, commandInfo);
    assert.strictEqual(actual, 'Specify either groupId or groupName, but not both');
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/onenote/notebooks`) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/2609af39-7775-4f94-a3dc-0dd67657e900/onenote/notebooks`) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks`) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { groupName: 'MyGroup' } } as any), new CommandError('An error has occurred'));
  });

  it('lists Microsoft OneNote notebooks in group by name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4",
              "description": "MyGroup",
              "displayName": "MyGroup"
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks`) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/sites/`) > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/testsite' } } as any), new CommandError('An error has occurred'));
  });

  it('lists Microsoft OneNote notebooks for site', async () => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return Promise.resolve({
            "id": "contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2",
            "name": "testsite",
            "webUrl": "https://contoso.sharepoint.com/sites/testsite",
            "displayName": "testsite"
          });
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2/onenote/notebooks`) {
          return Promise.resolve({
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
          });
        }

        return Promise.reject('Invalid request');
      });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/testsite' } });
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/user1@contoso.onmicrosoft.com/onenote/notebooks`) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
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
