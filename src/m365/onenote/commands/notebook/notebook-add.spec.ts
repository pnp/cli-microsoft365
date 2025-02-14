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
import command from './notebook-add.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';

describe(commands.NOTEBOOK_ADD, () => {
  const name = 'My Notebook';
  const addResponse = {
    id: '1-2ae2e5d0-2857-4b1a-99d3-cc5426799438',
    self: 'https://graph.microsoft.com/v1.0/users/fe36f75e-c103-410b-a18a-2bf6df06ac3a/onenote/notebooks/1-2ae2e5d0-2857-4b1a-99d3-cc5426799438',
    createdDateTime: '2024-04-05T17:30:28Z',
    displayName: name,
    lastModifiedDateTime: '2024-04-05T17:30:28Z',
    isDefault: false,
    userRole: 'Owner',
    isShared: false,
    sectionsUrl: 'https://graph.microsoft.com/v1.0/users/fe36f75e-c103-410b-a18a-2bf6df06ac3a/onenote/notebooks/1-2ae2e5d0-2857-4b1a-99d3-cc5426799438/sections',
    sectionGroupsUrl: 'https://graph.microsoft.com/v1.0/users/fe36f75e-c103-410b-a18a-2bf6df06ac3a/onenote/notebooks/1-2ae2e5d0-2857-4b1a-99d3-cc5426799438/sectionGroups',
    createdBy: {
      user: {
        id: 'fe36f75e-c103-410b-a18a-2bf6df06ac3a',
        displayName: 'John Doe'
      }
    },
    lastModifiedBy: {
      user: {
        id: 'fe36f75e-c103-410b-a18a-2bf6df06ac3a',
        displayName: 'John Doe'
      }
    },
    links: {
      oneNoteClientUrl: {
        href: 'onenote:https://contoso2-my.sharepoint.com/personal/john_contoso2_onmicrosoft_com/Documents/Notebooks/Dummy'
      },
      oneNoteWebUrl: {
        href: 'https://contoso2-my.sharepoint.com/personal/john_contoso2_onmicrosoft_com/Documents/Notebooks/Dummy'
      }
    }
  };

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.NOTEBOOK_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if name contains invalid characters', async () => {
    const actual = await command.validate({ options: { name: 'My notebook /' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if name is longer than 128 characters', async () => {
    const longString = 'x'.repeat(129);
    const actual = await command.validate({ options: { name: longString } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid webUrl', async () => {
    const actual = await command.validate({ options: { name: name, webUrl: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { name: name, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { name: name, groupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if no option but name specified', async () => {
    const actual = await command.validate({ options: { name: name } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds notebook for the currently logged in user', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/onenote/notebooks`) {
        return addResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: name, verbose: true } });
    assert(loggerLogSpy.calledWith(addResponse));
  });

  it('adds notebook for user by id', async () => {
    const userId = '2609af39-7775-4f94-a3dc-0dd67657e900';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onenote/notebooks`) {
        return addResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: name, userId: userId, verbose: true } });
    assert(loggerLogSpy.calledWith(addResponse));
  });

  it('adds notebook in group by id', async () => {
    const groupId = '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}/onenote/notebooks`) {
        return addResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: name, groupId: groupId, verbose: true } });
    assert(loggerLogSpy.calledWith(addResponse));
  });

  it('adds notebook in group by name', async () => {
    const groupId = '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4';
    const groupName = 'My group';
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves(groupId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}/onenote/notebooks`) {
        return addResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: name, groupName: groupName, verbose: true } });
    assert(loggerLogSpy.calledWith(addResponse));
  });

  it('adds notebook for site', async () => {
    const siteUrl = 'https://contoso.sharepoint.com/sites/testsite';
    const siteId = 'contoso.sharepoint.com,2C712604-1370-44E7-A1F5-426573FDA80A,2D2244C3-251A-49EA-93A8-39E1C3A060FE';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${siteId}/onenote/notebooks`) {
        return addResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getSpoGraphSiteId').resolves(siteId);

    await command.action(logger, { options: { name: name, webUrl: siteUrl, verbose: true } });
    assert(loggerLogSpy.calledWith(addResponse));
  });

  it('adds notebook for user by name', async () => {
    const userName = 'john@contoso.com';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userName}/onenote/notebooks`) {
        return addResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: name, userName: userName, verbose: true } });
    assert(loggerLogSpy.calledWith(addResponse));
  });

  it('handles error when adding notebook fails when it already exists', async () => {
    const error = {
      error: {
        code: '20117',
        message: 'An item with this name already exists in this location.',
        innerError: {
          date: '2024-04-05T17:49:42',
          'request-id': '47cd5f47-2158-4c43-ae0a-22e3b9073e7d',
          'client-request-id': '47cd5f47-2158-4c43-ae0a-22e3b9073e7d'
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/onenote/notebooks`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: name, verbose: true } } as any), new CommandError(error.error.message));
  });
});
