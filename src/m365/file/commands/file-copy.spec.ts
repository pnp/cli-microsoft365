import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { Cli } from '../../../cli/Cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import { CommandError } from '../../../Command.js';
import request from '../../../request.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import commands from '../commands.js';
import command from './file-copy.js';

describe(commands.COPY, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const defaultPostStub = () => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      const url: string = opts.url as string;

      if (
        url === 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/copy' ||
        url === 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/copy?@microsoft.graph.conflictBehavior=replace' ||
        url === 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/copy?@microsoft.graph.conflictBehavior=rename'
      ) {
        return Promise.resolve({ response: { status: 202 } });
      }

      throw 'Invalid request';
    });
  };

  const defaultGetStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'get').callsFake(async opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return {
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          };
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/teams/finance?$select=id':
          return {
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          };
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/teams/invalid?$select=id':
          throw {
            "error": {
              "code": "itemNotFound",
              "message": "Requested site could not be found",
              "innerError": {
                "date": "2023-09-20T19:13:33",
                "request-id": "0a1558c8-8078-4649-a570-5fd3d6518e83",
                "client-request-id": "0a1558c8-8078-4649-a570-5fd3d6518e83"
              }
            }
          };
        case 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root:/file.pdf?$select=id':
          return {
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          };
        case 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root?$select=id':
          return {
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          };
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/teams/finance?$select=id':
          return {
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          };
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return {
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/teams/finance/Shared%20Documents"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              }
            ]
          };
        case 'https://graph.microsoft.com/v1.0/sites/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/drives?$select=webUrl,id':
          return {
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KCswD4M9qeR6qB9K5J5Kvp",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LCxmZShRH-S4chwRsWoq23",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LxywkjzYwYSqUtcpywFv6S",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          };
        default:
          throw `Invalid GET request: ${url}`;
      }
    });
  };

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      request.put,
      fs.existsSync,
      fs.readFileSync
    ]);
    (command as any).sourceFileGraphUrl = undefined;
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.COPY);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if nameConflictBehavior is not a valid option', async () => {
    const actual = await command.validate({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: '/Shared Documents/file.pdf',
        targetUrl: '/teams/finance/Shared Documents',
        nameConflictBehavior: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options: {
        webUrl: 'foo',
        sourceUrl: '/Shared Documents/file.pdf',
        targetUrl: '/teams/finance/Shared Documents'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('copies file from source to target', async () => {
    const getStub: sinon.SinonStub = defaultGetStub();
    const postStub: sinon.SinonStub = defaultPostStub();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: '/Shared Documents/file.pdf',
        targetUrl: '/teams/finance/Shared Documents',
        verbose: true
      }
    });

    assert(getStub.called);
    assert(postStub.called);
  });

  it('copies file from source to target (absolute URLs)', async () => {
    defaultGetStub();
    defaultPostStub();

    await assert.doesNotReject(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'https://contoso.sharepoint.com/Shared Documents/file.pdf',
        targetUrl: 'https://contoso.sharepoint.com/teams/finance/Shared Documents',
        verbose: true
      }
    }));
  });

  it('copies file from source to the root site', async () => {
    defaultGetStub();
    defaultPostStub();

    await assert.doesNotReject(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/teams/finance',
        sourceUrl: '/teams/finance/Shared Documents/file.pdf',
        targetUrl: '/Shared Documents',
        debug: true
      }
    }));
  });

  it('replaces an existing file with a new one when a file with the same name already exists.', async () => {
    defaultGetStub();
    defaultPostStub();

    await assert.doesNotReject(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: '/Shared Documents/file.pdf',
        targetUrl: '/teams/finance/Shared Documents',
        nameConflictBehavior: 'replace'
      }
    }));
  });

  it('copies file by renaming when a file with the same name already exists.', async () => {
    defaultGetStub();
    defaultPostStub();

    await assert.doesNotReject(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: '/Shared Documents/file.pdf',
        targetUrl: '/teams/finance/Shared Documents',
        nameConflictBehavior: 'rename',
        newName: 'file1.pdf'
      }
    }));

  });

  it('handles error if unexpected error occurs', async () => {
    defaultGetStub();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: '/invalid/file.pdf',
        targetUrl: '/teams/finance/Shared Documents'
      }
    }), new CommandError(`Document library '/invalid/file.pdf' not found`));
  });
});