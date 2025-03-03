import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { cli } from '../../../cli/cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import { CommandError } from '../../../Command.js';
import request from '../../../request.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import commands from '../commands.js';
import command from './file-move.js';

describe(commands.MOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const defaultPostStub = () => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      const url: string = opts.url as string;

      if (
        url === 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/copy' ||
        url === 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/copy?@microsoft.graph.conflictBehavior=replace' ||
        url === 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/copy?@microsoft.graph.conflictBehavior=rename'
      ) {
        return Promise.resolve({
          status: 202,
          headers: {
            location: "https://contoso.sharepoint.com/_api/v2.1/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/operations/dd4f0967-51dd-4d4a-87d0-617b9ca1df6c"
          }
        });
      }

      if (
        url === 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYV/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRA/copy'
      ) {
        return Promise.resolve({
          status: 202,
          headers: {
            location: "https://contoso.sharepoint.com/_api/v2.1/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYV/operations/dd4f0967-51dd-4d4a-87d0-617b9ca1df6c"
          }
        });
      }

      throw 'Invalid request';
    });
  };

  const defaultPatchStub = () => {
    return sinon.stub(request, 'patch').callsFake(async (opts) => {
      const url: string = opts.url as string;

      if (
        url === 'https://graph.microsoft.com/v1.0//drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ' ||
        url === 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ?@microsoft.graph.conflictBehavior=rename'
      ) {
        return Promise.resolve({
          status: 202
        });
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
            "id": "contoso.sharepoint.com,f89617dc-8c96-4044-954d-4c690ce5fcd9,e0f173ba-2d5e-4098-a9d7-49530181130c"
          };
        case 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root:/file.pdf?$select=id':
          return {
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          };
        case 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root?$select=id':
          return {
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          };
        case 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root:/NewFolder?$select=id':
          return {
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          };
        case 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYV/root:/file.pdf?$select=id':
          return {
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRA"
          };
        case 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYV/root?$select=id':
          return {
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRA"
          };
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return {
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              }
            ]
          };
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,f89617dc-8c96-4044-954d-4c690ce5fcd9,e0f173ba-2d5e-4098-a9d7-49530181130c/drives?$select=webUrl,id':
          return {
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYV",
                "webUrl": "https://contoso.sharepoint.com/teams/finance/Shared%20Documents"
              }
            ]
          };
        case 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root:/folder1?$select=id':
          return {
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          };
        case 'https://contoso.sharepoint.com/_api/v2.1/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/operations/dd4f0967-51dd-4d4a-87d0-617b9ca1df6c':
          return {
            "@odata.context": "https://contoso.sharepoint.com/_api/v2.1/$metadata#drives('b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU')/operations/$entity",
            "id": "dd4f0967-51dd-4d4a-87d0-617b9ca1df6c",
            "createdDateTime": "0001-01-01T00:00:00Z",
            "lastActionDateTime": "0001-01-01T00:00:00Z",
            "percentageComplete": 100,
            "percentComplete": 100,
            "resourceId": "01QBVM576FDUYWKYENPBALUPTXALT65BFQ",
            "resourceLocation": "https://contoso.sharepoint.com/_api/v2.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYV/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRA",
            "status": "completed"
          };
        case 'https://contoso.sharepoint.com/_api/v2.1/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYV/operations/dd4f0967-51dd-4d4a-87d0-617b9ca1df6c':
          return {
            "@odata.context": "https://contoso.sharepoint.com/_api/v2.1/$metadata#drives('b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU')/operations/$entity",
            "id": "dd4f0967-51dd-4d4a-87d0-617b9ca1df6c",
            "createdDateTime": "0001-01-01T00:00:00Z",
            "lastActionDateTime": "0001-01-01T00:00:00Z",
            "status": "failed",
            "error": {
              "code": "nameAlreadyExists",
              "message": "Name already exists"
            }
          };

        default:
          throw `Invalid GET request: ${url}`;
      }
    });
  };

  const defaultDeleteStub = () => {
    return sinon.stub(request, 'delete').callsFake(async (opts) => {
      const url: string = opts.url as string;

      if (
        url === 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ' ||
        url === 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYV/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRA'
      ) {
        return;
      }

      throw 'Invalid request';
    });
  };

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
    (command as any).pollingInterval = 0;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      request.patch,
      request.delete,
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MOVE);
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

  it('passes validation with valid options', async () => {
    const actual = await command.validate({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: '/Shared Documents/file.pdf',
        targetUrl: '/teams/finance/Shared Documents',
        newName: 'file1',
        nameConflictBehavior: 'rename'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('moves a file by renaming when the same name already exists.', async () => {
    defaultGetStub();
    defaultPatchStub();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: '/Shared Documents/file.pdf',
        targetUrl: '/Shared Documents/NewFolder',
        nameConflictBehavior: 'rename',
        newName: 'file_renamed.pdf'
      }
    });
  });

  it('moves a file by renaming with a new name and no extension when a file with the same name already exists.', async () => {
    defaultGetStub();
    defaultPatchStub();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: '/Shared Documents/file.pdf',
        targetUrl: '/Shared Documents/NewFolder',
        nameConflictBehavior: 'rename',
        newName: 'file_renamed'
      }
    });
  });

  it('moves a file to a document library in another site collection', async () => {
    const getStub: sinon.SinonStub = defaultGetStub();
    const postStub: sinon.SinonStub = defaultPostStub();
    const deleteStub: sinon.SinonStub = defaultDeleteStub();

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
    assert(deleteStub.called);
  });

  it('moves a file to a document library in another site collection and waits for command to complete.', async () => {
    let callCount = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return {
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          };
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/teams/finance?$select=id':
          return {
            "id": "contoso.sharepoint.com,f89617dc-8c96-4044-954d-4c690ce5fcd9,e0f173ba-2d5e-4098-a9d7-49530181130c"
          };
        case 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root:/file.pdf?$select=id':
          return {
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          };
        case 'https://graph.microsoft.com/v1.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYV/root?$select=id':
          return {
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRA"
          };
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return {
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              }
            ]
          };
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,f89617dc-8c96-4044-954d-4c690ce5fcd9,e0f173ba-2d5e-4098-a9d7-49530181130c/drives?$select=webUrl,id':
          return {
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYV",
                "webUrl": "https://contoso.sharepoint.com/teams/finance/Shared%20Documents"
              }
            ]
          };
        case 'https://contoso.sharepoint.com/_api/v2.1/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/operations/dd4f0967-51dd-4d4a-87d0-617b9ca1df6c':
          if (callCount === 0) {
            callCount++;
            return {
              "@odata.context": "https://contoso.sharepoint.com/_api/v2.1/$metadata#drives('b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU')/operations/$entity",
              "id": "dd4f0967-51dd-4d4a-87d0-617b9ca1df6c",
              "createdDateTime": "0001-01-01T00:00:00Z",
              "lastActionDateTime": "0001-01-01T00:00:00Z",
              "percentageComplete": 98.64040641410281,
              "percentComplete": 98.64040641410281,
              "status": "inProgress"
            };
          }
          else {
            return {
              "@odata.context": "https://contoso.sharepoint.com/_api/v2.1/$metadata#drives('b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU')/operations/$entity",
              "id": "dd4f0967-51dd-4d4a-87d0-617b9ca1df6c",
              "createdDateTime": "0001-01-01T00:00:00Z",
              "lastActionDateTime": "0001-01-01T00:00:00Z",
              "percentageComplete": 100,
              "percentComplete": 100,
              "resourceId": "01QBVM576FDUYWKYENPBALUPTXALT65BFQ",
              "resourceLocation": "https://contoso.sharepoint.com/_api/v2.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYV/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRA",
              "status": "completed"
            };
          }

        default:
          throw `Invalid GET request: ${url}`;
      }
    });
    defaultPostStub();
    const deleteStub: sinon.SinonStub = defaultDeleteStub();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/',
        sourceUrl: '/shared documents/file.pdf',
        targetUrl: '/teams/finance/Shared Documents',
        verbose: true
      }
    });

    assert(deleteStub.called);
  });

  it('handles error if file already exists.', async () => {
    defaultGetStub();
    defaultPostStub();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/teams/finance',
        sourceUrl: '/teams/finance/shared documents/file.pdf',
        targetUrl: '/Shared Documents'
      }
    }), new CommandError(`Name already exists`));
  });

  it('moves a folder by renaming when a folder with the same name already exists.', async () => {
    defaultGetStub();
    defaultPostStub();
    defaultDeleteStub();

    await assert.doesNotReject(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: '/Shared Documents/folder1',
        targetUrl: '/teams/finance/Shared Documents',
        nameConflictBehavior: 'rename',
        newName: 'folder1_renamed'
      }
    }));
  });

  it('moves file from source to target (absolute URLs)', async () => {
    defaultGetStub();
    defaultPostStub();
    defaultDeleteStub();

    await assert.doesNotReject(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'https://contoso.sharepoint.com/Shared Documents/file.pdf',
        targetUrl: 'https://contoso.sharepoint.com/teams/finance/Shared Documents',
        verbose: true
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
    }), new CommandError(`Drive 'https://contoso.sharepoint.com/invalid/file.pdf' not found`));
  });
});