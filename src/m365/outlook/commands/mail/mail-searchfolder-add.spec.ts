import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './mail-searchfolder-add.js';

describe(commands.MAIL_SEARCHFOLDER_ADD, () => {
  const userId = 'ae0e8388-cd70-427f-9503-c57498ee3337';
  const userName = 'john.doe@contoso.com';
  const sourceFolderId1 = 'AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQA7DGC6yx9ARZqQFWs3P3q1AAAASBOHAAA=';
  const sourceFolderId2 = 'AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQA7DGC6yx9ARZqQFWs3P3q1AAAASBOHAAB=';
  const filterQuery = `subject eq 'Contoso'`;
  const response = {
    id: "AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQACAd7HWUedTo-i2ZIVhDiHAAoGOwIyAAA=",
    displayName: "Contoso",
    parentFolderId: "AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQA7DGC6yx9ARZqQFWs3P3q1AAAASBOLAAA=",
    childFolderCount: 0,
    unreadItemCount: 0,
    totalItemCount: 5,
    sizeInBytes: null,
    isHidden: false,
    isSupported: true,
    includeNestedFolders: false,
    sourceFolderIds: [
      sourceFolderId1
    ],
    filterQuery: filterQuery
  };
  const responseWithNestedFolders = {
    id: "AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQACAd7HWUedTo-i2ZIVhDiHAAoGOwIyAAA=",
    displayName: "Contoso",
    parentFolderId: "AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQA7DGC6yx9ARZqQFWs3P3q1AAAASBOLAAA=",
    childFolderCount: 0,
    unreadItemCount: 0,
    totalItemCount: 5,
    sizeInBytes: null,
    isHidden: false,
    isSupported: true,
    includeNestedFolders: true,
    sourceFolderIds: [
      sourceFolderId1,
      sourceFolderId2
    ],
    filterQuery: filterQuery
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    if (!auth.connection.accessTokens[auth.defaultResource]) {
      auth.connection.accessTokens[auth.defaultResource] = {
        expiresOn: 'abc',
        accessToken: 'abc'
      };
    }
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MAIL_SEARCHFOLDER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: 'foo',
      folderName: 'Contoso',
      messageFilter: filterQuery,
      sourceFoldersIds: sourceFolderId1
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid user principal name', () => {
    const actual = commandOptionsSchema.safeParse({
      userName: 'foo',
      folderName: 'Contoso',
      messageFilter: filterQuery,
      sourceFoldersIds: sourceFolderId1
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      userName: userName,
      folderName: 'Contoso',
      messageFilter: filterQuery,
      sourceFoldersIds: sourceFolderId1
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if folderName is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      messageFilter: filterQuery,
      sourceFoldersIds: sourceFolderId1
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if messageFilter is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      folderName: 'Contoso',
      sourceFoldersIds: sourceFolderId1
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('correctly creates a mail search folder in the mailbox of the signed-in user', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailFolders/searchFolders/childFolders') {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      folderName: 'Contoso',
      messageFilter: filterQuery,
      sourceFoldersIds: sourceFolderId1
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates a mail search folder in the mailbox of a user specified by id', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/mailFolders/searchFolders/childFolders`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      userId: userId,
      folderName: 'Contoso',
      messageFilter: filterQuery,
      sourceFoldersIds: sourceFolderId1
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates a mail search folder in the mailbox of a user specified by UPN', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/mailFolders/searchFolders/childFolders`) {
        return responseWithNestedFolders;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      userName: userName,
      folderName: 'Contoso',
      messageFilter: filterQuery,
      sourceFoldersIds: `${sourceFolderId1},${sourceFolderId2}`,
      includeNestedFolders: true,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledOnceWithExactly(responseWithNestedFolders));
  });

  it('fails creating a mail search folder if neither userId nor userName is specified in app-only mode', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const parsedSchema = commandOptionsSchema.safeParse({
      folderName: 'Contoso',
      messageFilter: filterQuery,
      sourceFoldersIds: `${sourceFolderId1},${sourceFolderId2}`,
      includeNestedFolders: true,
      verbose: true
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data! }), new CommandError('When running with application permissions either userId or userName is required'));
  });

  it('fails creating a mail search folder for signed-in user if userId is specified', async () => {
    const parsedSchema = commandOptionsSchema.safeParse({
      userId: userId,
      folderName: 'Contoso',
      messageFilter: filterQuery,
      sourceFoldersIds: `${sourceFolderId1},${sourceFolderId2}`,
      includeNestedFolders: true,
      verbose: true
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data! }), new CommandError('You can create mail search folder for other users only if CLI is authenticated in app-only mode'));
  });

  it('fails creating a mail search folder for signed-in user if userName is specified', async () => {
    const parsedSchema = commandOptionsSchema.safeParse({
      userName: userName,
      folderName: 'Contoso',
      messageFilter: filterQuery,
      sourceFoldersIds: `${sourceFolderId1},${sourceFolderId2}`,
      includeNestedFolders: true,
      verbose: true
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data! }), new CommandError('You can create mail search folder for other users only if CLI is authenticated in app-only mode'));
  });

  it('correctly handles error when invalid folder id is specified', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        error: {
          code: 'ErrorInvalidIdMalformed',
          message: 'Id is malformed.'
        }
      }
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      folderName: 'Contoso',
      messageFilter: filterQuery,
      sourceFoldersIds: 'foo'
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data! }), new CommandError('Id is malformed.'));
  });

  it('correctly handles error when invalid query is specified', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        error: {
          code: "ErrorParsingFilterQuery-ParseUri",
          message: "An unknown function with name 'contais' was found. This may also be a function import or a key lookup on a navigation property, which is not allowed."
        }
      }
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      folderName: 'Contoso',
      messageFilter: "contais(subject, 'CLI for Microsoft 365')",
      sourceFoldersIds: 'foo'
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data! }), new CommandError(`An unknown function with name 'contais' was found. This may also be a function import or a key lookup on a navigation property, which is not allowed.`));
  });
});