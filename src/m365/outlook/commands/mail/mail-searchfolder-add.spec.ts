import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import command from './mail-searchfolder-add.js';
import { cli } from '../../../../cli/cli.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { CommandError } from '../../../../Command.js';

describe(commands.MAIL_SEARCHFOLDER_ADD, () => {
  const userId = 'ae0e8388-cd70-427f-9503-c57498ee3337';
  const userName = 'john.doe@contoso.com';
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
      "AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQA7DGC6yx9ARZqQFWs3P3q1AAAASBOHAAA="
    ],
    filterQuery: "subject eq 'Contoso'"
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
      "AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQA7DGC6yx9ARZqQFWs3P3q1AAAASBOHAAA=",
      "AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQA7DGC6yx9ARZqQFWs3P3q1AAAASBOHAAB="
    ],
    filterQuery: "subject eq 'Contoso'"
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    assert.strictEqual(command.name, commands.MAIL_SEARCHFOLDER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: 'foo',
      folderName: 'Contoso',
      messageFilter: `subject eq 'Contoso'`,
      sourceFoldersIds: 'AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQDbzX51CNg_QrDdMlYeaLWqADbcwoABAAA='
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid user principal name', () => {
    const actual = commandOptionsSchema.safeParse({
      userName: 'foo',
      folderName: 'Contoso',
      messageFilter: `subject eq 'Contoso'`,
      sourceFoldersIds: 'AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQDbzX51CNg_QrDdMlYeaLWqADbcwoABAAA='
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      userName: userName,
      folderName: 'Contoso',
      messageFilter: `subject eq 'Contoso'`,
      sourceFoldersIds: 'AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQDbzX51CNg_QrDdMlYeaLWqADbcwoABAAA='
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if folderName is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      messageFilter: `subject eq 'Contoso'`,
      sourceFoldersIds: 'AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQDbzX51CNg_QrDdMlYeaLWqADbcwoABAAA='
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if messageFilter is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      folderName: 'Contoso',
      sourceFoldersIds: 'AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQDbzX51CNg_QrDdMlYeaLWqADbcwoABAAA='
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither userId nor userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      folderName: 'Contoso',
      messageFilter: `subject eq 'Contoso'`,
      sourceFoldersIds: 'AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQDbzX51CNg_QrDdMlYeaLWqADbcwoABAAA='
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('correctly creates a mail search folder in the mailbox of a user specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/searchFolders/childFolders`) {
        return response;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { userId: userId, folderName: 'Contoso', messageFilter: `subject eq 'Contoso'`, sourceFoldersIds: 'AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQA7DGC6yx9ARZqQFWs3P3q1AAAASBOHAAA=' } });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates a mail search folder in the mailbox of a user specified by UPN', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').resolves(userId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/searchFolders/childFolders`) {
        return responseWithNestedFolders;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { userName: userName, folderName: 'Contoso', messageFilter: `subject eq 'Contoso'`, sourceFoldersIds: 'AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQA7DGC6yx9ARZqQFWs3P3q1AAAASBOHAAA=,AAMkAGRkZTFiMDQxLWYzNDgtNGQ3ZS05Y2U3LWU1NWJhMTM5YTgwMAAuAAAAAABxI4iNfZK7SYRiWw9sza20AQA7DGC6yx9ARZqQFWs3P3q1AAAASBOHAAB=', includeNestedFodlers: true, verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly(responseWithNestedFolders));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        error: {
          code: 'ErrorInvalidIdMalformed',
          message: 'Id is malformed.'
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { userId: userId, folderName: 'Contoso', messageFilter: `subject eq 'Contoso'`, sourceFoldersIds: 'foo' } } as any), new CommandError('Id is malformed.'));
  });
});