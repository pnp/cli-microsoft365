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
import { formatting } from '../../../../utils/formatting.js';
import { z } from 'zod';
import commands from '../../commands.js';
import command from './folder-archive.js';

describe(commands.FOLDER_ARCHIVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let confirmationPromptStub: sinon.SinonStub;
  let loggerLogSpy: sinon.SinonSpy;

  const successResponse = {
    value: '{"IsArchive":true,"TotalFileCount":1,"CreatedUtcDateTime":"2026-04-30T16:34:57.3834786Z","LastStartedUtcDateTime":"0001-01-01T00:00:00","FolderArchiveStatus":"Unknown","ProcessedFileCount":0,"SuccessCount":0,"FailureCount":0,"NotArchivableFileCount":0,"ProgressPercentage":0.0}'
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
    confirmationPromptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_ARCHIVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['url']);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'invalid-url',
      id: '00000000-0000-0000-0000-000000000000',
      force: true
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both url and id are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      url: '/sites/test/Shared documents/folder',
      id: '00000000-0000-0000-0000-000000000000',
      force: true
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither url nor id are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      force: true
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      id: 'invalid-guid',
      force: true
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with valid options (url)', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      url: '/sites/test/Shared documents/folder',
      force: true
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with valid options (id)', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      id: '00000000-0000-0000-0000-000000000000',
      force: true
    });
    assert.strictEqual(actual.success, true);
  });

  it('prompts before archiving folder when confirmation argument not passed', async () => {
    sinon.stub(request, 'get').resolves({ Exists: true, ListItemAllFields: { Id: 1, ParentList: { Id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } });
    sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        id: '00000000-0000-0000-0000-000000000000'
      }
    });
    assert(confirmationPromptStub.calledOnce);
  });

  it('aborts archiving folder when prompt not confirmed', async () => {
    const getStub = sinon.stub(request, 'get').resolves({ Exists: true, ListItemAllFields: { Id: 1, ParentList: { Id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } });
    const postStub = sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        url: '/sites/test/Shared documents/folder'
      }
    });

    assert(getStub.notCalled);
    assert(postStub.notCalled);
  });

  it('archives folder by url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/test/Shared documents/folder')}')?$select=Exists,ListItemAllFields/Id,ListItemAllFields/ParentList/Id&$expand=ListItemAllFields,ListItemAllFields/ParentList`) {
        return {
          Exists: true,
          ListItemAllFields: {
            Id: 1,
            ParentList: {
              Id: 'b2307a39-e878-458b-bc90-03bc578531d6'
            }
          }
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/Lists(guid'b2307a39-e878-458b-bc90-03bc578531d6')/items(1)/Archive`) {
        return successResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        url: '/sites/test/Shared documents/folder',
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('archives folder by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/GetFolderById('${formatting.encodeQueryParameter('00000000-0000-0000-0000-000000000000')}')?$select=Exists,ListItemAllFields/Id,ListItemAllFields/ParentList/Id&$expand=ListItemAllFields,ListItemAllFields/ParentList`) {
        return {
          Exists: true,
          ListItemAllFields: {
            Id: 1,
            ParentList: {
              Id: 'b2307a39-e878-458b-bc90-03bc578531d6'
            }
          }
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/Lists(guid'b2307a39-e878-458b-bc90-03bc578531d6')/items(1)/Archive`) {
        return successResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        id: '00000000-0000-0000-0000-000000000000',
        verbose: true,
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('archives folder using site-relative url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/test/Shared Documents/folder')}')?$select=Exists,ListItemAllFields/Id,ListItemAllFields/ParentList/Id&$expand=ListItemAllFields,ListItemAllFields/ParentList`) {
        return {
          Exists: true,
          ListItemAllFields: {
            Id: 1,
            ParentList: {
              Id: 'b2307a39-e878-458b-bc90-03bc578531d6'
            }
          }
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/Lists(guid'b2307a39-e878-458b-bc90-03bc578531d6')/items(1)/Archive`) {
        return successResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        url: '/Shared Documents/folder',
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('outputs no result when archiving a folder', async () => {
    sinon.stub(request, 'get').resolves({ Exists: true, ListItemAllFields: { Id: 1, ParentList: { Id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } });
    sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        url: '/sites/test/Shared documents/folder',
        force: true
      }
    });

    assert(loggerLogSpy.notCalled);
  });

  it('throws an error when trying to archive the root folder of a document library by url', async () => {
    sinon.stub(request, 'get').resolves({ Exists: true, ListItemAllFields: {} });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        url: '/Shared Documents',
        force: true
      }
    }), new CommandError(`The folder '/Shared Documents' is the root folder of a document library and cannot be archived. Archive a subfolder instead.`));
  });

  it('throws an error when trying to archive the root folder of a document library by id', async () => {
    sinon.stub(request, 'get').resolves({ Exists: true, ListItemAllFields: {} });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        id: '00000000-0000-0000-0000-000000000000',
        force: true
      }
    }), new CommandError(`The folder '00000000-0000-0000-0000-000000000000' is the root folder of a document library and cannot be archived. Archive a subfolder instead.`));
  });

  it('throws an error when the folder does not exist by url', async () => {
    sinon.stub(request, 'get').resolves({});

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        url: '/Shared Documents/temp1',
        force: true
      }
    }), new CommandError(`The folder '/Shared Documents/temp1' does not exist.`));
  });

  it('throws an error when the folder does not exist by id', async () => {
    sinon.stub(request, 'get').resolves({});

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        id: '00000000-0000-0000-0000-000000000000',
        force: true
      }
    }), new CommandError(`The folder '00000000-0000-0000-0000-000000000000' does not exist.`));
  });

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          code: "-2130575338, Microsoft.SharePoint.SPException",
          message: {
            lang: "en-US",
            value: 'The folder /sites/test/Shared documents/folder does not exist.'
          }
        }
      }
    };

    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        url: '/sites/test/Shared documents/folder',
        force: true
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });
});
