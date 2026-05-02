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
import command from './folder-unarchive.js';

describe(commands.FOLDER_UNARCHIVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let confirmationPromptStub: sinon.SinonStub;
  let loggerLogSpy: sinon.SinonSpy;

  const folderInfoResponse = {
    ListItemAllFields: {
      Id: 1,
      ParentList: {
        Id: 'b2307a39-e878-458b-bc90-03bc578531d6'
      }
    }
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
    assert.strictEqual(command.name, commands.FOLDER_UNARCHIVE);
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
      url: '/sites/Marketing/documents/general',
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
      url: '/sites/Marketing/documents/general',
      force: true
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with valid options (id)', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      id: '7a8c9207-7745-4cda-b0e2-be2618ee3030',
      force: true
    });
    assert.strictEqual(actual.success, true);
  });

  it('prompts before unarchiving folder when confirmation argument not passed', async () => {
    sinon.stub(request, 'get').resolves(folderInfoResponse);
    sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        id: '7a8c9207-7745-4cda-b0e2-be2618ee3030'
      }
    });
    assert(confirmationPromptStub.calledOnce);
  });

  it('aborts unarchiving folder when prompt not confirmed', async () => {
    const getStub = sinon.stub(request, 'get').resolves(folderInfoResponse);
    const postStub = sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        url: '/sites/Marketing/documents/general'
      }
    });

    assert(getStub.notCalled);
    assert(postStub.notCalled);
  });

  it('unarchives folder by url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Marketing/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/Marketing/documents/general')}')?$select=ListItemAllFields/ParentList/Id,ListItemAllFields/Id&$expand=ListItemAllFields/ParentList`) {
        return folderInfoResponse;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Marketing/_api/Lists(guid'b2307a39-e878-458b-bc90-03bc578531d6')/items(1)/UnArchive`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/Marketing',
        url: '/sites/Marketing/documents/general',
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('unarchives folder by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Marketing/_api/web/GetFolderById('${formatting.encodeQueryParameter('7a8c9207-7745-4cda-b0e2-be2618ee3030')}')?$select=ListItemAllFields/ParentList/Id,ListItemAllFields/Id&$expand=ListItemAllFields/ParentList`) {
        return folderInfoResponse;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Marketing/_api/Lists(guid'b2307a39-e878-458b-bc90-03bc578531d6')/items(1)/UnArchive`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/Marketing',
        id: '7a8c9207-7745-4cda-b0e2-be2618ee3030',
        verbose: true,
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('unarchives folder using site-relative url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Marketing/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/Marketing/Shared Documents/general')}')?$select=ListItemAllFields/ParentList/Id,ListItemAllFields/Id&$expand=ListItemAllFields/ParentList`) {
        return folderInfoResponse;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Marketing/_api/Lists(guid'b2307a39-e878-458b-bc90-03bc578531d6')/items(1)/UnArchive`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/Marketing',
        url: '/Shared Documents/general',
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('outputs no result when unarchiving a folder', async () => {
    sinon.stub(request, 'get').resolves(folderInfoResponse);
    sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/Marketing',
        url: '/sites/Marketing/documents/general',
        force: true
      }
    });

    assert(loggerLogSpy.notCalled);
  });

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          code: "-2130575338, Microsoft.SharePoint.SPException",
          message: {
            lang: "en-US",
            value: 'The folder /sites/Marketing/documents/general does not exist.'
          }
        }
      }
    };

    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/Marketing',
        url: '/sites/Marketing/documents/general',
        force: true
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });
});
