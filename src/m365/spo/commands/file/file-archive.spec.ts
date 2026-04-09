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
import command from './file-archive.js';

describe(commands.FILE_ARCHIVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let confirmationPromptStub: sinon.SinonStub;
  let loggerLogSpy: sinon.SinonSpy;

  const successResponse = {
    value: "fullyArchived"
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
    assert.strictEqual(command.name, commands.FILE_ARCHIVE);
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
      url: '/sites/test/Shared documents/document.docx',
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
      url: '/sites/test/Shared documents/document.docx',
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

  it('prompts before archiving file when confirmation argument not passed', async () => {
    sinon.stub(request, 'get').resolves({ ListId: 'b2307a39-e878-458b-bc90-03bc578531d6', ListItemAllFields: { Id: 1 } });
    sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        id: '00000000-0000-0000-0000-000000000000'
      }
    });
    assert(confirmationPromptStub.calledOnce);
  });

  it('aborts archiving file when prompt not confirmed', async () => {
    const getStub = sinon.stub(request, 'get').resolves({ ListId: 'b2307a39-e878-458b-bc90-03bc578531d6', ListItemAllFields: { Id: 1 } });
    const postStub = sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        url: '/sites/test/Shared documents/document.docx'
      }
    });

    assert(getStub.notCalled);
    assert(postStub.notCalled);
  });

  it('archives file by url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/test/Shared documents/document.docx')}')?$select=ListId,ListItemAllFields/Id&$expand=ListItemAllFields`) {
        return {
          ListId: 'b2307a39-e878-458b-bc90-03bc578531d6',
          ListItemAllFields: {
            Id: 1
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
        url: '/sites/test/Shared documents/document.docx',
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('archives file by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/GetFileById('${formatting.encodeQueryParameter('00000000-0000-0000-0000-000000000000')}')?$select=ListId,ListItemAllFields/Id&$expand=ListItemAllFields`) {
        return {
          ListId: 'b2307a39-e878-458b-bc90-03bc578531d6',
          ListItemAllFields: {
            Id: 1
          }
        };
      }

      throw 'Invalid request';
    }
    );

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

  it('archives file using site-relative url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/test/Shared Documents/document.docx')}')?$select=ListId,ListItemAllFields/Id&$expand=ListItemAllFields`) {
        return {
          ListId: 'b2307a39-e878-458b-bc90-03bc578531d6',
          ListItemAllFields: {
            Id: 1
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
        url: '/Shared Documents/document.docx',
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('outputs no result when archiving a file', async () => {
    sinon.stub(request, 'get').resolves({ ListId: 'b2307a39-e878-458b-bc90-03bc578531d6', ListItemAllFields: { Id: 1 } });
    sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        url: '/sites/test/Shared documents/document.docx',
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
            value: 'The file /sites/test/Shared documents/document.docx does not exist.'
          }
        }
      }
    };

    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        url: '/sites/test/Shared documents/document.docx',
        force: true
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });
});
