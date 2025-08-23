import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { z } from 'zod';
import commands from '../../commands.js';
import command from './file-version-keep.js';

describe(commands.FILE_VERSION_KEEP, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  const validWebUrl = "https://contoso.sharepoint.com";
  const validFileUrl = "/Shared Documents/Document.docx";
  const validFileId = "7a9b8bb6-d5c4-4de9-ab76-5210a7879e89";
  const validLabel = "1.0";

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    auth.connection.active = true;
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_VERSION_KEEP);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webUrl is not a valid URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', label: validLabel, fileUrl: validFileUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if fileId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: validWebUrl, fileId: 'invalid', label: validLabel });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if fileUrl and fileId are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: validWebUrl, fileUrl: validFileUrl, fileId: validFileId, label: validLabel });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if label is not specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: validWebUrl, fileUrl: validFileUrl });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if fileUrl is specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: validWebUrl, fileUrl: validFileUrl, label: validLabel });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if fileId is specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: validWebUrl, fileId: validFileId, label: validLabel });
    assert.strictEqual(actual.success, true);
  });

  it('ensures that a specific file version will never expire (fileUrl)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(validFileUrl)}')/versions?$filter=VersionLabel eq '${validLabel}'&$select=ID`) {
        return {
          value: [
            {
              ID: 1
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(validFileUrl)}')/versions(1)/SetExpirationDate()`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: validWebUrl, fileUrl: validFileUrl, label: validLabel, verbose: true } });
    assert(postStub.calledOnce);
  });

  it('ensures that a specific file version will never expire (fileId)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileById('${validFileId}')/versions?$filter=VersionLabel eq '${validLabel}'&$select=ID`) {
        return {
          value: [
            {
              ID: 1
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileById('${validFileId}')/versions(1)/SetExpirationDate()`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: validWebUrl, fileId: validFileId, label: validLabel, verbose: true } });
    assert(postStub.calledOnce);
  });

  it('correctly handles error when the specified version does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(validFileUrl)}')/versions?$filter=VersionLabel eq '${validLabel}'&$select=ID`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { webUrl: validWebUrl, fileUrl: validFileUrl, label: validLabel } }),
      new CommandError(`Version with label '${validLabel}' not found.`));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'Invalid version request'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { webUrl: validWebUrl, fileId: validFileId, label: validLabel } }),
      new CommandError('Invalid version request'));
  });
});