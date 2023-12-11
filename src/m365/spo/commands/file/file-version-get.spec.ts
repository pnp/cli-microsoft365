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
import commands from '../../commands.js';
import command from './file-version-get.js';

describe(commands.FILE_VERSION_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const validWebUrl = "https://contoso.sharepoint.com";
  const validFileUrl = "/Shared Documents/Document.docx";
  const validFileId = "7a9b8bb6-d5c4-4de9-ab76-5210a7879e89";
  const validLabel = "1.0";
  const fileVersionResponse = {
    "value": [
      {
        "CheckInComment": "",
        "Created": "2022-10-30T12:03:06Z",
        "ID": 512,
        "IsCurrentVersion": false,
        "Length": "18898",
        "Size": 18898,
        "Url": "_vti_history/512/Shared Documents/Document.docx",
        "VersionLabel": "1.0"
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_VERSION_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if fileId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        webUrl: validWebUrl,
        label: validLabel,
        fileId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', label: validLabel, fileUrl: validFileUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (fileUrl)', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, label: validLabel, fileUrl: validFileUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (fileId)', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, label: validLabel, fileId: validFileId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves version from a file with the fileUrl options', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(validFileUrl)}')/versions/?$filter=VersionLabel eq '${validLabel}'`) {
        return fileVersionResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        label: validLabel,
        fileUrl: validFileUrl
      }
    });
    assert(loggerLogSpy.calledWith(fileVersionResponse.value[0]));
  });

  it('retrieves version from a file with the fileId options', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileById('${validFileId}')/versions/?$filter=VersionLabel eq '${validLabel}'`) {
        return fileVersionResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        label: validLabel,
        fileId: validFileId
      }
    });
    assert(loggerLogSpy.calledWith(fileVersionResponse.value[0]));
  });

  it('properly escapes single quotes in fileUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url = `${validWebUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='Shared%20Documents%2FFo''lde''r')/versions/?$filter=VersionLabel eq '${validLabel}'`) {
        return fileVersionResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        label: validLabel,
        fileUrl: `Shared Documents/Fo'lde'r`
      }
    });
    assert(loggerLogSpy.calledWith(fileVersionResponse.value[0]));
  });

  it('command correctly handles version list reject request', async () => {
    const err = 'Invalid version request';
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: err
          }
        }
      }
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        label: validLabel
      }
    }), new CommandError(err));
  });
});
