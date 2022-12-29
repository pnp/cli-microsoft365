import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./file-version-list');

describe(commands.FILE_VERSION_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const validWebUrl = "https://contoso.sharepoint.com";
  const validFileUrl = "/Shared Documents/Document.docx";
  const validFileId = "7a9b8bb6-d5c4-4de9-ab76-5210a7879e89";
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
      },
      {
        "CheckInComment": "",
        "Created": "2022-10-30T12:06:13Z",
        "ID": 1024,
        "IsCurrentVersion": false,
        "Length": "21098",
        "Size": 21098,
        "Url": "_vti_history/1024/Shared Documents/Document.docx",
        "VersionLabel": "2.0"
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
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
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_VERSION_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Created', 'ID', 'IsCurrentVersion', 'VersionLabel']);
  });

  it('fails validation if fileId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        webUrl: validWebUrl,
        fileId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileUrl: validFileUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (fileUrl)', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, fileUrl: validFileUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (fileId)', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, fileId: validFileId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves versions from a file with the fileUrl option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(validFileUrl)}')/versions`) {
        return fileVersionResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        fileUrl: validFileUrl
      }
    });
    assert(loggerLogSpy.calledWith(fileVersionResponse.value));
  });

  it('retrieves versions from a file with the fileId option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileById('${validFileId}')/versions`) {
        return fileVersionResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        fileId: validFileId
      }
    });
    assert(loggerLogSpy.calledWith(fileVersionResponse.value));
  });

  it('handles a random API error correctly', async () => {
    const err = 'Invalid versions request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileById('${validFileId}')/versions`) {
        throw { error: { 'odata.error': { message: { value: err } } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        fileId: validFileId
      }
    }), new CommandError(err));
  });
});
