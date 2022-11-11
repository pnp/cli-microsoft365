import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
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
const command: Command = require('./file-version-get');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_VERSION_GET), true);
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
      if (opts.url === `${validWebUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(validFileUrl)}')/versions/?$filter=VersionLabel eq '${validLabel}'`) {
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
      if (opts.url = `${validWebUrl}/_api/web/GetFileByServerRelativeUrl('Shared%20Documents%2FFo''lde''r')/versions/?$filter=VersionLabel eq '${validLabel}'`) {
        return fileVersionResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: false,
        webUrl: validWebUrl,
        label: validLabel,
        fileUrl: `Shared Documents/Fo'lde'r`
      }
    });
    assert(loggerLogSpy.calledWith(fileVersionResponse.value[0]));
  });

  it('command correctly handles version list reject request', async () => {
    const err = 'Invalid version request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        throw err;
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

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});