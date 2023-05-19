import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./file-rename');

describe(commands.FILE_RENAME, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const renameResponseJson = [
    {
      'ErrorCode': 0,
      'ErrorMessage': null,
      'FieldName': 'FileLeafRef',
      'FieldValue': 'test 2.docx',
      'HasException': false,
      'ItemId': 642
    }
  ];

  const renameValue = {
    value: renameResponseJson
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      Cli.executeCommand
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_RENAME), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', sourceUrl: 'abc', targetFileName: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', sourceUrl: 'abc', targetFileName: 'abc' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('forcefully renames file from a non-root site in the root folder of a document library when a file with the same name exists (or it doesn\'t?)', async () => {
    sinon.stub(Cli, 'executeCommand');

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string) === 'https://contoso.sharepoint.com/sites/portal/_api/web/GetFileByServerRelativeUrl(\'%2Fsites%2Fportal%2FShared%20Documents%2Fabc.pdf\')/ListItemAllFields/ValidateUpdateListItem()') {
        return Promise.resolve(renameValue);
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string) === 'https://contoso.sharepoint.com/sites/portal/_api/web/GetFileByServerRelativeUrl(\'%2Fsites%2Fportal%2FShared%20Documents%2Fabc.pdf\')?$select=UniqueId') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        sourceUrl: '/Shared Documents/abc.pdf',
        force: true,
        targetFileName: 'def.pdf'
      }
    });
    assert(loggerLogSpy.calledWith(renameResponseJson));
  });

  it('renames file from a non-root site in the root folder of a document library when a file with the same name doesn\'t exist', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string) === 'https://contoso.sharepoint.com/sites/portal/_api/web/GetFileByServerRelativeUrl(\'%2Fsites%2Fportal%2FShared%20Documents%2Fabc.pdf\')/ListItemAllFields/ValidateUpdateListItem()') {
        return Promise.resolve(renameValue);
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string) === 'https://contoso.sharepoint.com/sites/portal/_api/web/GetFileByServerRelativeUrl(\'%2Fsites%2Fportal%2FShared%20Documents%2Fabc.pdf\')?$select=UniqueId') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        sourceUrl: 'Shared Documents/abc.pdf',
        targetFileName: 'def.pdf'
      }
    });
    assert(loggerLogSpy.calledWith(renameResponseJson));
  });

  it('continues if file cannot be recycled because it does not exist', async () => {
    const fileDeleteError = {
      error: {
        message: 'File does not exist'
      }
    };
    sinon.stub(Cli, 'executeCommand').returns(Promise.reject(fileDeleteError));

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string) === 'https://contoso.sharepoint.com/sites/portal/_api/web/GetFileByServerRelativeUrl(\'%2Fsites%2Fportal%2FShared%20Documents%2Fabc.pdf\')/ListItemAllFields/ValidateUpdateListItem()') {
        return Promise.resolve(renameValue);
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string) === 'https://contoso.sharepoint.com/sites/portal/_api/web/GetFileByServerRelativeUrl(\'%2Fsites%2Fportal%2FShared%20Documents%2Fabc.pdf\')?$select=UniqueId') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        sourceUrl: 'Shared Documents/abc.pdf',
        force: true,
        targetFileName: 'def.pdf'
      }
    });
    assert(loggerLogSpy.calledWith(renameResponseJson));
  });

  it('throws error if file cannot be recycled', async () => {
    const fileDeleteError = {
      error: {
        message: 'Locked for use'
      },
      stderr: ''
    };

    sinon.stub(Cli, 'executeCommand').returns(Promise.reject(fileDeleteError));

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string) === `https://contoso.sharepoint.com/sites/portal/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fportal%2FShared%20Documents%2Fabc.pdf')?$select=UniqueId`) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        sourceUrl: 'Shared Documents/abc.pdf',
        force: true,
        targetFileName: 'def.pdf'
      }
    }), new CommandError(fileDeleteError.error.message));
  });
});
