import assert from 'assert';
import sinon from 'sinon';
import { telemetry } from '../../../../telemetry.js';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import fs from 'fs';
import request from '../../../../request.js';
import command, { options } from './listitem-attachment-add.js';

describe(commands.LISTITEM_ATTACHMENT_ADD, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listId = '236a0f92-482d-475b-ba8f-d0e4f78555e4';
  const listTitle = 'Test list';
  const listUrl = 'sites/project-x/lists/testlist';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
  const listItemId = 1;
  const filePath = 'C:\\Temp\\Test.pdf';
  const fileName = 'CLIRocks.pdf';

  const response = { 'FileName': 'CLIRocks.pdf', 'FileNameAsPath': { 'DecodedUrl': 'Testje.pdf' }, 'ServerRelativePath': { 'DecodedUrl': '/Lists/aaaaaa/Attachments/743/Testje.pdf' }, 'ServerRelativeUrl': '/Lists/aaaaaa/Attachments/743/Testje.pdf' };

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((_, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_ATTACHMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid', listTitle: listTitle, listItemId: listItemId, filePath: filePath });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'invalid' is not a valid SharePoint Online site URL.`);
  });
  it('passes validation if the webUrl option is a valid SharePoint site URL and filePath exists', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, filePath: filePath });
    assert.strictEqual(actual.success, true);
  });
  it('fails validation if the listId option is not a valid GUID', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: 'invalid', listItemId: listItemId, filePath: filePath });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'invalid in option listId is not a valid GUID.');
  });
  it('fails validation if the listItemId option is not a valid number', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: 'invalid', filePath: filePath });
    assert.strictEqual(actual.success, false);
  });
  it('passes validation if the listId option is a valid GUID', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: listItemId, filePath: filePath });
    assert.strictEqual(actual.success, true);
  });
  it('fails validation if filePath does not exist', () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, filePath: filePath });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `File with path '${filePath}' was not found.`);
  });

  it('fails validation with unknown options', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, filePath: filePath, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('adds attachment to listitem in list retrieved by id while specifying fileName', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('content read');
    sinon.stub(request, 'post').callsFake(async (args) => {
      if (args.url === `${webUrl}/_api/web/lists(guid'${listId}')/items(${listItemId})/AttachmentFiles/add(FileName='${fileName}')`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, webUrl: webUrl, listId: listId, listItemId: listItemId, filePath: filePath, fileName: fileName }) });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('adds attachment to listitem in list retrieved by url while not specifying fileName', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('content read');
    sinon.stub(request, 'post').callsFake(async (args) => {
      if (args.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(${listItemId})/AttachmentFiles/add(FileName='${filePath.replace(/^.*[\\/]/, '')}')`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, webUrl: webUrl, listUrl: listUrl, listItemId: listItemId, filePath: filePath }) });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('adds attachment to listitem in list retrieved by url while specifying fileName without extension', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('content read');
    const fileNameWithoutExtension = fileName.split('.')[0];
    const fileNameWithExtension = `${fileNameWithoutExtension}.${filePath.split('.').pop()}`;
    sinon.stub(request, 'post').callsFake(async (args) => {
      if (args.url === `${webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')/items(${listItemId})/AttachmentFiles/add(FileName='${fileNameWithExtension}')`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, filePath: filePath, fileName: fileNameWithoutExtension }) });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('handles error when file with specific name already exists', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('content read');
    const error = {
      error: {
        'odata.error': {
          code: '-2130575257, Microsoft.SharePoint.SPException',
          message: {
            lang: 'en-US',
            value: 'The specified name is already in use.\n\nThe document or folder name was not changed.  To change the name to a different value, close this dialog and edit the properties of the document or folder.'
          }
        }
      }
    };
    sinon.stub(request, 'post').callsFake(async (args) => {
      if (args.url === `${webUrl}/_api/web/lists(guid'${listId}')/items(${listItemId})/AttachmentFiles/add(FileName='${fileName}')`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, webUrl: webUrl, listId: listId, listItemId: listItemId, filePath: filePath, fileName: fileName }) }),
      new CommandError(error.error['odata.error'].message.value.split('\n')[0]));
  });

  it('handles API error', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('content read');
    const error = {
      error: {
        'odata.error': {
          code: '-2130575257, Microsoft.SharePoint.SPException',
          message: {
            lang: 'en-US',
            value: 'An error has occured.'
          }
        }
      }
    };
    sinon.stub(request, 'post').callsFake(async (args) => {
      if (args.url === `${webUrl}/_api/web/lists(guid'${listId}')/items(${listItemId})/AttachmentFiles/add(FileName='${fileName}')`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, webUrl: webUrl, listId: listId, listItemId: listItemId, filePath: filePath, fileName: fileName }) }),
      new CommandError(error.error['odata.error'].message.value));
  });
});
