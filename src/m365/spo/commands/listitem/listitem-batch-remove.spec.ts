import assert from 'assert';
import fs from 'fs';
import os from 'os';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './listitem-batch-remove.js';

describe(commands.LISTITEM_BATCH_REMOVE, () => {
  const filePath = 'C:\\Path\\To\\CSV\\CsvFile.csv';
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listId = 'f2978459-4e2a-4307-b57c-0c90eb4e5d6a';
  const listTitle = 'Random List';
  const listUrl = '/sites/project-x/lists/random-list';
  const csvContentHeaders = `ID`;
  const csvContentLine = `1`;
  const csvContent = `${csvContentHeaders}\n${csvContentLine}`;
  const ids = '1,2,3,4,5';

  //#region Mock Responses
  const mockBatchFailedResponse = '--batchresponse_062b14be-a782-413a-90e6-dc96852e0423\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary\r\n\r\nHTTP/1.1 404 Not Found\r\nCONTENT-TYPE: application/json;odata=verbose;charset=utf-8\r\n\r\n{"error":{"code":"-2130575338, System.ArgumentException","message":{"lang":"en-US","value":"Item does not exist. It may have been deleted by another user."}}}\r\n--batchresponse_062b14be-a782-413a-90e6-dc96852e0423--';
  const mockBatchSuccessfulResponse = '--batchresponse_125a3170-5162-45fb-ac7f-7c88d75e24c3\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary\r\n\r\nHTTP/1.1 200 OK\r\nCONTENT-TYPE: application/json;odata=nometadata;streaming=true;charset=utf-8\r\n\r\n\r\n--batchresponse_125a3170-5162-45fb-ac7f-7c88d75e24c3--';
  //#endregion

  let commandInfo: CommandInfo;
  let log: any[];
  let logger: Logger;
  let promptIssued: boolean = false;


  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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

    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      fs.existsSync,
      fs.readFileSync,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_BATCH_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing list items when confirmation argument not passed', async () => {
    await command.action(logger, { options: { webUrl: webUrl, listId: listId, ids: ids } });

    assert(promptIssued);
  });

  it('aborts removing list items when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, { options: { webUrl: webUrl, listId: listId, ids: ids } });

    assert(postSpy.notCalled);
  });

  it('removes items in batch from a sharepoint list retrieved by id when using a csv file', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(fs, 'readFileSync').returns(csvContent);
    const postStub = sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        return mockBatchSuccessfulResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listId: listId, recycle: true, verbose: true } });
    assert(postStub.calledOnce);
  });

  it('removes items in batch from a SharePoint list retrieved by id when using a csv file with different casing for the ID column', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(fs, 'readFileSync').returns(`id\n1`);
    const postStub = sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        return mockBatchSuccessfulResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listId: listId, recycle: true, verbose: true } });
    assert(postStub.calledOnce);
  });

  it('removes items from a sharepoint list retrieved by id when passing a list of ids via string', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        return mockBatchSuccessfulResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, ids: ids, listId: listId, force: true, verbose: true } });
    assert(postStub.calledOnce);
  });

  it('adds items in batch to a sharepoint list retrieved by title', async () => {
    sinon.stub(fs, 'readFileSync').returns(csvContent);
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        return mockBatchSuccessfulResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listTitle: listTitle, force: true, verbose: true } });
  });

  it('adds items in batch to a sharepoint list retrieved by title and removes empty items and final empty line if they exist', async () => {
    sinon.stub(fs, 'readFileSync').returns(`${csvContent}\n\n3\n`);
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        return mockBatchSuccessfulResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listTitle: listTitle, force: true, verbose: true } });
  });

  it('removes 150 items in batch from a sharepoint list retrieved by url and reading the content from a csv file and confirming the prompt', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    let csvContent150Items = csvContent;
    for (let i = 1; i < 150; i++) {
      csvContent150Items += `\n${csvContentLine}`;
    }
    let amountOfRequestsInBody = 0;
    sinon.stub(fs, 'readFileSync').returns(csvContent150Items);
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        amountOfRequestsInBody += opts.data.match(/DELETE/g).length;
        return mockBatchSuccessfulResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listUrl: listUrl, recycle: true, verbose: true } });
    assert.strictEqual(amountOfRequestsInBody, 150);
  });

  it('throws an error when batch api URL fails', async () => {
    const errorMessage = 'SharePoint REST Service Exception';
    sinon.stub(fs, 'readFileSync').returns(csvContent);
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        throw errorMessage;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listUrl: listUrl, force: true, verbose: true } }), new CommandError(errorMessage));
  });

  it('throws an error when batch api returns partly unsuccessful results', async () => {
    const errorMessage = `Creating some items failed with the following errors: ${os.EOL}- Item ID 1: Item does not exist. It may have been deleted by another user.`;
    sinon.stub(fs, 'readFileSync').returns(csvContent);
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        return mockBatchFailedResponse;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listUrl: listUrl, force: true, verbose: true } }), new CommandError(errorMessage));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', filePath: filePath, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId option is not a valid GUID', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = await command.validate({ options: { webUrl: webUrl, filePath: filePath, listId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if csv file does not exist', async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = await command.validate({ options: { webUrl: webUrl, filePath: filePath, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if filePath exists and listId is a valid guid, but the csvContent does not contain the ID column', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(`${csvContent}\n2\ninvalid`);
    const actual = await command.validate({ options: { webUrl: webUrl, filePath: filePath, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if filePath exists and listId is a valid guid, but the csvContent contains invalid ids', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(`${csvContentHeaders}invalid\n${csvContentLine}`);
    const actual = await command.validate({ options: { webUrl: webUrl, filePath: filePath, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if filePath exists, csv content is valid and listId is a valid guid', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(csvContent);
    const actual = await command.validate({ options: { webUrl: webUrl, filePath: filePath, listId: listId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if ids are passed and ids contains invalid numbers', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, ids: `${ids},invalid` } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if ids are passed and ids contains only valid numbers', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, ids: ids } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
