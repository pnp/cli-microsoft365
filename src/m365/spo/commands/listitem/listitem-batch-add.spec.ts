import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import request from '../../../../request';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import Command, { CommandError } from '../../../../Command';
import commands from '../../commands';
import { Logger } from '../../../../cli/Logger';
const command: Command = require('./listitem-batch-add');

describe(commands.LISTITEM_BATCH_ADD, () => {
  const filePath = 'C:\\Path\\To\\CSV\\CsvFile.csv';
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listId = 'f2978459-4e2a-4307-b57c-0c90eb4e5d6a';
  const listTitle = 'Random List';
  const listUrl = '/sites/project-x/lists/random-list';
  const csvContentHeaders = `ContentType,Title,SingleChoiceField,MultiChoiceField,SingleMetadataField,MultiMetadataField,SinglePeopleField,MultiPeopleField,CustomHyperlink,NumberField`;
  const csvContentLine = `Item,Title A,Choice 1,Choice 1;#Choice 2,Engineering|4a3cc5f3-a4a6-433e-a07a-746978ff1760;,Engineering|4a3cc5f3-a4a6-433e-a07a-746978ff1760;Finance|f994a4ac-cf34-448e-a22c-2b35fd9bbffa;,[{'Key':'i:0#.f|membership|markh@contoso.com'}],"[{'Key':'i:0#.f|membership|markh@contoso.com'},{'Key':'i:0#.f|membership|adamb@contoso.com'}]","https://bing.com, URL",5`;
  const csvContent = `${csvContentHeaders}\n${csvContentLine}`;

  let commandInfo: CommandInfo;
  let log: any[];
  let logger: Logger;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_BATCH_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds items in batch to a sharepoint list retrieved by id', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent);
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listId: listId, verbose: true } } as any);
  });

  it('adds items in batch to a sharepoint list retrieved by title', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent);
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listTitle: listTitle, verbose: true } } as any);
  });

  it('adds 150 items in batch to a sharepoint list retrieved by url', async () => {
    let csvContent150Items = csvContent;
    for (let i = 1; i < 150; i++) {
      csvContent150Items += `\n${csvContentLine}`;
    }
    let amountOfRequestsInBody = 0;
    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent150Items);
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        amountOfRequestsInBody += opts.data.match(/POST/g).length;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listUrl: listUrl, verbose: true } } as any);
    assert.strictEqual(amountOfRequestsInBody, 150);
  });

  it('throws an error when batch api URL fails', async () => {
    const errorMessage = 'SharePoint REST Service Exception';
    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent);
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/$batch`) {
        throw errorMessage;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listUrl: listUrl, verbose: true } } as any), new CommandError(errorMessage));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', filePath: filePath, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId option is not a valid GUID', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    const actual = await command.validate({ options: { webUrl: webUrl, filePath: filePath, listId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if csv file does not exist', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    const actual = await command.validate({ options: { webUrl: webUrl, filePath: filePath, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if filePath exists and listId is a valid guid', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    const actual = await command.validate({ options: { webUrl: webUrl, filePath: filePath, listId: listId } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
