import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './autofillcolumn-set.js';
import { z } from 'zod';
import { CommandError } from '../../../../Command.js';

describe(commands.AUTOFILLCOLUMN_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    assert.strictEqual(command.name, commands.AUTOFILLCOLUMN_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when required parameters are valid with column id and list id', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', prompt: 'test' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when required parameters are valid with column title and list id', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnTitle: 'ColumnName', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', prompt: 'test' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when required parameters are valid with column id and list title', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listTitle: 'Documents', prompt: 'test' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when required parameters are valid with column id and list URL', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listUrl: '/Shared Documents', prompt: 'test' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when siteUrl is not valid', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'invalidUrl', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', prompt: 'test' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when column id is not valid', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: 'invalidId', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', prompt: 'test' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when list id is not valid', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: 'invalidId', prompt: 'test' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when both columnId and columnTitle are provided', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', columnTitle: "DoubledColumn", listUrl: '/Shared Documents', prompt: 'test' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when both listTitle and ListUrl are provided', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listUrl: '/Shared Documents', listTitle: "Documents", prompt: 'test' });
    assert.notStrictEqual(actual.success, true);
  });

  it('apply autofill to column by id and list id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')/fields/getbyid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')?&$select=Id,Title,FieldTypeKind,AutofillInfo`) {
        return {
          Id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          Title: 'ColumnName',
          FieldTypeKind: 1,
          AutofillInfo: null
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=Id,BaseType`) {
        return {
          Id: "421b1e42-794b-4c71-93ac-5ed92488b67d",
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/machinelearning/SetSyntexPoweredColumnPrompts`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', prompt: 'test', verbose: true }) });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      docLibId: `{421b1e42-794b-4c71-93ac-5ed92488b67d}`,
      syntexPoweredColumnPrompts: JSON.stringify([{ columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', columnName: 'ColumnName', prompt: 'test', isEnabled: true }])
    });
  });

  it('apply autofill to column by id and list title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')/fields/getbyid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')?&$select=Id,Title,FieldTypeKind,AutofillInfo`) {
        return {
          Id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          Title: 'ColumnName',
          FieldTypeKind: 1,
          AutofillInfo: null
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists/getByTitle('Documents')?$select=Id,BaseType`) {
        return {
          Id: "421b1e42-794b-4c71-93ac-5ed92488b67d",
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/machinelearning/SetSyntexPoweredColumnPrompts`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listTitle: 'Documents', prompt: 'test' }) });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      docLibId: `{421b1e42-794b-4c71-93ac-5ed92488b67d}`,
      syntexPoweredColumnPrompts: JSON.stringify([{ columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', columnName: 'ColumnName', prompt: 'test', isEnabled: true }])
    });
  });

  it('apply autofill to column by title and list id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url ===
        `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')/fields/getbyinternalnameortitle('ColumnName')?&$select=Id,Title,FieldTypeKind,AutofillInfo`) {
        return {
          Id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          Title: 'ColumnName',
          FieldTypeKind: 1,
          AutofillInfo: null
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=Id,BaseType`) {
        return {
          Id: "421b1e42-794b-4c71-93ac-5ed92488b67d",
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/machinelearning/SetSyntexPoweredColumnPrompts`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnTitle: 'ColumnName', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', prompt: 'test', verbose: true }) });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      docLibId: `{421b1e42-794b-4c71-93ac-5ed92488b67d}`,
      syntexPoweredColumnPrompts: JSON.stringify([{ columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', columnName: 'ColumnName', prompt: 'test', isEnabled: true }])
    });
  });

  it('apply autofill to column by internalName and list id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url ===
        `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')/fields/getbyinternalnameortitle('ColumnInternalName')?&$select=Id,Title,FieldTypeKind,AutofillInfo`) {
        return {
          Id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          Title: 'ColumnInternalName',
          FieldTypeKind: 1,
          AutofillInfo: null
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=Id,BaseType`) {
        return {
          Id: "421b1e42-794b-4c71-93ac-5ed92488b67d",
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/machinelearning/SetSyntexPoweredColumnPrompts`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnInternalName: 'ColumnInternalName', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', prompt: 'test', verbose: true }) });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      docLibId: `{421b1e42-794b-4c71-93ac-5ed92488b67d}`,
      syntexPoweredColumnPrompts: JSON.stringify([{ columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', columnName: 'ColumnInternalName', prompt: 'test', isEnabled: true }])
    });
  });

  it('apply autofill to column by id and list url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')/fields/getbyinternalnameortitle('ColumnName')?&$select=Id,Title,FieldTypeKind,AutofillInfo`) {
        return {
          Id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          Title: 'ColumnName',
          FieldTypeKind: 1,
          AutofillInfo: null
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('%2Fsites%2Fsales%2FShared%20Documents')?$select=Id,BaseType`) {
        return {
          Id: "421b1e42-794b-4c71-93ac-5ed92488b67d",
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/machinelearning/SetSyntexPoweredColumnPrompts`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnTitle: 'ColumnName', prompt: 'test', listUrl: '/Shared Documents', isEnabled: false }) });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      docLibId: `{421b1e42-794b-4c71-93ac-5ed92488b67d}`,
      syntexPoweredColumnPrompts: JSON.stringify([{ columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', columnName: 'ColumnName', prompt: 'test', isEnabled: false }])
    });
  });

  it('set autofill prompt to column by id and list url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')/fields/getbyinternalnameortitle('ColumnName')?&$select=Id,Title,FieldTypeKind,AutofillInfo`) {
        return {
          Id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          Title: 'ColumnName',
          FieldTypeKind: 1,
          AutofillInfo: "{ \"LLM\": {\"IsEnabled\": true, \"Prompt\": \"test\" },\"PrebuiltModel\":null}"
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('%2Fsites%2Fsales%2FShared%20Documents')?$select=Id,BaseType`) {
        return {
          Id: "421b1e42-794b-4c71-93ac-5ed92488b67d",
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/machinelearning/SetColumnLLMInfo`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnTitle: 'ColumnName', prompt: 'new prompt', listUrl: '/Shared Documents' }) });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      autofillPrompt: 'new prompt',
      columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
      docLibId: `{421b1e42-794b-4c71-93ac-5ed92488b67d}`,
      isEnabled: true
    });
  });

  it('set autofill isEnable to false to column by id and list url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')/fields/getbyinternalnameortitle('ColumnName')?&$select=Id,Title,FieldTypeKind,AutofillInfo`) {
        return {
          Id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          Title: 'ColumnName',
          FieldTypeKind: 1,
          AutofillInfo: "{ \"LLM\": {\"IsEnabled\": true, \"Prompt\": \"test\" },\"PrebuiltModel\":null}"
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('%2Fsites%2Fsales%2FShared%20Documents')?$select=Id,BaseType`) {
        return {
          Id: "421b1e42-794b-4c71-93ac-5ed92488b67d",
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/machinelearning/SetColumnLLMInfo`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnTitle: 'ColumnName', isEnabled: false, listUrl: '/Shared Documents' }) });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      autofillPrompt: 'test',
      columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
      docLibId: `{421b1e42-794b-4c71-93ac-5ed92488b67d}`,
      isEnabled: false
    });
  });

  it('correctly handles error when list is not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')/fields/getbyid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')?&$select=Id,Title,FieldTypeKind,AutofillInfo`) {
        return {
          Id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          Title: 'ColumnName',
          FieldTypeKind: 1,
          AutofillInfo: null
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=Id,BaseType`) {
        throw {
          error: {
            "odata.error": {
              code: "-1, Microsoft.SharePoint.Client.ResourceNotFoundException",
              message: {
                lang: "en-US",
                value: "List does not exist. The page you selected contains a list that does not exist. It may have been deleted by another user."
              }
            }
          }
        };
      }

      throw `${opts.url} is invalid request`;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/machinelearning/SetSyntexPoweredColumnPrompts`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', prompt: 'test', verbose: true }) }), new CommandError('List does not exist. The page you selected contains a list that does not exist. It may have been deleted by another user.'));
  });

  it('correctly handles error when trying to apply autofill column to a SharePoint list that is not a document library and returns the error message "The specified list is not a document library."', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')/fields/getbyid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')?&$select=Id,Title,FieldTypeKind,AutofillInfo`) {
        return {
          Id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          Title: 'ColumnName',
          FieldTypeKind: 1,
          AutofillInfo: null
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=Id,BaseType`) {
        return {
          Id: "421b1e42-794b-4c71-93ac-5ed92488b67d",
          BaseType: 0
        };
      }

      throw `${opts.url} is invalid request`;
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', prompt: 'test', verbose: true }) }), new CommandError('The specified list is not a document library.'));
  });

  it('correctly handles error when trying to apply autofill to a column with an incorrect type."', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')/fields/getbyid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')?&$select=Id,Title,FieldTypeKind,AutofillInfo`) {
        return {
          Id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          Title: 'ColumnName',
          FieldTypeKind: 17,
          AutofillInfo: null
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=Id,BaseType`) {
        return {
          Id: "421b1e42-794b-4c71-93ac-5ed92488b67d",
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', prompt: 'test', verbose: true }) }), new CommandError('The specified column has incorrect type.'));
  });

  it('correctly handles error when applying autofill without the required prompt parameter.', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')/fields/getbyid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')?&$select=Id,Title,FieldTypeKind,AutofillInfo`) {
        return {
          Id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          Title: 'ColumnName',
          FieldTypeKind: 1,
          AutofillInfo: null
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=Id,BaseType`) {
        return {
          Id: "421b1e42-794b-4c71-93ac-5ed92488b67d",
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', columnId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d' }) }), new CommandError('The prompt parameter is required when setting the autofill column for the first time.'));
  });
});