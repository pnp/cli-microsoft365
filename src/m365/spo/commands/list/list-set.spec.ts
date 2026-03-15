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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command, { options } from './list-set.js';

describe(commands.LIST_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets specified title for list retrieved by title', async () => {
    const newTitle = 'List 1';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('Documents')/`) {
        actual = opts.data.Title;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ debug: true, title: 'Documents', newTitle: newTitle, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, newTitle);
  });

  it('sets specified title for list retrieved by url', async () => {
    const newTitle = 'List 1';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('%2Fsites%2Fproject-x%2Fdocuments')/`) {
        actual = opts.data.Title;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ debug: true, url: 'sites/project-x/documents', newTitle: newTitle, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, newTitle);
  });

  it('sets specified title for list', async () => {
    const newTitle = 'List 1';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.Title;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ debug: true, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', newTitle: newTitle, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, newTitle);
  });

  it('sets specified description for list', async () => {
    const expected = 'List 1 description';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.Description;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', description: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified templateFeatureId for list', async () => {
    const expected = '00bfea71-de22-43b2-a848-c05709900100';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.TemplateFeatureId;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', templateFeatureId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified allowDeletion for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.AllowDeletion;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowDeletion: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified allowEveryoneViewItems for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.AllowEveryoneViewItems;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowEveryoneViewItems: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified allowMultiResponses for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.AllowMultiResponses;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowMultiResponses: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified contentTypesEnabled for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.ContentTypesEnabled;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified crawlNonDefaultViews for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.CrawlNonDefaultViews;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', crawlNonDefaultViews: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified defaultContentApprovalWorkflowId for list', async () => {
    const expected = '00bfea71-de22-43b2-a848-c05709900100';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.DefaultContentApprovalWorkflowId;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultContentApprovalWorkflowId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified defaultDisplayFormUrl for list', async () => {
    const expected = '/sites/project-x/List%201/view.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.DefaultDisplayFormUrl;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultDisplayFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified defaultEditFormUrl for list', async () => {
    const expected = '/sites/project-x/List%201/edit.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.DefaultEditFormUrl;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultEditFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified direction for list', async () => {
    const expected = 'LTR';
    let actual = '';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.Direction;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', direction: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified disableCommenting for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.DisableCommenting;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', disableCommenting: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified disableGridEditing for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.DisableGridEditing;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', disableGridEditing: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified draftVersionVisibility for list', async () => {
    const expected = 1;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.DraftVersionVisibility;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', draftVersionVisibility: 'Author', webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified emailAlias for list', async () => {
    const expected = 'yourname@contoso.onmicrosoft.com';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EmailAlias;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', emailAlias: expected, enableAssignToEmail: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableAssignToEmail for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnableAssignToEmail;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAssignToEmail: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableAttachments for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnableAttachments;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAttachments: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableDeployWithDependentList for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnableDeployWithDependentList;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableDeployWithDependentList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableFolderCreation for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnableFolderCreation;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableFolderCreation: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableMinorVersions for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnableMinorVersions;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableMinorVersions: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableModeration for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnableModeration;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableModeration: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enablePeopleSelector for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnablePeopleSelector;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enablePeopleSelector: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableResourceSelector for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnableResourceSelector;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableResourceSelector: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableSchemaCaching for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnableSchemaCaching;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSchemaCaching: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableSyndication for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnableSyndication;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSyndication: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableThrottling for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnableThrottling;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableThrottling: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableVersioning for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnableVersioning;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableVersioning: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enforceDataValidation for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.EnforceDataValidation;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enforceDataValidation: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified excludeFromOfflineClient for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.ExcludeFromOfflineClient;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', excludeFromOfflineClient: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified fetchPropertyBagForListView for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.FetchPropertyBagForListView;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', fetchPropertyBagForListView: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified followable for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.Followable;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', followable: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified forceCheckout for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.ForceCheckout;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceCheckout: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified forceDefaultContentType for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.ForceDefaultContentType;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceDefaultContentType: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified hidden for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.Hidden;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', hidden: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified includedInMyFilesScope for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.IncludedInMyFilesScope;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', includedInMyFilesScope: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified irmEnabled for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.IrmEnabled;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified irmExpire for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.IrmExpire;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmExpire: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified irmReject for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.IrmReject;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmReject: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified isApplicationList for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.IsApplicationList;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', isApplicationList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified listExperienceOptions for list', async () => {
    const expected = 1;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.ListExperienceOptions;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', listExperienceOptions: 'NewExperience', webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified majorVersionLimit for list', async () => {
    const expected = 34;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.MajorVersionLimit;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: expected, enableVersioning: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified majorWithMinorVersionsLimit for list', async () => {
    const expected = 20;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.MajorWithMinorVersionsLimit;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorWithMinorVersionsLimit: expected, enableMinorVersions: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified multipleDataList for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.MultipleDataList;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', multipleDataList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified navigateForFormsPages for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.NavigateForFormsPages;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', navigateForFormsPages: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified needUpdateSiteClientTag for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.NeedUpdateSiteClientTag;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', needUpdateSiteClientTag: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified noCrawl for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.NoCrawl;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', noCrawl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified onQuickLaunch for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.OnQuickLaunch;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', onQuickLaunch: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified ordered for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.Ordered;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', ordered: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified parserDisabled for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.ParserDisabled;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', parserDisabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified readOnlyUI for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.ReadOnlyUI;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readOnlyUI: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified readSecurity for list', async () => {
    const expected = 2;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.ReadSecurity;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readSecurity: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified requestAccessEnabled for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.RequestAccessEnabled;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', requestAccessEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified restrictUserUpdates for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.RestrictUserUpdates;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', restrictUserUpdates: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified sendToLocationName for list', async () => {
    const expected = 'SendToLocation';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.SendToLocationName;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', sendToLocationName: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified sendToLocationUrl for list', async () => {
    const expected = '/sites/project-x/SendToLocation.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.SendToLocationUrl;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', sendToLocationUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified showUser for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.ShowUser;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', showUser: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified useFormsForDisplay for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.UseFormsForDisplay;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', useFormsForDisplay: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified validationFormula for list', async () => {
    const expected = `IF(fieldName=true);'truetest':'falsetest'`;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.ValidationFormula;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', validationFormula: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified validationMessage for list', async () => {
    const expected = 'Error on field x';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.ValidationMessage;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', validationMessage: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('sets specified writeSecurity for list', async () => {
    const expected = 4;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        actual = opts.data.WriteSecurity;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', writeSecurity: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(actual, expected);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      throw 'An error has occurred';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', newTitle: 'Test', webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await assert.rejects(command.action(logger, { options: parsedOptions.data! }), new CommandError('An error has occurred'));
  });

  it('automatically enables versioning when majorVersionLimit is specified', async () => {
    let enableVersioningValue: boolean | undefined;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        enableVersioningValue = opts.data.EnableVersioning;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: 50, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(enableVersioningValue, true);
  });

  it('does not override explicit enableVersioning value when majorVersionLimit is specified', async () => {
    let enableVersioningValue: boolean | undefined;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        enableVersioningValue = opts.data.EnableVersioning;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: 50, enableVersioning: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(enableVersioningValue, true);
  });

  it('sets versionAutoExpireTrim to true using CSOM for list retrieved by id', async () => {
    let csomData = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_vti_bin/client.svc/ProcessQuery`) {
        csomData = opts.data;
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.0.0", "ErrorInfo": null, "TraceCorrelationId": "fake" }]);
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', versionAutoExpireTrim: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert(csomData.indexOf('Name="DefaultTrimMode"><Parameter Type="Int32">2</Parameter>') > -1);
  });

  it('sets versionAutoExpireTrim to false using CSOM for list retrieved by id', async () => {
    let csomData = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_vti_bin/client.svc/ProcessQuery`) {
        csomData = opts.data;
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.0.0", "ErrorInfo": null, "TraceCorrelationId": "fake" }]);
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', versionAutoExpireTrim: false, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert(csomData.indexOf('Name="DefaultTrimMode"><Parameter Type="Int32">0</Parameter>') > -1);
  });

  it('sets versionExpireAfterDays using CSOM for list retrieved by id', async () => {
    let csomData = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_vti_bin/client.svc/ProcessQuery`) {
        csomData = opts.data;
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.0.0", "ErrorInfo": null, "TraceCorrelationId": "fake" }]);
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', versionExpireAfterDays: 30, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert(csomData.indexOf('Name="DefaultTrimMode"><Parameter Type="Int32">1</Parameter>') > -1);
    assert(csomData.indexOf('Name="DefaultExpireAfterDays"><Parameter Type="Int32">30</Parameter>') > -1);
  });

  it('sets versionAutoExpireTrim using CSOM for list retrieved by title', async () => {
    let csomData = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_vti_bin/client.svc/ProcessQuery`) {
        csomData = opts.data;
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.0.0", "ErrorInfo": null, "TraceCorrelationId": "fake" }]);
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ title: 'Documents', versionAutoExpireTrim: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert(csomData.indexOf('Name="GetByTitle"') > -1);
    assert(csomData.indexOf('Name="DefaultTrimMode"><Parameter Type="Int32">2</Parameter>') > -1);
  });

  it('sets versionExpireAfterDays using CSOM for list retrieved by url', async () => {
    let csomData = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_vti_bin/client.svc/ProcessQuery`) {
        csomData = opts.data;
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.0.0", "ErrorInfo": null, "TraceCorrelationId": "fake" }]);
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ url: 'sites/project-x/documents', versionExpireAfterDays: 60, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert(csomData.indexOf('Name="GetList"') > -1);
    assert(csomData.indexOf('Name="DefaultTrimMode"><Parameter Type="Int32">1</Parameter>') > -1);
    assert(csomData.indexOf('Name="DefaultExpireAfterDays"><Parameter Type="Int32">60</Parameter>') > -1);
  });

  it('handles CSOM error when setting version policies', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.0.0", "ErrorInfo": { "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "fake", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.SPException" }, "TraceCorrelationId": "fake" }]);
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', versionAutoExpireTrim: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await assert.rejects(command.action(logger, { options: parsedOptions.data! }), new CommandError('An error has occurred'));
  });

  it('sends both REST and CSOM requests when setting regular and version policy options', async () => {
    let restCalled = false;
    let csomCalled = false;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        restCalled = true;
        return { ErrorMessage: null };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_vti_bin/client.svc/ProcessQuery`) {
        csomCalled = true;
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.0.0", "ErrorInfo": null, "TraceCorrelationId": "fake" }]);
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', newTitle: 'New Title', versionAutoExpireTrim: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(restCalled, true);
    assert.strictEqual(csomCalled, true);
  });

  it('skips REST request when only version policy options are specified', async () => {
    let restCalled = false;
    let csomCalled = false;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')/`) {
        restCalled = true;
        return { ErrorMessage: null };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_vti_bin/client.svc/ProcessQuery`) {
        csomCalled = true;
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.0.0", "ErrorInfo": null, "TraceCorrelationId": "fake" }]);
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', versionAutoExpireTrim: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert.strictEqual(restCalled, false);
    assert.strictEqual(csomCalled, true);
  });

  it('uses new title for CSOM lookup when title is changed via REST', async () => {
    let csomData = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('Documents')/`) {
        return { ErrorMessage: null };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_vti_bin/client.svc/ProcessQuery`) {
        csomData = opts.data;
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.0.0", "ErrorInfo": null, "TraceCorrelationId": "fake" }]);
      }

      throw 'Invalid request';
    });

    const parsedOptions = commandOptionsSchema.safeParse({ title: 'Documents', newTitle: 'Documents Updated', versionAutoExpireTrim: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' });
    await command.action(logger, { options: parsedOptions.data! });
    assert(csomData.indexOf('Name="GetByTitle"') > -1);
    assert(csomData.indexOf('Documents Updated') > -1);
  });

  it('fails validation if the neither id, title or url is set', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if no option to update is specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if id and title is set', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', title: 'Documents' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', newTitle: 'Test' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: true });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', newTitle: 'Test' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the templateFeatureId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', templateFeatureId: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if the templateFeatureId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', templateFeatureId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the defaultContentApprovalWorkflowId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultContentApprovalWorkflowId: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if the defaultContentApprovalWorkflowId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultContentApprovalWorkflowId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' });
    assert.strictEqual(actual.success, true);
  });

  it('fails if non existing draftVersionVisibility specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', draftVersionVisibility: 'NonExistingDraftVersionVisibility' });
    assert.notStrictEqual(actual.success, true);
  });

  it('has correct draftVersionVisibility specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', draftVersionVisibility: 'Approver' });
    assert.strictEqual(actual.success, true);
  });

  it('fails if emailAlias specified, but enableAssignToEmail is not true', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', emailAlias: 'yourname@contoso.onmicrosoft.com' });
    assert.notStrictEqual(actual.success, true);
  });

  it('has correct emailAlias and enableAssignToEmail values specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', emailAlias: 'yourname@contoso.onmicrosoft.com', enableAssignToEmail: true });
    assert.strictEqual(actual.success, true);
  });

  it('fails if non existing direction specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', direction: 'abc' });
    assert.notStrictEqual(actual.success, true);
  });

  it('has correct direction specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', direction: 'LTR' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if majorVersionLimit specified without enableVersioning', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: 20 });
    assert.strictEqual(actual.success, true);
  });

  it('has correct majorVersionLimit and enableVersioning values specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: 20, enableVersioning: true });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if versionExpireAfterDays is not a valid positive integer', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', versionExpireAfterDays: -1 });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if versionExpireAfterDays is zero', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', versionExpireAfterDays: 0 });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if versionExpireAfterDays is a valid positive integer', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', versionExpireAfterDays: 30 });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if versionExpireAfterDays and versionAutoExpireTrim true are both specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', versionExpireAfterDays: 30, versionAutoExpireTrim: true });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if versionExpireAfterDays and versionAutoExpireTrim false are both specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', versionExpireAfterDays: 30, versionAutoExpireTrim: false });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if only versionAutoExpireTrim is specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', versionAutoExpireTrim: true });
    assert.strictEqual(actual.success, true);
  });

  it('fails if majorWithMinorVersionsLimit specified, but enableModeration is not true', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorWithMinorVersionsLimit: 20 });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails if non existing readSecurity specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readSecurity: 5 });
    assert.notStrictEqual(actual.success, true);
  });

  it('has correct readSecurity specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readSecurity: 2 });
    assert.strictEqual(actual.success, true);
  });

  it('fails if non existing listExperienceOptions specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', listExperienceOptions: 'NonExistingExperience' });
    assert.notStrictEqual(actual.success, true);
  });

  it('has correct listExperienceOptions specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', listExperienceOptions: 'NewExperience' });
    assert.strictEqual(actual.success, true);
  });

  it('fails if non existing writeSecurity specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', writeSecurity: 5 });
    assert.notStrictEqual(actual.success, true);
  });

  it('has correct writeSecurity specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', writeSecurity: 4 });
    assert.strictEqual(actual.success, true);
  });
});
