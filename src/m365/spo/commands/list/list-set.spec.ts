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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./list-set');

describe(commands.LIST_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.LIST_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets specified title for list retrieved by title', async () => {
    const newTitle = 'List 1';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/lists/getByTitle(\'Documents\')`) > -1) {
        actual = opts.data.Title;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, title: 'Documents', newTitle: newTitle, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, newTitle);
  });

  it('sets specified title for list retrieved by url', async () => {
    const newTitle = 'List 1';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fproject-x%2Fdocuments')`) > -1) {
        actual = opts.data.Title;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, url: 'sites/project-x/documents', newTitle: newTitle, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, newTitle);
  });

  it('sets specified title for list', async () => {
    const newTitle = 'List 1';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'3EA5A977-315E-4E25-8B0F-E4F949BF6B8F')`) > -1) {
        actual = opts.data.Title;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', newTitle: newTitle, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, newTitle);
  });

  it('sets specified description for list', async () => {
    const expected = 'List 1 description';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.Description;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', description: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified templateFeatureId for list', async () => {
    const expected = '00bfea71-de22-43b2-a848-c05709900100';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.TemplateFeatureId;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', templateFeatureId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified schemaXml for list', async () => {
    const expected = `<List Title=\'List 1' ID='BE9CE88C-EF3A-4A61-9A8E-F8C038442227'></List>`;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.SchemaXml;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', schemaXml: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified allowDeletion for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.AllowDeletion;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowDeletion: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified allowEveryoneViewItems for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.AllowEveryoneViewItems;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowEveryoneViewItems: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified allowMultiResponses for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.AllowMultiResponses;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowMultiResponses: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified contentTypesEnabled for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ContentTypesEnabled;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified crawlNonDefaultViews for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.CrawlNonDefaultViews;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', crawlNonDefaultViews: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified defaultContentApprovalWorkflowId for list', async () => {
    const expected = '00bfea71-de22-43b2-a848-c05709900100';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.DefaultContentApprovalWorkflowId;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultContentApprovalWorkflowId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified defaultDisplayFormUrl for list', async () => {
    const expected = '/sites/project-x/List%201/view.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.DefaultDisplayFormUrl;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultDisplayFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified defaultEditFormUrl for list', async () => {
    const expected = '/sites/project-x/List%201/edit.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.DefaultEditFormUrl;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultEditFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified direction for list', async () => {
    const expected = 'LTR';
    let actual = '';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid`) > -1) {
        actual = opts.data.Direction;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', direction: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified disableGridEditing for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.DisableGridEditing;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', disableGridEditing: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified draftVersionVisibility for list', async () => {
    const expected = 1;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.DraftVersionVisibility;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', draftVersionVisibility: 'Author', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified emailAlias for list', async () => {
    const expected = 'yourname@contoso.onmicrosoft.com';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EmailAlias;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', emailAlias: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableAssignToEmail for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableAssignToEmail;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAssignToEmail: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableAttachments for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableAttachments;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAttachments: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableDeployWithDependentList for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableDeployWithDependentList;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableDeployWithDependentList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableFolderCreation for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableFolderCreation;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableFolderCreation: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableMinorVersions for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableMinorVersions;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableMinorVersions: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableModeration for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableModeration;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableModeration: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enablePeopleSelector for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnablePeopleSelector;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enablePeopleSelector: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableResourceSelector for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableResourceSelector;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableResourceSelector: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableSchemaCaching for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableSchemaCaching;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSchemaCaching: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableSyndication for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableSyndication;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSyndication: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableThrottling for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableThrottling;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableThrottling: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enableVersioning for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableVersioning;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableVersioning: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified enforceDataValidation for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnforceDataValidation;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enforceDataValidation: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified excludeFromOfflineClient for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ExcludeFromOfflineClient;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', excludeFromOfflineClient: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified fetchPropertyBagForListView for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.FetchPropertyBagForListView;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', fetchPropertyBagForListView: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified followable for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.Followable;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', followable: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified forceCheckout for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ForceCheckout;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceCheckout: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified forceDefaultContentType for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ForceDefaultContentType;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceDefaultContentType: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified hidden for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.Hidden;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', hidden: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified includedInMyFilesScope for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.IncludedInMyFilesScope;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', includedInMyFilesScope: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified irmEnabled for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.IrmEnabled;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified irmExpire for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.IrmExpire;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmExpire: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified irmReject for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.IrmReject;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmReject: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified isApplicationList for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.IsApplicationList;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', isApplicationList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified listExperienceOptions for list', async () => {
    const expected = 1;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ListExperienceOptions;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', listExperienceOptions: 'NewExperience', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified majorVersionLimit for list', async () => {
    const expected = 34;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.MajorVersionLimit;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: expected, enableVersioning: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified majorWithMinorVersionsLimit for list', async () => {
    const expected = 20;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.MajorWithMinorVersionsLimit;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorWithMinorVersionsLimit: expected, enableMinorVersions: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified multipleDataList for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.MultipleDataList;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', multipleDataList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified navigateForFormsPages for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.NavigateForFormsPages;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', navigateForFormsPages: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified needUpdateSiteClientTag for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.NeedUpdateSiteClientTag;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', needUpdateSiteClientTag: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified noCrawl for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.NoCrawl;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', noCrawl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified onQuickLaunch for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.OnQuickLaunch;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', onQuickLaunch: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified ordered for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.Ordered;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', ordered: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified parserDisabled for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ParserDisabled;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', parserDisabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified readOnlyUI for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ReadOnlyUI;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readOnlyUI: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified readSecurity for list', async () => {
    const expected = 2;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ReadSecurity;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readSecurity: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified requestAccessEnabled for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.RequestAccessEnabled;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', requestAccessEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified restrictUserUpdates for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.RestrictUserUpdates;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', restrictUserUpdates: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified sendToLocationName for list', async () => {
    const expected = 'SendToLocation';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.SendToLocationName;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', sendToLocationName: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified sendToLocationUrl for list', async () => {
    const expected = '/sites/project-x/SendToLocation.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.SendToLocationUrl;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', sendToLocationUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified showUser for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ShowUser;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', showUser: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified useFormsForDisplay for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.UseFormsForDisplay;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', useFormsForDisplay: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified validationFormula for list', async () => {
    const expected = `IF(fieldName=true);'truetest':'falsetest'`;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ValidationFormula;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', validationFormula: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified validationMessage for list', async () => {
    const expected = 'Error on field x';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ValidationMessage;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', validationMessage: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified writeSecurity for list', async () => {
    const expected = 4;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.WriteSecurity;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', writeSecurity: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, { options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } } as any), new CommandError('An error has occurred'));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('offers autocomplete for the direction option', () => {
    const options = command.options;
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--direction') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('fails validation if the neither id, title or url is set', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id and title is set', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', title: 'Documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the templateFeatureId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', templateFeatureId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the templateFeatureId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', templateFeatureId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the defaultContentApprovalWorkflowId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultContentApprovalWorkflowId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the defaultContentApprovalWorkflowId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultContentApprovalWorkflowId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails if non existing draftVersionVisibility specified', async () => {
    const draftVersionValue = 'NonExistingDraftVersionVisibility';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', draftVersionVisibility: draftVersionValue } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('has correct draftVersionVisibility specified', async () => {
    const draftVersionValue = 'Approver';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', draftVersionVisibility: draftVersionValue } }, commandInfo);
    assert(actual === true);
  });

  it('fails if emailAlias specified, but enableAssignToEmail is not true', async () => {
    const emailAliasValue = 'yourname@contoso.onmicrosoft.com';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', emailAlias: emailAliasValue } }, commandInfo);
    assert.strictEqual(actual, `emailAlias could not be set if enableAssignToEmail is not set to true. Please set enableAssignToEmail.`);
  });

  it('has correct emailAlias and enableAssignToEmail values specified', async () => {
    const emailAliasValue = 'yourname@contoso.onmicrosoft.com';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', emailAlias: emailAliasValue, enableAssignToEmail: true } }, commandInfo);
    assert(actual === true);
  });

  it('fails if non existing direction specified', async () => {
    const directionValue = 'abc';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', direction: directionValue } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('has correct direction specified', async () => {
    const directionValue = 'LTR';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', direction: directionValue } }, commandInfo);
    assert(actual === true);
  });

  it('fails if majorVersionLimit specified, but enableVersioning is not true', async () => {
    const majorVersionLimitValue = 20;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: majorVersionLimitValue } }, commandInfo);
    assert.strictEqual(actual, `majorVersionLimit option is only valid in combination with enableVersioning.`);
  });

  it('has correct majorVersionLimit and enableVersioning values specified', async () => {
    const majorVersionLimitValue = 20;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: majorVersionLimitValue, enableVersioning: true } }, commandInfo);
    assert(actual === true);
  });

  it('fails if majorWithMinorVersionsLimit specified, but enableModeration is not true', async () => {
    const majorWithMinorVersionLimitValue = 20;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorWithMinorVersionsLimit: majorWithMinorVersionLimitValue } }, commandInfo);
    assert.strictEqual(actual, `majorWithMinorVersionsLimit option is only valid in combination with enableMinorVersions or enableModeration.`);
  });


  it('fails if non existing readSecurity specified', async () => {
    const readSecurityValue = 5;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readSecurity: readSecurityValue } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('has correct readSecurity specified', async () => {
    const readSecurityValue = 2;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readSecurity: readSecurityValue } }, commandInfo);
    assert(actual === true);
  });

  it('fails if non existing listExperienceOptions specified', async () => {
    const listExperienceValue = 'NonExistingExperience';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', listExperienceOptions: listExperienceValue } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('has correct listExperienceOptions specified', async () => {
    const listExperienceValue = 'NewExperience';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', listExperienceOptions: listExperienceValue } }, commandInfo);
    assert(actual === true);
  });

  it('fails if non existing writeSecurity specified', async () => {
    const writeSecurityValue = 5;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', writeSecurity: writeSecurityValue } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('has correct writeSecurity specified', async () => {
    const writeSecurityValue = 4;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', writeSecurity: writeSecurityValue } }, commandInfo);
    assert(actual === true);
  });
});
