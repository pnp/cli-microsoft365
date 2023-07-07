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
const command: Command = require('./list-add');

describe(commands.LIST_ADD, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets specified title for list', async () => {
    const expected = 'List 1';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.Title;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: expected, baseTemplate: 'GenericList', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified baseTemplate for list', async () => {
    const expected = 100;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.BaseTemplate;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets default baseTemplate for list', async () => {
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists`) {
        actual = opts.data.BaseTemplate;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'List 1', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, 100);
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

    await command.action(logger, { options: { title: 'List 1', description: expected, baseTemplate: 'GenericList', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', templateFeatureId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', schemaXml: expected, templateFeatureId: '00bfea71-de22-43b2-a848-c05709900100', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', allowDeletion: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', allowEveryoneViewItems: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', allowMultiResponses: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', contentTypesEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', crawlNonDefaultViews: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', defaultContentApprovalWorkflowId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', defaultDisplayFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', defaultEditFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified direction for list', async () => {
    const expected = 'LTR';
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.Direction;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', direction: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('sets specified disableCommenting for list', async () => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists`) {
        actual = opts.data.DisableCommenting;
        return { ErrorMessage: null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', disableCommenting: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', disableGridEditing: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', draftVersionVisibility: 'Author', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', emailAlias: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enableAssignToEmail: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enableAttachments: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enableDeployWithDependentList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enableFolderCreation: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enableMinorVersions: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enableModeration: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enablePeopleSelector: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enableResourceSelector: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enableSchemaCaching: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enableSyndication: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enableThrottling: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enableVersioning: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', enforceDataValidation: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', excludeFromOfflineClient: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', fetchPropertyBagForListView: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', followable: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', forceCheckout: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', forceDefaultContentType: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', hidden: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', includedInMyFilesScope: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', irmEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', irmExpire: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', irmReject: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', isApplicationList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', listExperienceOptions: 'NewExperience', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', majorVersionLimit: expected, enableVersioning: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', majorWithMinorVersionsLimit: expected, enableMinorVersions: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', multipleDataList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', navigateForFormsPages: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', needUpdateSiteClientTag: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', noCrawl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', onQuickLaunch: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', ordered: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', parserDisabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', readOnlyUI: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', readSecurity: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', requestAccessEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', restrictUserUpdates: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', sendToLocationName: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', sendToLocationUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', showUser: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', useFormsForDisplay: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', validationFormula: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', validationMessage: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
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

    await command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', writeSecurity: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, expected);
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, { options: { title: 'List 1', baseTemplate: 'GenericList', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } } as any),
      new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', title: 'List 1', baseTemplate: 'GenericList' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList' } }, commandInfo);
    assert(actual);
  });

  it('has correct baseTemplate specified', async () => {
    const baseTemplateValue = 'DocumentLibrary';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: baseTemplateValue } }, commandInfo);
    assert(actual === true);
  });

  it('fails if non existing baseTemplate specified', async () => {
    const baseTemplateValue = 'foo';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: baseTemplateValue } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the templateFeatureId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', templateFeatureId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the templateFeatureId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', templateFeatureId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the defaultContentApprovalWorkflowId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', defaultContentApprovalWorkflowId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the defaultContentApprovalWorkflowId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', defaultContentApprovalWorkflowId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails if non existing draftVersionVisibility specified', async () => {
    const draftVersionValue = 'NonExistingDraftVersionVisibility';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', draftVersionVisibility: draftVersionValue } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('has correct draftVersionVisibility specified', async () => {
    const draftVersionValue = 'Approver';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', draftVersionVisibility: draftVersionValue } }, commandInfo);
    assert(actual === true);
  });

  it('fails if emailAlias specified, but enableAssignToEmail is not true', async () => {
    const emailAliasValue = 'yourname@contoso.onmicrosoft.com';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', emailAlias: emailAliasValue } }, commandInfo);
    assert.strictEqual(actual, `emailAlias could not be set if enableAssignToEmail is not set to true. Please set enableAssignToEmail.`);
  });

  it('has correct emailAlias and enableAssignToEmail values specified', async () => {
    const emailAliasValue = 'yourname@contoso.onmicrosoft.com';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', emailAlias: emailAliasValue, enableAssignToEmail: true } }, commandInfo);
    assert(actual === true);
  });

  it('fails if non existing direction specified', async () => {
    const directionValue = 'abc';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', direction: directionValue } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('has correct direction value specified', async () => {
    const directionValue = 'LTR';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', direction: directionValue } }, commandInfo);
    assert(actual === true);
  });

  it('fails if majorVersionLimit specified, but enableVersioning is not true', async () => {
    const majorVersionLimitValue = 20;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', majorVersionLimit: majorVersionLimitValue } }, commandInfo);
    assert.strictEqual(actual, `majorVersionLimit option is only valid in combination with enableVersioning.`);
  });

  it('has correct majorVersionLimit and enableVersioning values specified', async () => {
    const majorVersionLimitValue = 20;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', majorVersionLimit: majorVersionLimitValue, enableVersioning: true } }, commandInfo);
    assert(actual === true);
  });

  it('fails if majorWithMinorVersionsLimit specified, but enableModeration is not true', async () => {
    const majorWithMinorVersionLimitValue = 20;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', majorWithMinorVersionsLimit: majorWithMinorVersionLimitValue } }, commandInfo);
    assert.strictEqual(actual, `majorWithMinorVersionsLimit option is only valid in combination with enableMinorVersions or enableModeration.`);
  });

  it('fails if non existing readSecurity specified', async () => {
    const readSecurityValue = 5;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', readSecurity: readSecurityValue } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails if non existing writeSecurity specified', async () => {
    const writeSecurityValue = 5;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', writeSecurity: writeSecurityValue } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('has correct readSecurity specified', async () => {
    const readSecurityValue = 2;
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', readSecurity: readSecurityValue } }, commandInfo);
    assert(actual === true);
  });

  it('fails if non existing listExperienceOptions specified', async () => {
    const listExperienceValue = 'NonExistingExperience';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', listExperienceOptions: listExperienceValue } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('has correct listExperienceOptions specified', async () => {
    const listExperienceValue = 'NewExperience';
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', listExperienceOptions: listExperienceValue } }, commandInfo);
    assert(actual === true);
  });

  it('returns listInstance object when list is added with correct values', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        return {
          "AllowContentTypes": true,
          "BaseTemplate": 100,
          "BaseType": 1,
          "ContentTypesEnabled": false,
          "CrawlNonDefaultViews": false,
          "Created": null,
          "CurrentChangeToken": null,
          "CustomActionElements": null,
          "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000",
          "DefaultItemOpenUseListSetting": false,
          "Description": "",
          "Direction": "none",
          "DocumentTemplateUrl": null,
          "DraftVersionVisibility": 0,
          "EnableAttachments": false,
          "EnableFolderCreation": true,
          "EnableMinorVersions": false,
          "EnableModeration": false,
          "EnableVersioning": false,
          "EntityTypeName": "Documents",
          "ExemptFromBlockDownloadOfNonViewableFiles": false,
          "FileSavePostProcessingEnabled": false,
          "ForceCheckout": false,
          "HasExternalDataSource": false,
          "Hidden": false,
          "Id": "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
          "ImagePath": null,
          "ImageUrl": null,
          "IrmEnabled": false,
          "IrmExpire": false,
          "IrmReject": false,
          "IsApplicationList": false,
          "IsCatalog": false,
          "IsPrivate": false,
          "ItemCount": 69,
          "LastItemDeletedDate": null,
          "LastItemModifiedDate": null,
          "LastItemUserModifiedDate": null,
          "ListExperienceOptions": 0,
          "ListItemEntityTypeFullName": null,
          "MajorVersionLimit": 0,
          "MajorWithMinorVersionsLimit": 0,
          "MultipleDataList": false,
          "NoCrawl": false,
          "ParentWebPath": null,
          "ParentWebUrl": null,
          "ParserDisabled": false,
          "ServerTemplateCanCreateFolders": true,
          "TemplateFeatureId": null,
          "Title": "List 1"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, title: 'List 1', baseTemplate: 'GenericList', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert(loggerLogSpy.calledWith({
      AllowContentTypes: true,
      BaseTemplate: 100,
      BaseType: 1,
      ContentTypesEnabled: false,
      CrawlNonDefaultViews: false,
      Created: null,
      CurrentChangeToken: null,
      CustomActionElements: null,
      DefaultContentApprovalWorkflowId: '00000000-0000-0000-0000-000000000000',
      DefaultItemOpenUseListSetting: false,
      Description: '',
      Direction: 'none',
      DocumentTemplateUrl: null,
      DraftVersionVisibility: 0,
      EnableAttachments: false,
      EnableFolderCreation: true,
      EnableMinorVersions: false,
      EnableModeration: false,
      EnableVersioning: false,
      EntityTypeName: 'Documents',
      ExemptFromBlockDownloadOfNonViewableFiles: false,
      FileSavePostProcessingEnabled: false,
      ForceCheckout: false,
      HasExternalDataSource: false,
      Hidden: false,
      Id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
      ImagePath: null,
      ImageUrl: null,
      IrmEnabled: false,
      IrmExpire: false,
      IrmReject: false,
      IsApplicationList: false,
      IsCatalog: false,
      IsPrivate: false,
      ItemCount: 69,
      LastItemDeletedDate: null,
      LastItemModifiedDate: null,
      LastItemUserModifiedDate: null,
      ListExperienceOptions: 0,
      ListItemEntityTypeFullName: null,
      MajorVersionLimit: 0,
      MajorWithMinorVersionsLimit: 0,
      MultipleDataList: false,
      NoCrawl: false,
      ParentWebPath: null,
      ParentWebUrl: null,
      ParserDisabled: false,
      ServerTemplateCanCreateFolders: true,
      TemplateFeatureId: null,
      Title: 'List 1'
    }));
  });
});
