import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./list-set');

describe(commands.LIST_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets specified title for list', (done) => {
    const expected = 'List 1';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.Title;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', title: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified description for list', (done) => {
    const expected = 'List 1 description';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.Description;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', description: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified templateFeatureId for list', (done) => {
    const expected = '00bfea71-de22-43b2-a848-c05709900100';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.TemplateFeatureId;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', templateFeatureId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified schemaXml for list', (done) => {
    const expected = `<List Title=\'List 1' ID='BE9CE88C-EF3A-4A61-9A8E-F8C038442227'></List>`;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.SchemaXml;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', schemaXml: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified allowDeletion for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.AllowDeletion;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowDeletion: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified allowEveryoneViewItems for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.AllowEveryoneViewItems;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowEveryoneViewItems: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified allowMultiResponses for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.AllowMultiResponses;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowMultiResponses: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified contentTypesEnabled for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ContentTypesEnabled;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified crawlNonDefaultViews for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.CrawlNonDefaultViews;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', crawlNonDefaultViews: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified defaultContentApprovalWorkflowId for list', (done) => {
    const expected = '00bfea71-de22-43b2-a848-c05709900100';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.DefaultContentApprovalWorkflowId;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultContentApprovalWorkflowId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified defaultDisplayFormUrl for list', (done) => {
    const expected = '/sites/project-x/List%201/view.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.DefaultDisplayFormUrl;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultDisplayFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified defaultEditFormUrl for list', (done) => {
    const expected = '/sites/project-x/List%201/edit.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.DefaultEditFormUrl;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultEditFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified direction for list', (done) => {
    const expected = 'LTR';
    let actual = '';

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid`) > -1) {
        actual = opts.data.Direction;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', direction: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified disableGridEditing for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.DisableGridEditing;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', disableGridEditing: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified draftVersionVisibility for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.DraftVersionVisibility;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', draftVersionVisibility: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified emailAlias for list', (done) => {
    const expected = 'yourname@contoso.onmicrosoft.com';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EmailAlias;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', emailAlias: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableAssignToEmail for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableAssignToEmail;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAssignToEmail: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableAttachments for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableAttachments;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAttachments: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableDeployWithDependentList for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableDeployWithDependentList;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableDeployWithDependentList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableFolderCreation for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableFolderCreation;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableFolderCreation: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableMinorVersions for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableMinorVersions;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableMinorVersions: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableModeration for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableModeration;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableModeration: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enablePeopleSelector for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnablePeopleSelector;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enablePeopleSelector: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableResourceSelector for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableResourceSelector;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableResourceSelector: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableSchemaCaching for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableSchemaCaching;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSchemaCaching: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableSyndication for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableSyndication;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSyndication: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableThrottling for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableThrottling;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableThrottling: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableVersioning for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnableVersioning;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableVersioning: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enforceDataValidation for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.EnforceDataValidation;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enforceDataValidation: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified excludeFromOfflineClient for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ExcludeFromOfflineClient;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', excludeFromOfflineClient: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified fetchPropertyBagForListView for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.FetchPropertyBagForListView;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', fetchPropertyBagForListView: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified followable for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.Followable;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', followable: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified forceCheckout for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ForceCheckout;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceCheckout: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified forceDefaultContentType for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ForceDefaultContentType;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceDefaultContentType: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified hidden for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.Hidden;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', hidden: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified includedInMyFilesScope for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.IncludedInMyFilesScope;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', includedInMyFilesScope: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified irmEnabled for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.IrmEnabled;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified irmExpire for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.IrmExpire;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmExpire: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified irmReject for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.IrmReject;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmReject: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified isApplicationList for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.IsApplicationList;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', isApplicationList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified listExperienceOptions for list', (done) => {
    const expected = 'NewExperience';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ListExperienceOptions;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', listExperienceOptions: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified majorVersionLimit for list', (done) => {
    const expected = 34;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.MajorVersionLimit;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: expected, enableVersioning: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified majorWithMinorVersionsLimit for list', (done) => {
    const expected = 20;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.MajorWithMinorVersionsLimit;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorWithMinorVersionsLimit: expected, enableMinorVersions: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified multipleDataList for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.MultipleDataList;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', multipleDataList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified navigateForFormsPages for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.NavigateForFormsPages;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', navigateForFormsPages: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified needUpdateSiteClientTag for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.NeedUpdateSiteClientTag;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', needUpdateSiteClientTag: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified noCrawl for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.NoCrawl;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', noCrawl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified onQuickLaunch for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.OnQuickLaunch;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', onQuickLaunch: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified ordered for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.Ordered;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', ordered: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified parserDisabled for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ParserDisabled;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', parserDisabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified readOnlyUI for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ReadOnlyUI;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readOnlyUI: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified readSecurity for list', (done) => {
    const expected = 2;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ReadSecurity;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readSecurity: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified requestAccessEnabled for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.RequestAccessEnabled;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', requestAccessEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified restrictUserUpdates for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.RestrictUserUpdates;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', restrictUserUpdates: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified sendToLocationName for list', (done) => {
    const expected = 'SendToLocation';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.SendToLocationName;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', sendToLocationName: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified sendToLocationUrl for list', (done) => {
    const expected = '/sites/project-x/SendToLocation.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.SendToLocationUrl;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', sendToLocationUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified showUser for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ShowUser;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', showUser: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified useFormsForDisplay for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.UseFormsForDisplay;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', useFormsForDisplay: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified validationFormula for list', (done) => {
    const expected = `IF(fieldName=true);'truetest':'falsetest'`;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ValidationFormula;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', validationFormula: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified validationMessage for list', (done) => {
    const expected = 'Error on field x';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.ValidationMessage;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', validationMessage: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified writeSecurity for list', (done) => {
    const expected = 4;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        actual = opts.data.WriteSecurity;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', writeSecurity: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, { options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: 'true' } }, commandInfo);
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

  it('fails validation if the allowDeletion option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowDeletion: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the allowDeletion option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowDeletion: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the allowEveryoneViewItems option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowEveryoneViewItems: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the allowEveryoneViewItems option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowEveryoneViewItems: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the allowMultiResponses option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowMultiResponses: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the allowMultiResponses option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowMultiResponses: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the contentTypesEnabled option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the contentTypesEnabled option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the crawlNonDefaultViews option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', crawlNonDefaultViews: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the crawlNonDefaultViews option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', crawlNonDefaultViews: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the disableGridEditing option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', disableGridEditing: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the disableGridEditing option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', disableGridEditing: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enableAssignToEmail option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAssignToEmail: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enableAssignToEmail option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAssignToEmail: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enableAttachments option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAttachments: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enableAttachments option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAttachments: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enableDeployWithDependentList option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableDeployWithDependentList: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enableDeployWithDependentList option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableDeployWithDependentList: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enableFolderCreation option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableFolderCreation: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enableFolderCreation option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableFolderCreation: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enableMinorVersions option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableMinorVersions: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enableMinorVersions option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableMinorVersions: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enableModeration option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableModeration: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enableModeration option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableModeration: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enablePeopleSelector option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enablePeopleSelector: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enablePeopleSelector option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enablePeopleSelector: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enableResourceSelector option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableResourceSelector: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enableResourceSelector option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableResourceSelector: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enableSchemaCaching option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSchemaCaching: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enableSchemaCaching option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSchemaCaching: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enableSyndication option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSyndication: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enableSyndication option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSyndication: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enableThrottling option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableThrottling: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enableThrottling option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableThrottling: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enableVersioning option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableVersioning: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enableVersioning option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableVersioning: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the enforceDataValidation option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enforceDataValidation: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the enforceDataValidation option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enforceDataValidation: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the excludeFromOfflineClient option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', excludeFromOfflineClient: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the excludeFromOfflineClient option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', excludeFromOfflineClient: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the fetchPropertyBagForListView option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', fetchPropertyBagForListView: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the fetchPropertyBagForListView option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', fetchPropertyBagForListView: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the followable option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', followable: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the followable option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', followable: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the forceCheckout option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceCheckout: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the forceCheckout option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceCheckout: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the forceDefaultContentType option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceDefaultContentType: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the forceDefaultContentType option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceDefaultContentType: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the hidden option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', hidden: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the hidden option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', hidden: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the includedInMyFilesScope option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', includedInMyFilesScope: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the includedInMyFilesScope option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', includedInMyFilesScope: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the irmEnabled option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmEnabled: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the irmEnabled option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmEnabled: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the irmExpire option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmExpire: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the irmExpire option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmExpire: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the irmReject option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmReject: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the irmReject option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmReject: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the isApplicationList option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', isApplicationList: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the isApplicationList option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', isApplicationList: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the multipleDataList option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', multipleDataList: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the multipleDataList option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', multipleDataList: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the navigateForFormsPages option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', navigateForFormsPages: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the navigateForFormsPages option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', navigateForFormsPages: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the needUpdateSiteClientTag option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', needUpdateSiteClientTag: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the needUpdateSiteClientTag option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', needUpdateSiteClientTag: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the noCrawl option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', noCrawl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the noCrawl option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', noCrawl: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the onQuickLaunch option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', onQuickLaunch: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the onQuickLaunch option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', onQuickLaunch: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the ordered option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', ordered: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the ordered option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', ordered: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the parserDisabled option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', parserDisabled: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the parserDisabled option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', parserDisabled: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the readOnlyUI option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readOnlyUI: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the readOnlyUI option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readOnlyUI: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the requestAccessEnabled option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', requestAccessEnabled: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the requestAccessEnabled option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', requestAccessEnabled: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the restrictUserUpdates option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', restrictUserUpdates: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the restrictUserUpdates option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', restrictUserUpdates: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the showUser option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', showUser: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the showUser option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', showUser: 'true' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the useFormsForDisplay option is not a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', useFormsForDisplay: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the useFormsForDisplay option is a valid Boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', useFormsForDisplay: 'true' } }, commandInfo);
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
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', emailAlias: emailAliasValue, enableAssignToEmail: 'true' } }, commandInfo);
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
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: majorVersionLimitValue, enableVersioning: 'true' } }, commandInfo);
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