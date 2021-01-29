import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./apppage-add');

describe(commands.APPPAGE_ADD, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    Utils.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APPPAGE_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a single-part app page', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/CreateAppPage`) > -1 &&
        opts.data.webPartDataAsJson ===
        "{\"id\": \"e84a4f44-30d2-4962-b203-f8bf42114860\", \"instanceId\": \"15353e8b-cb55-4794-b871-4cd74abf78b4\", \"title\": \"Milestone Tracking\", \"description\": \"A tool used for tracking project milestones\", \"serverProcessedContent\": { \"htmlStrings\": {}, \"searchablePlainTexts\": {}, \"imageSources\": {}, \"links\": {} }, \"dataVersion\": \"1.0\", \"properties\": {\"description\": \"Milestone Tracking\"}}") {
        return Promise.resolve({
          value: "SitePages/lp4blf70.aspx"
        });
      }
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/UpdateAppPage`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          "pageId": 20,
          "webPartDataAsJson": "{\"id\": \"e84a4f44-30d2-4962-b203-f8bf42114860\", \"instanceId\": \"15353e8b-cb55-4794-b871-4cd74abf78b4\", \"title\": \"Milestone Tracking\", \"description\": \"A tool used for tracking project milestones\", \"serverProcessedContent\": { \"htmlStrings\": {}, \"searchablePlainTexts\": {}, \"imageSources\": {}, \"links\": {} }, \"dataVersion\": \"1.0\", \"properties\": {\"description\": \"Milestone Tracking\"}}",
          "title": "test-single"
        })) {
        return Promise.resolve({
          value: "SitePages/lp4blf70.aspx"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/SitePages/lp4blf70.aspx')?$expand=ListItemAllFields`) > -1) {
        return Promise.resolve({
          "ListItemAllFields": {
            "FileSystemObjectType": 0,
            "Id": 20,
            "ServerRedirectedEmbedUri": null,
            "ServerRedirectedEmbedUrl": "",
            "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180031B3A1F0639B70418205BE3DDA9C3044",
            "ComplianceAssetId": null,
            "WikiField": null,
            "Title": null,
            "CanvasContent1": "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;addedFromPersistedData&quot;&#58;&quot;True&quot;,&quot;controlType&quot;&#58;3,&quot;id&quot;&#58;&quot;15353e8b-cb55-4794-b871-4cd74abf78b4&quot;,&quot;webPartId&quot;&#58;&quot;e84a4f44-30d2-4962-b203-f8bf42114860&quot;&#125;\"><div data-sp-webpart=\"\" data-sp-webpartdataversion=\"1.0\" data-sp-webpartdata=\"&#123;&quot;id&quot;&#58;&quot;e84a4f44-30d2-4962-b203-f8bf42114860&quot;,&quot;instanceId&quot;&#58;&quot;15353e8b-cb55-4794-b871-4cd74abf78b4&quot;,&quot;title&quot;&#58;&quot;Milestone Tracking&quot;,&quot;description&quot;&#58;&quot;A tool used for tracking project milestones&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.0&quot;,&quot;properties&quot;&#58;&#123;&quot;description&quot;&#58;&quot;Milestone Tracking&quot;&#125;&#125;\"><div data-sp-componentid=\"\">e84a4f44-30d2-4962-b203-f8bf42114860</div><div data-sp-htmlproperties=\"\"></div></div></div></div>",
            "BannerImageUrl": null,
            "Description": null,
            "PromotedState": 0,
            "FirstPublishedDate": null,
            "LayoutWebpartsContent": null,
            "OData__AuthorBylineId": null,
            "_AuthorBylineStringId": null,
            "OData__TopicHeader": null,
            "OData__SPSitePageFlags": null,
            "OData__OriginalSourceUrl": null,
            "OData__OriginalSourceSiteId": null,
            "OData__OriginalSourceWebId": null,
            "OData__OriginalSourceListId": null,
            "OData__OriginalSourceItemId": null,
            "ID": 20,
            "Created": "2020-12-11T06:40:05",
            "AuthorId": 11,
            "Modified": "2020-12-11T06:40:05",
            "EditorId": 11,
            "OData__CopySource": null,
            "CheckoutUserId": 11,
            "OData__UIVersionString": "0.1",
            "GUID": "fba3d8ca-790d-4134-a276-7528c32d6b9c"
          },
          "CheckInComment": "",
          "CheckOutType": 0,
          "ContentTag": "{D1408169-EBC1-4A96-B839-95E2D4F439B3},2,2",
          "CustomizedPageStatus": 0,
          "ETag": "\"{D1408169-EBC1-4A96-B839-95E2D4F439B3},2\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "2940",
          "Level": 255,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "lp4blf70.aspx",
          "ServerRelativeUrl": "/SitePages/lp4blf70.aspx",
          "TimeCreated": "2020-12-11T14:40:05Z",
          "TimeLastModified": "2020-12-11T14:40:05Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "d1408169-ebc1-4a96-b839-95e2d4f439b3"
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, title: 'test-single', webUrl: 'https://contoso.sharepoint.com/', webPartData: "{\"id\": \"e84a4f44-30d2-4962-b203-f8bf42114860\", \"instanceId\": \"15353e8b-cb55-4794-b871-4cd74abf78b4\", \"title\": \"Milestone Tracking\", \"description\": \"A tool used for tracking project milestones\", \"serverProcessedContent\": { \"htmlStrings\": {}, \"searchablePlainTexts\": {}, \"imageSources\": {}, \"links\": {} }, \"dataVersion\": \"1.0\", \"properties\": {\"description\": \"Milestone Tracking\"}}" } }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined')
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a single-part app page showing on quicklaunch', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/CreateAppPage`) > -1 &&
        opts.data.webPartDataAsJson ===
        "{\"id\": \"e84a4f44-30d2-4962-b203-f8bf42114860\", \"instanceId\": \"15353e8b-cb55-4794-b871-4cd74abf78b4\", \"title\": \"Milestone Tracking\", \"description\": \"A tool used for tracking project milestones\", \"serverProcessedContent\": { \"htmlStrings\": {}, \"searchablePlainTexts\": {}, \"imageSources\": {}, \"links\": {} }, \"dataVersion\": \"1.0\", \"properties\": {\"description\": \"Milestone Tracking\"}}") {
        return Promise.resolve({
          value: "SitePages/lp4blf70.aspx"
        });
      }
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/UpdateAppPage`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          "pageId": 20,
          "webPartDataAsJson": "{\"id\": \"e84a4f44-30d2-4962-b203-f8bf42114860\", \"instanceId\": \"15353e8b-cb55-4794-b871-4cd74abf78b4\", \"title\": \"Milestone Tracking\", \"description\": \"A tool used for tracking project milestones\", \"serverProcessedContent\": { \"htmlStrings\": {}, \"searchablePlainTexts\": {}, \"imageSources\": {}, \"links\": {} }, \"dataVersion\": \"1.0\", \"properties\": {\"description\": \"Milestone Tracking\"}}",
          "title": "test-single",
          "includeInNavigation": true
        })) {
        return Promise.resolve({
          value: "SitePages/lp4blf70.aspx"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/SitePages/lp4blf70.aspx')?$expand=ListItemAllFields`) > -1) {
        return Promise.resolve({
          "ListItemAllFields": {
            "FileSystemObjectType": 0,
            "Id": 20,
            "ServerRedirectedEmbedUri": null,
            "ServerRedirectedEmbedUrl": "",
            "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180031B3A1F0639B70418205BE3DDA9C3044",
            "ComplianceAssetId": null,
            "WikiField": null,
            "Title": null,
            "CanvasContent1": "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;addedFromPersistedData&quot;&#58;&quot;True&quot;,&quot;controlType&quot;&#58;3,&quot;id&quot;&#58;&quot;15353e8b-cb55-4794-b871-4cd74abf78b4&quot;,&quot;webPartId&quot;&#58;&quot;e84a4f44-30d2-4962-b203-f8bf42114860&quot;&#125;\"><div data-sp-webpart=\"\" data-sp-webpartdataversion=\"1.0\" data-sp-webpartdata=\"&#123;&quot;id&quot;&#58;&quot;e84a4f44-30d2-4962-b203-f8bf42114860&quot;,&quot;instanceId&quot;&#58;&quot;15353e8b-cb55-4794-b871-4cd74abf78b4&quot;,&quot;title&quot;&#58;&quot;Milestone Tracking&quot;,&quot;description&quot;&#58;&quot;A tool used for tracking project milestones&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.0&quot;,&quot;properties&quot;&#58;&#123;&quot;description&quot;&#58;&quot;Milestone Tracking&quot;&#125;&#125;\"><div data-sp-componentid=\"\">e84a4f44-30d2-4962-b203-f8bf42114860</div><div data-sp-htmlproperties=\"\"></div></div></div></div>",
            "BannerImageUrl": null,
            "Description": null,
            "PromotedState": 0,
            "FirstPublishedDate": null,
            "LayoutWebpartsContent": null,
            "OData__AuthorBylineId": null,
            "_AuthorBylineStringId": null,
            "OData__TopicHeader": null,
            "OData__SPSitePageFlags": null,
            "OData__OriginalSourceUrl": null,
            "OData__OriginalSourceSiteId": null,
            "OData__OriginalSourceWebId": null,
            "OData__OriginalSourceListId": null,
            "OData__OriginalSourceItemId": null,
            "ID": 20,
            "Created": "2020-12-11T06:40:05",
            "AuthorId": 11,
            "Modified": "2020-12-11T06:40:05",
            "EditorId": 11,
            "OData__CopySource": null,
            "CheckoutUserId": 11,
            "OData__UIVersionString": "0.1",
            "GUID": "fba3d8ca-790d-4134-a276-7528c32d6b9c"
          },
          "CheckInComment": "",
          "CheckOutType": 0,
          "ContentTag": "{D1408169-EBC1-4A96-B839-95E2D4F439B3},2,2",
          "CustomizedPageStatus": 0,
          "ETag": "\"{D1408169-EBC1-4A96-B839-95E2D4F439B3},2\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "2940",
          "Level": 255,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "lp4blf70.aspx",
          "ServerRelativeUrl": "/SitePages/lp4blf70.aspx",
          "TimeCreated": "2020-12-11T14:40:05Z",
          "TimeLastModified": "2020-12-11T14:40:05Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "d1408169-ebc1-4a96-b839-95e2d4f439b3"
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, addToQuickLaunch: true, title: 'test-single', webUrl: 'https://contoso.sharepoint.com/', webPartData: "{\"id\": \"e84a4f44-30d2-4962-b203-f8bf42114860\", \"instanceId\": \"15353e8b-cb55-4794-b871-4cd74abf78b4\", \"title\": \"Milestone Tracking\", \"description\": \"A tool used for tracking project milestones\", \"serverProcessedContent\": { \"htmlStrings\": {}, \"searchablePlainTexts\": {}, \"imageSources\": {}, \"links\": {} }, \"dataVersion\": \"1.0\", \"properties\": {\"description\": \"Milestone Tracking\"}}" } }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined')
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to create a single-part app page if creating page failed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/CreateAppPage`) > -1) {
        return Promise.reject('Failed to create a single-part app page');
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'failme', webUrl: 'https://contoso.sharepoint.com/', webPartData: JSON.stringify({}) } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Failed to create a single-part app page`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to create a single-part app page if retrieving the created page failed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/CreateAppPage`) > -1) {
        return Promise.resolve({
          value: "SitePages/lp4blf70.aspx"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/SitePages/lp4blf70.aspx')?$expand=ListItemAllFields`) > -1) {
        return Promise.reject('Page not found');
      }
      return Promise.reject('Invalid request');
    });


    command.action(logger, { options: { debug: false, title: 'failme', webUrl: 'https://contoso.sharepoint.com/', webPartData: JSON.stringify({}) } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Page not found`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to create a single-part app page if updating the created page failed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/CreateAppPage`) > -1) {
        return Promise.resolve({
          value: "SitePages/lp4blf70.aspx"
        });
      }
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/UpdateAppPage`) > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/SitePages/lp4blf70.aspx')?$expand=ListItemAllFields`) > -1) {
        return Promise.resolve({
          "ListItemAllFields": {
            "FileSystemObjectType": 0,
            "Id": 20,
            "ServerRedirectedEmbedUri": null,
            "ServerRedirectedEmbedUrl": "",
            "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180031B3A1F0639B70418205BE3DDA9C3044",
            "ComplianceAssetId": null,
            "WikiField": null,
            "Title": null,
            "CanvasContent1": "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;addedFromPersistedData&quot;&#58;&quot;True&quot;,&quot;controlType&quot;&#58;3,&quot;id&quot;&#58;&quot;15353e8b-cb55-4794-b871-4cd74abf78b4&quot;,&quot;webPartId&quot;&#58;&quot;e84a4f44-30d2-4962-b203-f8bf42114860&quot;&#125;\"><div data-sp-webpart=\"\" data-sp-webpartdataversion=\"1.0\" data-sp-webpartdata=\"&#123;&quot;id&quot;&#58;&quot;e84a4f44-30d2-4962-b203-f8bf42114860&quot;,&quot;instanceId&quot;&#58;&quot;15353e8b-cb55-4794-b871-4cd74abf78b4&quot;,&quot;title&quot;&#58;&quot;Milestone Tracking&quot;,&quot;description&quot;&#58;&quot;A tool used for tracking project milestones&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.0&quot;,&quot;properties&quot;&#58;&#123;&quot;description&quot;&#58;&quot;Milestone Tracking&quot;&#125;&#125;\"><div data-sp-componentid=\"\">e84a4f44-30d2-4962-b203-f8bf42114860</div><div data-sp-htmlproperties=\"\"></div></div></div></div>",
            "BannerImageUrl": null,
            "Description": null,
            "PromotedState": 0,
            "FirstPublishedDate": null,
            "LayoutWebpartsContent": null,
            "OData__AuthorBylineId": null,
            "_AuthorBylineStringId": null,
            "OData__TopicHeader": null,
            "OData__SPSitePageFlags": null,
            "OData__OriginalSourceUrl": null,
            "OData__OriginalSourceSiteId": null,
            "OData__OriginalSourceWebId": null,
            "OData__OriginalSourceListId": null,
            "OData__OriginalSourceItemId": null,
            "ID": 20,
            "Created": "2020-12-11T06:40:05",
            "AuthorId": 11,
            "Modified": "2020-12-11T06:40:05",
            "EditorId": 11,
            "OData__CopySource": null,
            "CheckoutUserId": 11,
            "OData__UIVersionString": "0.1",
            "GUID": "fba3d8ca-790d-4134-a276-7528c32d6b9c"
          },
          "CheckInComment": "",
          "CheckOutType": 0,
          "ContentTag": "{D1408169-EBC1-4A96-B839-95E2D4F439B3},2,2",
          "CustomizedPageStatus": 0,
          "ETag": "\"{D1408169-EBC1-4A96-B839-95E2D4F439B3},2\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "2940",
          "Level": 255,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "lp4blf70.aspx",
          "ServerRelativeUrl": "/SitePages/lp4blf70.aspx",
          "TimeCreated": "2020-12-11T14:40:05Z",
          "TimeLastModified": "2020-12-11T14:40:05Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "d1408169-ebc1-4a96-b839-95e2d4f439b3"
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'failme', webUrl: 'https://contoso.sharepoint.com/', webPartData: JSON.stringify({}) } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying title', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--title') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webPartData', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webPartData') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
  it('fails validation if webPartData is not a valid JSON string', () => {
    const actual = command.validate({ options: { title: 'Contoso', webUrl: 'https://contoso', webPartData: 'abc' } });
    assert.notStrictEqual(actual, true);
  });
  it('validation passes on all required options', () => {
    const actual = command.validate({ options: { title: 'Contoso', webPartData: '{}', webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });
});