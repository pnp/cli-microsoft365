import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
import { mockCanvasContent, mockPage } from './page-control-set.mock';
const command: Command = require('./page-header-set');

describe(commands.PAGE_HEADER_SET, () => {
  let log: string[];
  let logger: Logger;
  let data: string;

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

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) > -1) {
        return Promise.resolve({
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        });
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) > -1) {
        return Promise.resolve({ CanvasContent1: mockCanvasContent });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) > -1) {
        data = opts.data;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
    data = '';
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_HEADER_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['imageUrl']);
  });

  it('checks out page if not checked out by the current user', (done) => {
    sinonUtil.restore([request.get, request.post]);
    let checkedOut = false;
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) > -1) {
        return Promise.resolve({
          IsPageCheckedOutToCurrentUser: false,
          Title: 'Page'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`) > -1) {
        checkedOut = true;
        return Promise.resolve(mockPage.ListItemAllFields);
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/SavePageAsDraft`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        pageName: 'home',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter'
      }
    }, () => {
      try {
        assert.strictEqual(checkedOut, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t check out page if not checked out by the current user', (done) => {
    sinonUtil.restore([request.get, request.post]);
    let checkingOut = false;
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) > -1) {
        return Promise.resolve({
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`) > -1) {
        checkingOut = true;
        return Promise.resolve({});
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/SavePageAsDraft`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter'
      }
    }, () => {
      try {
        assert.deepStrictEqual(checkingOut, false);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets page header to default when no type specified', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"topicHeader":""}}]',
      CanvasContent1: '<div>just some test content</div>'
    };

    command.action(logger, { options: { debug: true, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets page header to default when default type specified', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"topicHeader":""}}]',
      CanvasContent1: '<div>just some test content</div>'
    };

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'Default' } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets page header to none when none specified', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"NoImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"topicHeader":""}}]',
      CanvasContent1: '<div>just some test content</div>'
    };

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'None' } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('check when no CanvasContent1 is provided', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"title":"Page","imageSourceType":4,"layoutType":"NoImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"topicHeader":""}}]',
      Title: 'Page',
      AuthorByline: []
    };

    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) > -1) {
        return Promise.resolve({
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        });
      }

      if ((opts.url as string).indexOf(`/_api/site?`) > -1) {
        return Promise.resolve({
          Id: 'c7678ab2-c9dc-454b-b2ee-7fcffb983d4e'
        });
      }

      if ((opts.url as string).indexOf(`/_api/web?`) > -1) {
        return Promise.resolve({
          Id: '0df4d2d2-5ecf-45e9-94f5-c638106bfc65'
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('%2Fsites%2Fteam-a%2Fsiteassets%2Fhero.jpg')?$select=ListId,UniqueId`) > -1) {
        return Promise.resolve({
          ListId: 'e1557527-d333-49f2-9d60-ea8a3003fda8',
          UniqueId: '102f496d-23a2-415f-803a-232b8a6c7613'
        });
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) > -1) {
        return Promise.resolve(null);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'None' } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets page header to custom when custom type specified', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"imageSource":"/sites/team-a/siteassets/hero.jpg"},"links":{},"customMetadata":{"imageSource":{"siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613"}}},"dataVersion":"1.4","properties":{"imageSourceType":2,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"topicHeader":"","authors":[],"altText":"","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613","translateX":42.3837520042758,"translateY":56.4285714285714}}]',
      CanvasContent1: '<div>just some test content</div>'
    };

    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) > -1) {
        return Promise.resolve({
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        });
      }

      if ((opts.url as string).indexOf(`/_api/site?`) > -1) {
        return Promise.resolve({
          Id: 'c7678ab2-c9dc-454b-b2ee-7fcffb983d4e'
        });
      }

      if ((opts.url as string).indexOf(`/_api/web?`) > -1) {
        return Promise.resolve({
          Id: '0df4d2d2-5ecf-45e9-94f5-c638106bfc65'
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('%2Fsites%2Fteam-a%2Fsiteassets%2Fhero.jpg')?$select=ListId,UniqueId`) > -1) {
        return Promise.resolve({
          ListId: 'e1557527-d333-49f2-9d60-ea8a3003fda8',
          UniqueId: '102f496d-23a2-415f-803a-232b8a6c7613'
        });
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) > -1) {
        return Promise.resolve({ CanvasContent1: mockCanvasContent });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'Custom', imageUrl: '/sites/team-a/siteassets/hero.jpg', translateX: 42.3837520042758, translateY: 56.4285714285714 } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets page header to custom when custom type specified (debug)', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"imageSource":"/sites/team-a/siteassets/hero.jpg"},"links":{},"customMetadata":{"imageSource":{"siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613"}}},"dataVersion":"1.4","properties":{"imageSourceType":2,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"topicHeader":"","authors":[],"altText":"","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613","translateX":42.3837520042758,"translateY":56.4285714285714}}]',
      CanvasContent1: '<div>just some test content</div>'
    };

    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) > -1) {
        return Promise.resolve({
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        });
      }

      if ((opts.url as string).indexOf(`/_api/site?`) > -1) {
        return Promise.resolve({
          Id: 'c7678ab2-c9dc-454b-b2ee-7fcffb983d4e'
        });
      }

      if ((opts.url as string).indexOf(`/_api/web?`) > -1) {
        return Promise.resolve({
          Id: '0df4d2d2-5ecf-45e9-94f5-c638106bfc65'
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('%2Fsites%2Fteam-a%2Fsiteassets%2Fhero.jpg')?$select=ListId,UniqueId`) > -1) {
        return Promise.resolve({
          ListId: 'e1557527-d333-49f2-9d60-ea8a3003fda8',
          UniqueId: '102f496d-23a2-415f-803a-232b8a6c7613'
        });
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) > -1) {
        return Promise.resolve({ CanvasContent1: mockCanvasContent });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'Custom', imageUrl: '/sites/team-a/siteassets/hero.jpg', translateX: 42.3837520042758, translateY: 56.4285714285714 } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets image to empty when header set to custom and no image specified', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"imageSource":""},"links":{},"customMetadata":{"imageSource":{"siteId":"","webId":"","listId":"","uniqueId":""}}},"dataVersion":"1.4","properties":{"imageSourceType":2,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"topicHeader":"","authors":[],"altText":"","webId":"","siteId":"","listId":"","uniqueId":"","translateX":0,"translateY":0}}]',
      CanvasContent1: '<div>just some test content</div>'
    };

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'Custom' } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets focus coordinates to 0 0 if none specified', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"imageSource":"/sites/team-a/siteassets/hero.jpg"},"links":{},"customMetadata":{"imageSource":{"siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613"}}},"dataVersion":"1.4","properties":{"imageSourceType":2,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"topicHeader":"","authors":[],"altText":"","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613","translateX":0,"translateY":0}}]',
      CanvasContent1: '<div>just some test content</div>'
    };

    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) > -1) {
        return Promise.resolve({
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        });
      }

      if ((opts.url as string).indexOf(`/_api/site?`) > -1) {
        return Promise.resolve({
          Id: 'c7678ab2-c9dc-454b-b2ee-7fcffb983d4e'
        });
      }

      if ((opts.url as string).indexOf(`/_api/web?`) > -1) {
        return Promise.resolve({
          Id: '0df4d2d2-5ecf-45e9-94f5-c638106bfc65'
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('%2Fsites%2Fteam-a%2Fsiteassets%2Fhero.jpg')?$select=ListId,UniqueId`) > -1) {
        return Promise.resolve({
          ListId: 'e1557527-d333-49f2-9d60-ea8a3003fda8',
          UniqueId: '102f496d-23a2-415f-803a-232b8a6c7613'
        });
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) > -1) {
        return Promise.resolve({ CanvasContent1: mockCanvasContent });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'Custom', imageUrl: '/sites/team-a/siteassets/hero.jpg' } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('centers text when textAlignment set to Center', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Center","showTopicHeader":false,"showPublishDate":false,"topicHeader":""}}]',
      CanvasContent1: '<div>just some test content</div>'
    };

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'Default', textAlignment: 'Center' } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows topicHeader with the specified topicHeader text', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":true,"showPublishDate":false,"topicHeader":"Team Awesome"}}]',
      TopicHeader: 'Team Awesome',
      CanvasContent1: '<div>just some test content</div>'
    };

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'Default', showTopicHeader: true, topicHeader: 'Team Awesome' } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows publish date', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":true,"topicHeader":""}}]',
      CanvasContent1: '<div>just some test content</div>'
    };

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'Default', showPublishDate: true } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows page authors', (done) => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"imageSource":""},"links":{},"customMetadata":{"imageSource":{"siteId":"","webId":"","listId":"","uniqueId":""}}},"dataVersion":"1.4","properties":{"imageSourceType":2,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"topicHeader":"","authors":[],"altText":"","webId":"","siteId":"","listId":"","uniqueId":"","translateX":0,"translateY":0}}]',
      AuthorByline: [ 'Joe Doe', 'Jane Doe' ],
      CanvasContent1: '<div>just some test content</div>'
    };

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'Custom', authors: 'Joe Doe, Jane Doe' } }, () => {
      try {
        assert.strictEqual(JSON.stringify(data), JSON.stringify(mockData));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('automatically appends the .aspx extension', (done) => {
    command.action(logger, { options: { debug: false, pageName: 'page', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when retrieving modern page', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified image doesn\'t exist', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) > -1) {
        return Promise.resolve({
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        });
      }

      if ((opts.url as string).indexOf(`/_api/site?`) > -1) {
        return Promise.resolve({
          Id: 'c7678ab2-c9dc-454b-b2ee-7fcffb983d4e'
        });
      }

      if ((opts.url as string).indexOf(`/_api/web?`) > -1) {
        return Promise.resolve({
          Id: '0df4d2d2-5ecf-45e9-94f5-c638106bfc65'
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('%2Fsites%2Fteam-a%2Fsiteassets%2Fhero.jpg')?$select=ListId,UniqueId`) > -1) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) > -1) {
        return Promise.resolve({ CanvasContent1: mockCanvasContent });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', type: 'Custom', imageUrl: '/sites/team-a/siteassets/hero.jpg', translateX: 42.3837520042758, translateY: 56.4285714285714 } } as any, (err?: any) => {
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
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not an absolute URL', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'http://foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when pageName has no extension', () => {
    const actual = command.validate({ options: { pageName: 'page', webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if type is invalid', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', type: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if type is None', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', type: 'None' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if type is Default', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', type: 'Default' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if type is Custom', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', type: 'Custom' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if translateX is not a valid number', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', translateX: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if translateY is not a valid number', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', translateY: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if layout is invalid', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layout: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if layout is FullWidthImage', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layout: 'FullWidthImage' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout is NoImage', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layout: 'NoImage' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout is ColorBlock', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layout: 'ColorBlock' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout is CutInShape', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layout: 'CutInShape' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if textAlignment is invalid', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', textAlignment: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if textAlignment is Left', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', textAlignment: 'Left' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if textAlignment is Center', () => {
    const actual = command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', textAlignment: 'Center' } });
    assert.strictEqual(actual, true);
  });
});