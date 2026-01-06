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
import commands from '../../commands.js';
import { mockCanvasContentStringified, mockPageJsonCanvasContent } from './page-control-set.mock.js';
import command from './page-header-set.js';

describe(commands.PAGE_HEADER_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

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
    assert.strictEqual(command.name, commands.PAGE_HEADER_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['imageUrl']);
  });

  it('checks out page if not checked out by the current user', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: false,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockPageJsonCanvasContent.ListItemAllFields;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) {
        return mockPageJsonCanvasContent.ListItemAllFields;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        pageName: 'page',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter'
      }
    });
    assert.strictEqual(postStub.firstCall.args[0].url, `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`);
  });

  it('doesn\'t check out page if page is checked out by the current user', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockPageJsonCanvasContent.ListItemAllFields;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter'
      }
    });
    assert.notStrictEqual(postStub.firstCall.args[0].url, `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`);
  });

  it('sets page header to default when no type specified', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":""}}]',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('sets page header to default when default type specified', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":""}}]',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'Default' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('sets page header to none when none specified', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"NoImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":""}}]',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'None' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('sets page header to custom when custom type specified', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"imageSource":"/sites/newsletter/siteassets/hero.jpg"},"links":{},"customMetadata":{"imageSource":{"siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613"}}},"dataVersion":"1.4","properties":{"imageSourceType":2,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":"","authors":[],"altText":"","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613","translateX":42.3837520042758,"translateY":56.4285714285714}}]',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/site?$select=Id`) {
        return {
          Id: 'c7678ab2-c9dc-454b-b2ee-7fcffb983d4e'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/web?$select=Id`) {
        return {
          Id: '0df4d2d2-5ecf-45e9-94f5-c638106bfc65'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/web/GetFileByServerRelativePath(DecodedUrl='%2Fsites%2Fnewsletter%2Fsiteassets%2Fhero.jpg')?$select=ListId,UniqueId`) {
        return {
          ListId: 'e1557527-d333-49f2-9d60-ea8a3003fda8',
          UniqueId: '102f496d-23a2-415f-803a-232b8a6c7613'
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'Custom', imageUrl: '/sites/newsletter/siteassets/hero.jpg', translateX: 42.3837520042758, translateY: 56.4285714285714 } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('sets page header to custom when custom type specified (debug)', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"imageSource":"/sites/newsletter/siteassets/hero.jpg"},"links":{},"customMetadata":{"imageSource":{"siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613"}}},"dataVersion":"1.4","properties":{"imageSourceType":2,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":"","authors":[],"altText":"","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613","translateX":42.3837520042758,"translateY":56.4285714285714}}]',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/site?$select=Id`) {
        return {
          Id: 'c7678ab2-c9dc-454b-b2ee-7fcffb983d4e'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/web?$select=Id`) {
        return {
          Id: '0df4d2d2-5ecf-45e9-94f5-c638106bfc65'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/web/GetFileByServerRelativePath(DecodedUrl='%2Fsites%2Fnewsletter%2Fsiteassets%2Fhero.jpg')?$select=ListId,UniqueId`) {
        return {
          ListId: 'e1557527-d333-49f2-9d60-ea8a3003fda8',
          UniqueId: '102f496d-23a2-415f-803a-232b8a6c7613'
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'Custom', imageUrl: '/sites/newsletter/siteassets/hero.jpg', translateX: 42.3837520042758, translateY: 56.4285714285714 } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('sets image to empty when header set to custom and no image specified', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"imageSource":""},"links":{},"customMetadata":{"imageSource":{"siteId":"","webId":"","listId":"","uniqueId":""}}},"dataVersion":"1.4","properties":{"imageSourceType":2,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":"","authors":[],"altText":"","webId":"","siteId":"","listId":"","uniqueId":"","translateX":0,"translateY":0}}]',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'Custom' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('sets focus coordinates to 0 0 if none specified', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"imageSource":"/sites/newsletter/siteassets/hero.jpg"},"links":{},"customMetadata":{"imageSource":{"siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613"}}},"dataVersion":"1.4","properties":{"imageSourceType":2,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":"","authors":[],"altText":"","webId":"0df4d2d2-5ecf-45e9-94f5-c638106bfc65","siteId":"c7678ab2-c9dc-454b-b2ee-7fcffb983d4e","listId":"e1557527-d333-49f2-9d60-ea8a3003fda8","uniqueId":"102f496d-23a2-415f-803a-232b8a6c7613","translateX":0,"translateY":0}}]',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/site?$select=Id`) {
        return {
          Id: 'c7678ab2-c9dc-454b-b2ee-7fcffb983d4e'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/web?$select=Id`) {
        return {
          Id: '0df4d2d2-5ecf-45e9-94f5-c638106bfc65'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/web/GetFileByServerRelativePath(DecodedUrl='%2Fsites%2Fnewsletter%2Fsiteassets%2Fhero.jpg')?$select=ListId,UniqueId`) {
        return {
          ListId: 'e1557527-d333-49f2-9d60-ea8a3003fda8',
          UniqueId: '102f496d-23a2-415f-803a-232b8a6c7613'
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'Custom', imageUrl: '/sites/newsletter/siteassets/hero.jpg' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('centers text when textAlignment set to Center', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Center","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":""}}]',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'Default', textAlignment: 'Center' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('shows topicHeader with the specified topicHeader text', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":true,"showPublishDate":false,"showTimeToRead":false,"topicHeader":"Team Awesome"}}]',
      TopicHeader: 'Team Awesome',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'Default', showTopicHeader: true, topicHeader: 'Team Awesome' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('shows publish date', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":true,"showTimeToRead":false,"topicHeader":""}}]',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'Default', showPublishDate: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('shows correctly shows time to read', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":true,"topicHeader":""}}]',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'Default', showTimeToRead: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('shows page authors', async () => {
    const mockData = {
      LayoutWebpartsContent: '[{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"imageSource":""},"links":{},"customMetadata":{"imageSource":{"siteId":"","webId":"","listId":"","uniqueId":""}}},"dataVersion":"1.4","properties":{"imageSourceType":2,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":"","authors":[],"altText":"","webId":"","siteId":"","listId":"","uniqueId":"","translateX":0,"translateY":0}}]',
      AuthorByline: ['Joe Doe', 'Jane Doe'],
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'Custom', authors: 'Joe Doe, Jane Doe' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, mockData);
  });

  it('automatically appends the .aspx extension', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockPageJsonCanvasContent.ListItemAllFields;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { pageName: 'page', webUrl: 'https://contoso.sharepoint.com/sites/newsletter' } } as any);
    assert.deepStrictEqual(getStub.firstCall.args[0].url, `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`);
  });

  it('sets page header to default when no type specified without any header', async () => {
    const mockData = {
      LayoutWebpartsContent: '[]',
      CanvasContent1: mockCanvasContentStringified
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      LayoutWebpartsContent: '[]',
      CanvasContent1: '[{"controlType":3,"displayMode":2,"id":"ede2ee65-157d-4523-b4ed-87b9b64374a6","position":{"zoneIndex":1,"sectionIndex":2,"sectionFactor":12,"layoutIndex":1,"controlIndex":1},"webPartId":"ede2ee65-157d-4523-b4ed-87b9b64374a6","emphasis":{},"addedFromPersistedData":true,"reservedHeight":600,"reservedWidth":969,"webPartData":{"id":"ede2ee65-157d-4523-b4ed-87b9b64374a6","instanceId":"dcd01c36-24f9-42e5-8e03-76e4af572468","title":"valo-markdown","description":"Use markdown to add text, tables, links, and images to your page.","serverProcessedContent":{"htmlStrings":{"html":"<h2 id=\\"this-is-just-a-test\\">This is just a test</h2><p>Test for playbook 123</p><div class=\\"react-codemirror2\\"><div class=\\"CodeMirror cm-s-monokai CodeMirror-wrap\\"><div class=\\"CodeMirror-vscrollbar\\" tabindex=\\"-1\\" style=\\"bottom:0px;\\"><div style=\\"min-width:1px;height:0px;\\"></div></div><div class=\\"CodeMirror-hscrollbar\\" tabindex=\\"-1\\"><div style=\\"height:100%;min-height:1px;width:0px;\\"></div></div><div class=\\"CodeMirror-scrollbar-filler\\"></div><div class=\\"CodeMirror-gutter-filler\\"></div><div class=\\"CodeMirror-scroll\\" tabindex=\\"-1\\"><div class=\\"CodeMirror-sizer\\" style=\\"margin-left:0px;margin-bottom:-8px;border-right-width:22px;min-height:290px;padding-right:0px;padding-bottom:0px;\\"><div style=\\"position:relative;top:0px;\\"><div class=\\"CodeMirror-lines\\" role=\\"presentation\\"><div role=\\"presentation\\" style=\\"position:relative;outline:none;\\"><div class=\\"CodeMirror-measure\\"></div><div class=\\"CodeMirror-measure\\"></div><div style=\\"position:relative;z-index:1;\\"></div><div class=\\"CodeMirror-cursors\\"><div class=\\"CodeMirror-cursor\\" style=\\"left:55.0089px;top:260.571px;height:21.7143px;\\">&nbsp;</div></div><div class=\\"CodeMirror-code\\" role=\\"presentation\\"><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> { <span class=\\"cm-def\\">exec</span> } <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">child_process</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">fs</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">fs</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">path</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">path</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">parseMarkdown</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">frontmatter</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span>&nbsp;</span></span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">valoWpTitle</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-string-2\\">`valo-markdown`</span>;</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">siteUrl</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-string-2\\">`https://contoso.sharepoint.com/sites/StaticPages`</span>;</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span>&nbsp;</span></span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\">(() <span class=\\"cm-operator\\">=&gt;</span> {</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"> &nbsp;<span class=\\"cm-keyword\\">if</span> (<span class=\\"cm-atom\\">true</span>) {</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span>&nbsp;</span></span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\">  }</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\">})();</span></pre></div></div></div></div></div><div style=\\"position:absolute;height:22px;width:1px;border-bottom:0px solid transparent;top:290px;\\"></div><div class=\\"CodeMirror-gutters\\" style=\\"display:none;height:312px;\\"></div></div></div></div><p>!!! Warning\\n    This is a warning of mkdocs</p><p>&lt;p style=\\"color:red\\"&gt;This is a paragraph&lt;/p&gt;</p>"},"searchablePlainTexts":{"code":"\\n# This is just a test\\n\\nTest for playbook 123\\n\\n```typescript\\nconst { exec } = require(child_process);\\nconst fs = require(fs);\\nconst path = require(path);\\nconst parseMarkdown = require(frontmatter);\\n\\nconst valoWpTitle = `valo-markdown`;\\nconst siteUrl = `https://contoso.sharepoint.com/sites/StaticPages`;\\n\\n(() => {\\n  if (true) {\\n\\n  }\\n})();\\n```\\n\\n\\n!!! Warning\\n    This is a warning of mkdocs\\n\\n\\n<p style=\\"color:red\\">This is a paragraph</p>"},"imageSources":{},"links":{}},"dataVersion":"2.0","properties":{"displayPreview":true,"lineWrapping":true,"miniMap":{"enabled":false},"previewState":"Show","theme":"Monokai"}}},{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","controlType":3,"displayMode":2,"emphasis":{},"position":{"zoneIndex":1,"sectionFactor":0,"layoutIndex":1,"controlIndex":1,"sectionIndex":1},"webPartId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","webPartData":{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"title":"","imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":""}}}]'
    });
  });

  it('sets page header to default when no type specified with Banner Webpart in first section', async () => {
    const mockData = {
      LayoutWebpartsContent: '[]',
      CanvasContent1: '[{"controlType":3,"displayMode":2,"id":"ede2ee65-157d-4523-b4ed-87b9b64374a6","position":{"zoneIndex":1,"sectionIndex":2,"sectionFactor":12,"layoutIndex":1,"controlIndex":1},"webPartId":"ede2ee65-157d-4523-b4ed-87b9b64374a6","emphasis":{},"addedFromPersistedData":true,"reservedHeight":600,"reservedWidth":969,"webPartData":{"id":"ede2ee65-157d-4523-b4ed-87b9b64374a6","instanceId":"dcd01c36-24f9-42e5-8e03-76e4af572468","title":"valo-markdown","description":"Use markdown to add text, tables, links, and images to your page.","serverProcessedContent":{"htmlStrings":{"html":"<h2 id=\\"this-is-just-a-test\\">This is just a test</h2><p>Test for playbook 123</p><div class=\\"react-codemirror2\\"><div class=\\"CodeMirror cm-s-monokai CodeMirror-wrap\\"><div class=\\"CodeMirror-vscrollbar\\" tabindex=\\"-1\\" style=\\"bottom:0px;\\"><div style=\\"min-width:1px;height:0px;\\"></div></div><div class=\\"CodeMirror-hscrollbar\\" tabindex=\\"-1\\"><div style=\\"height:100%;min-height:1px;width:0px;\\"></div></div><div class=\\"CodeMirror-scrollbar-filler\\"></div><div class=\\"CodeMirror-gutter-filler\\"></div><div class=\\"CodeMirror-scroll\\" tabindex=\\"-1\\"><div class=\\"CodeMirror-sizer\\" style=\\"margin-left:0px;margin-bottom:-8px;border-right-width:22px;min-height:290px;padding-right:0px;padding-bottom:0px;\\"><div style=\\"position:relative;top:0px;\\"><div class=\\"CodeMirror-lines\\" role=\\"presentation\\"><div role=\\"presentation\\" style=\\"position:relative;outline:none;\\"><div class=\\"CodeMirror-measure\\"></div><div class=\\"CodeMirror-measure\\"></div><div style=\\"position:relative;z-index:1;\\"></div><div class=\\"CodeMirror-cursors\\"><div class=\\"CodeMirror-cursor\\" style=\\"left:55.0089px;top:260.571px;height:21.7143px;\\">&nbsp;</div></div><div class=\\"CodeMirror-code\\" role=\\"presentation\\"><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> { <span class=\\"cm-def\\">exec</span> } <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">child_process</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">fs</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">fs</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">path</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">path</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">parseMarkdown</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">frontmatter</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span>&nbsp;</span></span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">valoWpTitle</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-string-2\\">`valo-markdown`</span>;</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">siteUrl</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-string-2\\">`https://contoso.sharepoint.com/sites/StaticPages`</span>;</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span>&nbsp;</span></span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\">(() <span class=\\"cm-operator\\">=&gt;</span> {</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"> &nbsp;<span class=\\"cm-keyword\\">if</span> (<span class=\\"cm-atom\\">true</span>) {</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span>&nbsp;</span></span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\">  }</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\">})();</span></pre></div></div></div></div></div><div style=\\"position:absolute;height:22px;width:1px;border-bottom:0px solid transparent;top:290px;\\"></div><div class=\\"CodeMirror-gutters\\" style=\\"display:none;height:312px;\\"></div></div></div></div><p>!!! Warning\\n    This is a warning of mkdocs</p><p>&lt;p style=\\"color:red\\"&gt;This is a paragraph&lt;/p&gt;</p>"},"searchablePlainTexts":{"code":"\\n# This is just a test\\n\\nTest for playbook 123\\n\\n```typescript\\nconst { exec } = require(child_process);\\nconst fs = require(fs);\\nconst path = require(path);\\nconst parseMarkdown = require(frontmatter);\\n\\nconst valoWpTitle = `valo-markdown`;\\nconst siteUrl = `https://contoso.sharepoint.com/sites/StaticPages`;\\n\\n(() => {\\n  if (true) {\\n\\n  }\\n})();\\n```\\n\\n\\n!!! Warning\\n    This is a warning of mkdocs\\n\\n\\n<p style=\\"color:red\\">This is a paragraph</p>"},"imageSources":{},"links":{}},"dataVersion":"2.0","properties":{"displayPreview":true,"lineWrapping":true,"miniMap":{"enabled":false},"previewState":"Show","theme":"Monokai"}}},{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","controlType":3,"displayMode":2,"emphasis":{},"position":{"zoneIndex":1,"sectionFactor":0,"layoutIndex":1,"controlIndex":1,"sectionIndex":1},"webPartId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","webPartData":{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"title":"","imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":""}}}]'
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockData;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      LayoutWebpartsContent: '[]',
      CanvasContent1: '[{"controlType":3,"displayMode":2,"id":"ede2ee65-157d-4523-b4ed-87b9b64374a6","position":{"zoneIndex":1,"sectionIndex":2,"sectionFactor":12,"layoutIndex":1,"controlIndex":1},"webPartId":"ede2ee65-157d-4523-b4ed-87b9b64374a6","emphasis":{},"addedFromPersistedData":true,"reservedHeight":600,"reservedWidth":969,"webPartData":{"id":"ede2ee65-157d-4523-b4ed-87b9b64374a6","instanceId":"dcd01c36-24f9-42e5-8e03-76e4af572468","title":"valo-markdown","description":"Use markdown to add text, tables, links, and images to your page.","serverProcessedContent":{"htmlStrings":{"html":"<h2 id=\\"this-is-just-a-test\\">This is just a test</h2><p>Test for playbook 123</p><div class=\\"react-codemirror2\\"><div class=\\"CodeMirror cm-s-monokai CodeMirror-wrap\\"><div class=\\"CodeMirror-vscrollbar\\" tabindex=\\"-1\\" style=\\"bottom:0px;\\"><div style=\\"min-width:1px;height:0px;\\"></div></div><div class=\\"CodeMirror-hscrollbar\\" tabindex=\\"-1\\"><div style=\\"height:100%;min-height:1px;width:0px;\\"></div></div><div class=\\"CodeMirror-scrollbar-filler\\"></div><div class=\\"CodeMirror-gutter-filler\\"></div><div class=\\"CodeMirror-scroll\\" tabindex=\\"-1\\"><div class=\\"CodeMirror-sizer\\" style=\\"margin-left:0px;margin-bottom:-8px;border-right-width:22px;min-height:290px;padding-right:0px;padding-bottom:0px;\\"><div style=\\"position:relative;top:0px;\\"><div class=\\"CodeMirror-lines\\" role=\\"presentation\\"><div role=\\"presentation\\" style=\\"position:relative;outline:none;\\"><div class=\\"CodeMirror-measure\\"></div><div class=\\"CodeMirror-measure\\"></div><div style=\\"position:relative;z-index:1;\\"></div><div class=\\"CodeMirror-cursors\\"><div class=\\"CodeMirror-cursor\\" style=\\"left:55.0089px;top:260.571px;height:21.7143px;\\">&nbsp;</div></div><div class=\\"CodeMirror-code\\" role=\\"presentation\\"><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> { <span class=\\"cm-def\\">exec</span> } <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">child_process</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">fs</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">fs</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">path</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">path</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">parseMarkdown</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-variable\\">require</span>(<span class=\\"cm-variable\\">frontmatter</span>);</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span>&nbsp;</span></span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">valoWpTitle</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-string-2\\">`valo-markdown`</span>;</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span class=\\"cm-keyword\\">const</span> <span class=\\"cm-def\\">siteUrl</span> <span class=\\"cm-operator\\">=</span> <span class=\\"cm-string-2\\">`https://contoso.sharepoint.com/sites/StaticPages`</span>;</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span>&nbsp;</span></span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\">(() <span class=\\"cm-operator\\">=&gt;</span> {</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"> &nbsp;<span class=\\"cm-keyword\\">if</span> (<span class=\\"cm-atom\\">true</span>) {</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\"><span>&nbsp;</span></span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\">  }</span></pre><pre class=\\" CodeMirror-line \\" role=\\"presentation\\"><span role=\\"presentation\\" style=\\"padding-right:0.1px;\\">})();</span></pre></div></div></div></div></div><div style=\\"position:absolute;height:22px;width:1px;border-bottom:0px solid transparent;top:290px;\\"></div><div class=\\"CodeMirror-gutters\\" style=\\"display:none;height:312px;\\"></div></div></div></div><p>!!! Warning\\n    This is a warning of mkdocs</p><p>&lt;p style=\\"color:red\\"&gt;This is a paragraph&lt;/p&gt;</p>"},"searchablePlainTexts":{"code":"\\n# This is just a test\\n\\nTest for playbook 123\\n\\n```typescript\\nconst { exec } = require(child_process);\\nconst fs = require(fs);\\nconst path = require(path);\\nconst parseMarkdown = require(frontmatter);\\n\\nconst valoWpTitle = `valo-markdown`;\\nconst siteUrl = `https://contoso.sharepoint.com/sites/StaticPages`;\\n\\n(() => {\\n  if (true) {\\n\\n  }\\n})();\\n```\\n\\n\\n!!! Warning\\n    This is a warning of mkdocs\\n\\n\\n<p style=\\"color:red\\">This is a paragraph</p>"},"imageSources":{},"links":{}},"dataVersion":"2.0","properties":{"displayPreview":true,"lineWrapping":true,"miniMap":{"enabled":false},"previewState":"Show","theme":"Monokai"}}},{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","controlType":3,"displayMode":2,"emphasis":{},"position":{"zoneIndex":1,"sectionFactor":0,"layoutIndex":1,"controlIndex":1,"sectionIndex":1},"webPartId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","webPartData":{"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","instanceId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","title":"Title Region","description":"Title Region Description","serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.4","properties":{"title":"","imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"showTimeToRead":false,"topicHeader":""}}}]'
    });
  });

  it('correctly handles OData error when retrieving modern page', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    await assert.rejects(command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter' } }),
      new CommandError('An error has occurred'));
  });

  it('correctly handles page not found error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw {
        status: 404,
        error: { 'odata.error': { message: { value: "Exception of type 'Microsoft.SharePoint.Client.ClientServiceException' was thrown." } } }
      };
    });

    await assert.rejects(command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter' } }),
      new CommandError(`The specified page 'page.aspx' does not exist.`));
  });

  it('correctly handles error when the specified image doesn\'t exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=IsPageCheckedOutToCurrentUser,Title`) > -1) {
        return {
          IsPageCheckedOutToCurrentUser: true,
          Title: 'Page'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/site?$select=Id`) {
        return {
          Id: 'c7678ab2-c9dc-454b-b2ee-7fcffb983d4e'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/web?$select=Id`) {
        return {
          Id: '0df4d2d2-5ecf-45e9-94f5-c638106bfc65'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/web/GetFileByServerRelativePath(DecodedUrl='%2Fsites%2Fnewsletter%2Fsiteassets%2Fhero.jpg')?$select=ListId,UniqueId`) {
        throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$expand=ListItemAllFields`) {
        return mockPageJsonCanvasContent.ListItemAllFields;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/newsletter', type: 'Custom', imageUrl: '/sites/newsletter/siteassets/hero.jpg', translateX: 42.3837520042758, translateY: 56.4285714285714 } } as any),
      new CommandError('An error has occurred'));
  });

  it('fails validation if webUrl is not an absolute URL', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'http://foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when pageName has no extension', async () => {
    const actual = await command.validate({ options: { pageName: 'page', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if type is invalid', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', type: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if type is None', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', type: 'None' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if type is Default', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', type: 'Default' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if type is Custom', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', type: 'Custom' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if translateX is not a valid number', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', translateX: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if translateY is not a valid number', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', translateY: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if layout is invalid', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layout: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if layout is FullWidthImage', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layout: 'FullWidthImage' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout is NoImage', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layout: 'NoImage' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout is ColorBlock', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layout: 'ColorBlock' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout is CutInShape', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layout: 'CutInShape' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if textAlignment is invalid', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', textAlignment: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if textAlignment is Left', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', textAlignment: 'Left' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if textAlignment is Center', async () => {
    const actual = await command.validate({ options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', textAlignment: 'Center' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
