import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./page-text-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.PAGE_TEXT_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon
      .stub(command as any, 'getRequestDigest')
      .callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_TEXT_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds text to an empty modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`) > -1) {
        return Promise.resolve({
          ListItemAllFields: {
            CommentsDisabled: false,
            FileSystemObjectType: 0,
            Id: 1,
            ServerRedirectedEmbedUri: null,
            ServerRedirectedEmbedUrl: '',
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
            FileLeafRef: 'Home.aspx',
            ComplianceAssetId: null,
            WikiField: null,
            Title: 'Home',
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: 'Home',
            CanvasContent1:
              '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true&#125;&#125;"></div></div>',
            BannerImageUrl: {
              Description: '/_layouts/15/images/sitepagethumbnail.png',
              Url: 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'
            },
            Description: 'Lorem ipsum Dolor samet Lorem ipsum',
            PromotedState: null,
            FirstPublishedDate: null,
            LayoutWebpartsContent: null,
            AuthorsId: null,
            AuthorsStringId: null,
            OriginalSourceUrl: null,
            ID: 1,
            Created: '2018-01-20T09:54:41',
            AuthorId: 1073741823,
            Modified: '2018-04-12T12:42:47',
            EditorId: 12,
            OData__CopySource: null,
            CheckoutUserId: null,
            OData__UIVersionString: '7.0',
            GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
          },
          CheckInComment: '',
          CheckOutType: 2,
          ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
          CustomizedPageStatus: 1,
          ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
          Exists: true,
          IrmEnabled: false,
          Length: '805',
          Level: 1,
          LinkingUri: null,
          LinkingUrl: '',
          MajorVersion: 7,
          MinorVersion: 0,
          Name: 'home.aspx',
          ServerRelativeUrl: '/sites/team-a/SitePages/home.aspx',
          TimeCreated: '2018-01-20T08:54:41Z',
          TimeLastModified: '2018-04-12T10:42:46Z',
          Title: 'Home',
          UIVersion: 3584,
          UIVersionLabel: '7.0',
          UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body).indexOf(`&quot;,&quot;position&quot;&#58;&#123;&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;\\"><div data-sp-rte=\\"\\"><p>Hello world</p></div></div></div>"}`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          text: 'Hello world'
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.notCalled);
          done();
        }
        catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds text to an empty modern page (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`) > -1) {
        return Promise.resolve({
          ListItemAllFields: {
            CommentsDisabled: false,
            FileSystemObjectType: 0,
            Id: 1,
            ServerRedirectedEmbedUri: null,
            ServerRedirectedEmbedUrl: '',
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
            FileLeafRef: 'Home.aspx',
            ComplianceAssetId: null,
            WikiField: null,
            Title: 'Page',
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: 'Article',
            CanvasContent1:
              '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true&#125;&#125;"></div></div>',
            BannerImageUrl: {
              Description: '/_layouts/15/images/sitepagethumbnail.png',
              Url: 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'
            },
            Description: 'Lorem ipsum Dolor samet Lorem ipsum',
            PromotedState: null,
            FirstPublishedDate: null,
            LayoutWebpartsContent: null,
            AuthorsId: null,
            AuthorsStringId: null,
            OriginalSourceUrl: null,
            ID: 1,
            Created: '2018-01-20T09:54:41',
            AuthorId: 1073741823,
            Modified: '2018-04-12T12:42:47',
            EditorId: 12,
            OData__CopySource: null,
            CheckoutUserId: null,
            OData__UIVersionString: '7.0',
            GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
          },
          CheckInComment: '',
          CheckOutType: 2,
          ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
          CustomizedPageStatus: 1,
          ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
          Exists: true,
          IrmEnabled: false,
          Length: '805',
          Level: 1,
          LinkingUri: null,
          LinkingUrl: '',
          MajorVersion: 7,
          MinorVersion: 0,
          Name: 'page.aspx',
          ServerRelativeUrl: '/sites/team-a/SitePages/page.aspx',
          TimeCreated: '2018-01-20T08:54:41Z',
          TimeLastModified: '2018-04-12T10:42:46Z',
          Title: 'Home',
          UIVersion: 3584,
          UIVersionLabel: '7.0',
          UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body).indexOf(`&quot;,&quot;position&quot;&#58;&#123;&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;\\"><div data-sp-rte=\\"\\"><p>Hello world</p></div></div></div>"}`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: true,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          text: 'Hello world'
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
          done();
        }
        catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds text to an empty modern page on root of tenant (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/_api/web/getfilebyserverrelativeurl('/SitePages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`) > -1) {
        return Promise.resolve({
          ListItemAllFields: {
            CommentsDisabled: false,
            FileSystemObjectType: 0,
            Id: 1,
            ServerRedirectedEmbedUri: null,
            ServerRedirectedEmbedUrl: '',
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
            FileLeafRef: 'Home.aspx',
            ComplianceAssetId: null,
            WikiField: null,
            Title: 'Page',
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: 'Article',
            CanvasContent1:
              '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true&#125;&#125;"></div></div>',
            BannerImageUrl: {
              Description: '/_layouts/15/images/sitepagethumbnail.png',
              Url: 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'
            },
            Description: 'Lorem ipsum Dolor samet Lorem ipsum',
            PromotedState: null,
            FirstPublishedDate: null,
            LayoutWebpartsContent: null,
            AuthorsId: null,
            AuthorsStringId: null,
            OriginalSourceUrl: null,
            ID: 1,
            Created: '2018-01-20T09:54:41',
            AuthorId: 1073741823,
            Modified: '2018-04-12T12:42:47',
            EditorId: 12,
            OData__CopySource: null,
            CheckoutUserId: null,
            OData__UIVersionString: '7.0',
            GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
          },
          CheckInComment: '',
          CheckOutType: 2,
          ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
          CustomizedPageStatus: 1,
          ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
          Exists: true,
          IrmEnabled: false,
          Length: '805',
          Level: 1,
          LinkingUri: null,
          LinkingUrl: '',
          MajorVersion: 7,
          MinorVersion: 0,
          Name: 'page.aspx',
          ServerRelativeUrl: '/SitePages/page.aspx',
          TimeCreated: '2018-01-20T08:54:41Z',
          TimeLastModified: '2018-04-12T10:42:46Z',
          Title: 'Home',
          UIVersion: 3584,
          UIVersionLabel: '7.0',
          UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/_api/web/getfilebyserverrelativeurl('/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body).indexOf(`&quot;,&quot;position&quot;&#58;&#123;&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;\\"><div data-sp-rte=\\"\\"><p>Hello world</p></div></div></div>"}`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: true,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com',
          text: 'Hello world'
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
          done();
        }
        catch (e) {
          done(e);
        }
      }
    );
  });

  it('appends text to a modern page which already had some text', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (
        (opts.url as string).indexOf(
          `/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`
        ) > -1
      ) {
        return Promise.resolve({
          ListItemAllFields: {
            CommentsDisabled: false,
            FileSystemObjectType: 0,
            Id: 1,
            ServerRedirectedEmbedUri: null,
            ServerRedirectedEmbedUrl: '',
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
            FileLeafRef: 'Home.aspx',
            ComplianceAssetId: null,
            WikiField: null,
            Title: 'Home',
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: 'Home',
            CanvasContent1:
              '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;4,&quot;editorType&quot;&#58;&quot;CKEditor&quot;,&quot;id&quot;&#58;&quot;e278967c-6f89-4601-a30b-f132dc48d55b&quot;,&quot;position&quot;&#58;&#123;&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;"><div data-sp-rte=""><p>Hello world</p></div></div></div>',
            BannerImageUrl: {
              Description: '/_layouts/15/images/sitepagethumbnail.png',
              Url: 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'
            },
            Description: 'Lorem ipsum Dolor samet Lorem ipsum',
            PromotedState: null,
            FirstPublishedDate: null,
            LayoutWebpartsContent: null,
            AuthorsId: null,
            AuthorsStringId: null,
            OriginalSourceUrl: null,
            ID: 1,
            Created: '2018-01-20T09:54:41',
            AuthorId: 1073741823,
            Modified: '2018-04-12T12:42:47',
            EditorId: 12,
            OData__CopySource: null,
            CheckoutUserId: null,
            OData__UIVersionString: '7.0',
            GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
          },
          CheckInComment: '',
          CheckOutType: 2,
          ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
          CustomizedPageStatus: 1,
          ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
          Exists: true,
          IrmEnabled: false,
          Length: '805',
          Level: 1,
          LinkingUri: null,
          LinkingUrl: '',
          MajorVersion: 7,
          MinorVersion: 0,
          Name: 'home.aspx',
          ServerRelativeUrl: '/sites/team-a/SitePages/home.aspx',
          TimeCreated: '2018-01-20T08:54:41Z',
          TimeLastModified: '2018-04-12T10:42:46Z',
          Title: 'Home',
          UIVersion: 3584,
          UIVersionLabel: '7.0',
          UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body).endsWith(`&quot;position&quot;&#58;&#123;&quot;controlIndex&quot;&#58;2,&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;"><div data-sp-rte=""><p>Hello world 2</p></div></div></div>`)) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          text: 'Hello world'
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.notCalled);
          done();
        }
        catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds text in the specified order to a modern page which already had some text', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (
        (opts.url as string).indexOf(
          `/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`
        ) > -1
      ) {
        return Promise.resolve({
          ListItemAllFields: {
            CommentsDisabled: false,
            FileSystemObjectType: 0,
            Id: 1,
            ServerRedirectedEmbedUri: null,
            ServerRedirectedEmbedUrl: '',
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
            FileLeafRef: 'Home.aspx',
            ComplianceAssetId: null,
            WikiField: null,
            Title: 'Home',
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: 'Home',
            CanvasContent1:
              '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;4,&quot;editorType&quot;&#58;&quot;CKEditor&quot;,&quot;id&quot;&#58;&quot;e278967c-6f89-4601-a30b-f132dc48d55b&quot;,&quot;position&quot;&#58;&#123;&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;"><div data-sp-rte=""><p>Hello world</p></div></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;4,&quot;editorType&quot;&#58;&quot;CKEditor&quot;,&quot;id&quot;&#58;&quot;cc988078-be29-4999-a5e2-4aa0f9a04ab4&quot;,&quot;position&quot;&#58;&#123;&quot;controlIndex&quot;&#58;2,&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;"><div data-sp-rte=""><p>Hello world 2</p></div></div></div>',
            BannerImageUrl: {
              Description: '/_layouts/15/images/sitepagethumbnail.png',
              Url: 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'
            },
            Description: 'Lorem ipsum Dolor samet Lorem ipsum',
            PromotedState: null,
            FirstPublishedDate: null,
            LayoutWebpartsContent: null,
            AuthorsId: null,
            AuthorsStringId: null,
            OriginalSourceUrl: null,
            ID: 1,
            Created: '2018-01-20T09:54:41',
            AuthorId: 1073741823,
            Modified: '2018-04-12T12:42:47',
            EditorId: 12,
            OData__CopySource: null,
            CheckoutUserId: null,
            OData__UIVersionString: '7.0',
            GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
          },
          CheckInComment: '',
          CheckOutType: 2,
          ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
          CustomizedPageStatus: 1,
          ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
          Exists: true,
          IrmEnabled: false,
          Length: '805',
          Level: 1,
          LinkingUri: null,
          LinkingUrl: '',
          MajorVersion: 7,
          MinorVersion: 0,
          Name: 'home.aspx',
          ServerRelativeUrl: '/sites/team-a/SitePages/home.aspx',
          TimeCreated: '2018-01-20T08:54:41Z',
          TimeLastModified: '2018-04-12T10:42:46Z',
          Title: 'Home',
          UIVersion: 3584,
          UIVersionLabel: '7.0',
          UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body).endsWith(`position&quot;&#58;&#123;&quot;controlIndex&quot;&#58;3,&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;"><div data-sp-rte=""><p>Hello world 2</p></div></div></div>`)) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          text: 'Hello world 1.1',
          order: 2
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.notCalled);
          done();
        }
        catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds text to a modern page without specifying the page file extension', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (
        (opts.url as string).indexOf(
          `/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`
        ) > -1
      ) {
        return Promise.resolve({
          ListItemAllFields: {
            CommentsDisabled: false,
            FileSystemObjectType: 0,
            Id: 1,
            ServerRedirectedEmbedUri: null,
            ServerRedirectedEmbedUrl: '',
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
            FileLeafRef: 'Home.aspx',
            ComplianceAssetId: null,
            WikiField: null,
            Title: 'Home',
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: 'Home',
            CanvasContent1:
              '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true&#125;&#125;"></div></div>',
            BannerImageUrl: {
              Description: '/_layouts/15/images/sitepagethumbnail.png',
              Url: 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'
            },
            Description: 'Lorem ipsum Dolor samet Lorem ipsum',
            PromotedState: null,
            FirstPublishedDate: null,
            LayoutWebpartsContent: null,
            AuthorsId: null,
            AuthorsStringId: null,
            OriginalSourceUrl: null,
            ID: 1,
            Created: '2018-01-20T09:54:41',
            AuthorId: 1073741823,
            Modified: '2018-04-12T12:42:47',
            EditorId: 12,
            OData__CopySource: null,
            CheckoutUserId: null,
            OData__UIVersionString: '7.0',
            GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
          },
          CheckInComment: '',
          CheckOutType: 2,
          ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
          CustomizedPageStatus: 1,
          ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
          Exists: true,
          IrmEnabled: false,
          Length: '805',
          Level: 1,
          LinkingUri: null,
          LinkingUrl: '',
          MajorVersion: 7,
          MinorVersion: 0,
          Name: 'home.aspx',
          ServerRelativeUrl: '/sites/team-a/SitePages/home.aspx',
          TimeCreated: '2018-01-20T08:54:41Z',
          TimeLastModified: '2018-04-12T10:42:46Z',
          Title: 'Home',
          UIVersion: 3584,
          UIVersionLabel: '7.0',
          UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body).endsWith(`position&quot;&#58;&#123;&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;\"><div data-sp-rte=\"\"><p>Hello world</p></div></div></div>`)) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          text: 'Hello world'
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.notCalled);
          done();
        }
        catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles OData error when adding text to a non-existing page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/foo.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`) > -1) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'The file /sites/team-a/SitePages/foo.aspx does not exist' } } } });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'foo.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          text: 'Hello world'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The file /sites/team-a/SitePages/foo.aspx does not exist')));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles OData error when adding text to a page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`) > -1) {
        return Promise.resolve({
          ListItemAllFields: {
            CommentsDisabled: false,
            FileSystemObjectType: 0,
            Id: 1,
            ServerRedirectedEmbedUri: null,
            ServerRedirectedEmbedUrl: '',
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
            FileLeafRef: 'page.aspx',
            ComplianceAssetId: null,
            WikiField: null,
            Title: 'Page',
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: 'Article',
            CanvasContent1:
              '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true&#125;&#125;"></div></div>',
            BannerImageUrl: {
              Description: '/_layouts/15/images/sitepagethumbnail.png',
              Url: 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'
            },
            Description: 'Lorem ipsum Dolor samet Lorem ipsum',
            PromotedState: null,
            FirstPublishedDate: null,
            LayoutWebpartsContent: null,
            AuthorsId: null,
            AuthorsStringId: null,
            OriginalSourceUrl: null,
            ID: 1,
            Created: '2018-01-20T09:54:41',
            AuthorId: 1073741823,
            Modified: '2018-04-12T12:42:47',
            EditorId: 12,
            OData__CopySource: null,
            CheckoutUserId: null,
            OData__UIVersionString: '7.0',
            GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
          },
          CheckInComment: '',
          CheckOutType: 2,
          ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
          CustomizedPageStatus: 1,
          ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
          Exists: true,
          IrmEnabled: false,
          Length: '805',
          Level: 1,
          LinkingUri: null,
          LinkingUrl: '',
          MajorVersion: 7,
          MinorVersion: 0,
          Name: 'home.aspx',
          ServerRelativeUrl: '/sites/team-a/SitePages/page.aspx',
          TimeCreated: '2018-01-20T08:54:41Z',
          TimeLastModified: '2018-04-12T10:42:46Z',
          Title: 'Home',
          UIVersion: 3584,
          UIVersionLabel: '7.0',
          UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          text: 'Hello world'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles error if target page is not a modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`) > -1) {
        return Promise.resolve({
          ListItemAllFields: {
            CommentsDisabled: false,
            FileSystemObjectType: 0,
            Id: 1,
            ServerRedirectedEmbedUri: null,
            ServerRedirectedEmbedUrl: '',
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
            FileLeafRef: 'Page.aspx',
            ComplianceAssetId: null,
            WikiField: null,
            Title: 'Page',
            Description: 'Lorem ipsum Dolor samet Lorem ipsum',
            PromotedState: null,
            FirstPublishedDate: null,
            LayoutWebpartsContent: null,
            AuthorsId: null,
            AuthorsStringId: null,
            OriginalSourceUrl: null,
            ID: 1,
            Created: '2018-01-20T09:54:41',
            AuthorId: 1073741823,
            Modified: '2018-04-12T12:42:47',
            EditorId: 12,
            OData__CopySource: null,
            CheckoutUserId: null,
            OData__UIVersionString: '7.0',
            GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
          },
          CheckInComment: '',
          CheckOutType: 2,
          ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
          CustomizedPageStatus: 1,
          ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
          Exists: true,
          IrmEnabled: false,
          Length: '805',
          Level: 1,
          LinkingUri: null,
          LinkingUrl: '',
          MajorVersion: 7,
          MinorVersion: 0,
          Name: 'home.aspx',
          ServerRelativeUrl: '/sites/team-a/SitePages/page.aspx',
          TimeCreated: '2018-01-20T08:54:41Z',
          TimeLastModified: '2018-04-12T10:42:46Z',
          Title: 'Home',
          UIVersion: 3584,
          UIVersionLabel: '7.0',
          UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          text: 'Hello world'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Page page.aspx is not a modern page.`)));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles invalid section error when adding text to modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`) > -1) {
        return Promise.resolve({
          ListItemAllFields: {
            CommentsDisabled: false,
            FileSystemObjectType: 0,
            Id: 1,
            ServerRedirectedEmbedUri: null,
            ServerRedirectedEmbedUrl: '',
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
            FileLeafRef: 'page.aspx',
            ComplianceAssetId: null,
            WikiField: null,
            Title: 'Page',
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: 'Article',
            CanvasContent1:
              '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true&#125;&#125;"></div></div>',
            BannerImageUrl: {
              Description: '/_layouts/15/images/sitepagethumbnail.png',
              Url: 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'
            },
            Description: 'Lorem ipsum Dolor samet Lorem ipsum',
            PromotedState: null,
            FirstPublishedDate: null,
            LayoutWebpartsContent: null,
            AuthorsId: null,
            AuthorsStringId: null,
            OriginalSourceUrl: null,
            ID: 1,
            Created: '2018-01-20T09:54:41',
            AuthorId: 1073741823,
            Modified: '2018-04-12T12:42:47',
            EditorId: 12,
            OData__CopySource: null,
            CheckoutUserId: null,
            OData__UIVersionString: '7.0',
            GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
          },
          CheckInComment: '',
          CheckOutType: 2,
          ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
          CustomizedPageStatus: 1,
          ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
          Exists: true,
          IrmEnabled: false,
          Length: '805',
          Level: 1,
          LinkingUri: null,
          LinkingUrl: '',
          MajorVersion: 7,
          MinorVersion: 0,
          Name: 'home.aspx',
          ServerRelativeUrl: '/sites/team-a/SitePages/page.aspx',
          TimeCreated: '2018-01-20T08:54:41Z',
          TimeLastModified: '2018-04-12T10:42:46Z',
          Title: 'Home',
          UIVersion: 3584,
          UIVersionLabel: '7.0',
          UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          text: 'Hello world',
          section: 8
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Invalid section '8'")));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles invalid column error when adding text to modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`) > -1) {
        return Promise.resolve({
          ListItemAllFields: {
            CommentsDisabled: false,
            FileSystemObjectType: 0,
            Id: 1,
            ServerRedirectedEmbedUri: null,
            ServerRedirectedEmbedUrl: '',
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
            FileLeafRef: 'page.aspx',
            ComplianceAssetId: null,
            WikiField: null,
            Title: 'Page',
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: 'Article',
            CanvasContent1:
              '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true&#125;&#125;"></div></div>',
            BannerImageUrl: {
              Description: '/_layouts/15/images/sitepagethumbnail.png',
              Url: 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'
            },
            Description: 'Lorem ipsum Dolor samet Lorem ipsum',
            PromotedState: null,
            FirstPublishedDate: null,
            LayoutWebpartsContent: null,
            AuthorsId: null,
            AuthorsStringId: null,
            OriginalSourceUrl: null,
            ID: 1,
            Created: '2018-01-20T09:54:41',
            AuthorId: 1073741823,
            Modified: '2018-04-12T12:42:47',
            EditorId: 12,
            OData__CopySource: null,
            CheckoutUserId: null,
            OData__UIVersionString: '7.0',
            GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
          },
          CheckInComment: '',
          CheckOutType: 2,
          ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
          CustomizedPageStatus: 1,
          ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
          Exists: true,
          IrmEnabled: false,
          Length: '805',
          Level: 1,
          LinkingUri: null,
          LinkingUrl: '',
          MajorVersion: 7,
          MinorVersion: 0,
          Name: 'home.aspx',
          ServerRelativeUrl: '/sites/team-a/SitePages/page.aspx',
          TimeCreated: '2018-01-20T08:54:41Z',
          TimeLastModified: '2018-04-12T10:42:46Z',
          Title: 'Home',
          UIVersion: 3584,
          UIVersionLabel: '7.0',
          UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          text: 'Hello world',
          section: 1,
          column: 7
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Invalid column '7'")));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles error when parsing modern page contents', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`) > -1) {
        return Promise.resolve({
          ListItemAllFields: {
            CommentsDisabled: false,
            FileSystemObjectType: 0,
            Id: 1,
            ServerRedirectedEmbedUri: null,
            ServerRedirectedEmbedUrl: '',
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
            FileLeafRef: 'page.aspx',
            ComplianceAssetId: null,
            WikiField: null,
            Title: 'Page',
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: 'Article',
            CanvasContent1:
              '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&qu"></div></div>',
            BannerImageUrl: {
              Description: '/_layouts/15/images/sitepagethumbnail.png',
              Url: 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'
            },
            Description: 'Lorem ipsum Dolor samet Lorem ipsum',
            PromotedState: null,
            FirstPublishedDate: null,
            LayoutWebpartsContent: null,
            AuthorsId: null,
            AuthorsStringId: null,
            OriginalSourceUrl: null,
            ID: 1,
            Created: '2018-01-20T09:54:41',
            AuthorId: 1073741823,
            Modified: '2018-04-12T12:42:47',
            EditorId: 12,
            OData__CopySource: null,
            CheckoutUserId: null,
            OData__UIVersionString: '7.0',
            GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
          },
          CheckInComment: '',
          CheckOutType: 2,
          ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
          CustomizedPageStatus: 1,
          ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
          Exists: true,
          IrmEnabled: false,
          Length: '805',
          Level: 1,
          LinkingUri: null,
          LinkingUrl: '',
          MajorVersion: 7,
          MinorVersion: 0,
          Name: 'home.aspx',
          ServerRelativeUrl: '/sites/team-a/SitePages/page.aspx',
          TimeCreated: '2018-01-20T08:54:41Z',
          TimeLastModified: '2018-04-12T10:42:46Z',
          Title: 'Home',
          UIVersion: 3584,
          UIVersionLabel: '7.0',
          UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          text: 'Hello world',
          section: 1,
          column: 1
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(new CommandError("Unexpected end of JSON input")), JSON.stringify(err));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('supports debug mode', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports verbose mode', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--verbose') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page name', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--pageName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying section', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--section') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying column', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--column') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying order', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--order') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not an absolute URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { pageName: 'page.aspx', webUrl: 'foo', text: 'Hello world' }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'http://foo',
        text: 'Hello world'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webUrl specified, webUrl is a valid SharePoint URL and text is specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        text: 'Hello world'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name has no extension', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page',
        webUrl: 'https://contoso.sharepoint.com',
        text: 'Hello world'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if section has invalid (negative) value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        text: 'Hello world',
        section: -1
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if section has invalid (non number) value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        text: 'Hello world',
        section: 'foobar'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if column has invalid (negative) value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        text: 'Hello world',
        column: -1
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if column has invalid (non number) value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        text: 'Hello world',
        column: 'foobar'
      }
    });
    assert.notStrictEqual(actual, true);
  });
});
