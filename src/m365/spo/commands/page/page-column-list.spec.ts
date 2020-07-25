import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./page-column-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { ClientSidePage } from './clientsidepages';

describe(commands.PAGE_COLUMN_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  const apiResponse = {
    "ListItemAllFields": {
      "FileSystemObjectType": 0,
      "Id": 9,
      "ServerRedirectedEmbedUri": null,
      "ServerRedirectedEmbedUrl": "",
      "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180070E97A63FCC58F47B8FE04D0654FD44E",
      "WikiField": null,
      "Title": "Nova",
      "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
      "CanvasContent1": "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;controlType&quot;&#58;4,&quot;displayMode&quot;&#58;2,&quot;id&quot;&#58;&quot;ccaa96dc-4d16-4940-bf0d-7b179628a8fd&quot;,&quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;1,&quot;sectionIndex&quot;&#58;1,&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;6&#125;,&quot;editorType&quot;&#58;&quot;CKEditor&quot;&#125;\"><div data-sp-rte=\"\"><p>asd</p>\n</div></div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionIndex&quot;&#58;2,&quot;sectionFactor&quot;&#58;6,&quot;zoneIndex&quot;&#58;1&#125;&#125;\"></div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;6,&quot;zoneIndex&quot;&#58;2&#125;&#125;\"></div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionIndex&quot;&#58;2,&quot;sectionFactor&quot;&#58;6,&quot;zoneIndex&quot;&#58;2&#125;&#125;\"></div></div>",
      "BannerImageUrl": {
        "Description": "/_layouts/15/images/sitepagethumbnail.png",
        "Url": "/_layouts/15/images/sitepagethumbnail.png"
      },
      "Description": "asd",
      "PromotedState": 0,
      "FirstPublishedDate": null,
      "LayoutWebpartsContent": "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.4\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Região do Título&quot;,&quot;description&quot;&#58;&quot;Descrição da Região de Título&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;Nova&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showKicker&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;kicker&quot;&#58;&quot;&quot;&#125;&#125;\"></div></div>",
      "ComplianceAssetId": null,
      "OData__AuthorBylineId": null,
      "_AuthorBylineStringId": null,
      "OData__OriginalSourceUrl": null,
      "OData__OriginalSourceSiteId": null,
      "OData__OriginalSourceWebId": null,
      "OData__OriginalSourceListId": null,
      "OData__OriginalSourceItemId": null,
      "ID": 9,
      "Created": "2018-07-11T16:24:12",
      "AuthorId": 9,
      "Modified": "2018-07-11T16:33:57",
      "EditorId": 9,
      "OData__CopySource": null,
      "CheckoutUserId": 9,
      "OData__UIVersionString": "1.0",
      "GUID": "903cdabe-7a28-4e96-a55e-c768185d7d9a"
    },
    "CheckInComment": "",
    "CheckOutType": 0,
    "ContentTag": "{16035D61-EDB9-4758-A490-3D13FCD9FDAA},10,8",
    "CustomizedPageStatus": 0,
    "ETag": "\"{16035D61-EDB9-4758-A490-3D13FCD9FDAA},10\"",
    "Exists": true,
    "IrmEnabled": false,
    "Length": "4708",
    "Level": 255,
    "LinkingUri": null,
    "LinkingUrl": "",
    "MajorVersion": 1,
    "MinorVersion": 0,
    "Name": "Nova.aspx",
    "ServerRelativeUrl": "/SitePages/Nova.aspx",
    "TimeCreated": "2018-07-11T19:24:12Z",
    "TimeLastModified": "2018-07-11T19:33:57Z",
    "Title": "Nova",
    "UIVersion": 512,
    "UIVersionLabel": "1.0",
    "UniqueId": "16035d61-edb9-4758-a490-3d13fcd9fdaa"
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
      request.get,
      ClientSidePage.fromHtml
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_COLUMN_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists columns on the modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(apiResponse);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', section: 1 } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "order": 1,
            "factor": 6,
            "controls": 1
          },
          {
            "order": 2,
            "factor": 6,
            "controls": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists columns on the modern page - no sections available', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve({
          "ListItemAllFields": {
            "FileSystemObjectType": 0,
            "Id": 9,
            "ServerRedirectedEmbedUri": null,
            "ServerRedirectedEmbedUrl": "",
            "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180070E97A63FCC58F47B8FE04D0654FD44E",
            "WikiField": null,
            "Title": "Nova",
            "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
            "CanvasContent1": "<div></div>",
            "BannerImageUrl": {
              "Description": "/_layouts/15/images/sitepagethumbnail.png",
              "Url": "/_layouts/15/images/sitepagethumbnail.png"
            },
            "Description": "asd",
            "PromotedState": 0,
            "FirstPublishedDate": null,
            "LayoutWebpartsContent": "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.4\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Região do Título&quot;,&quot;description&quot;&#58;&quot;Descrição da Região de Título&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;Nova&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showKicker&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;kicker&quot;&#58;&quot;&quot;&#125;&#125;\"></div></div>",
            "ComplianceAssetId": null,
            "OData__AuthorBylineId": null,
            "_AuthorBylineStringId": null,
            "OData__OriginalSourceUrl": null,
            "OData__OriginalSourceSiteId": null,
            "OData__OriginalSourceWebId": null,
            "OData__OriginalSourceListId": null,
            "OData__OriginalSourceItemId": null,
            "ID": 9,
            "Created": "2018-07-11T16:24:12",
            "AuthorId": 9,
            "Modified": "2018-07-11T16:33:57",
            "EditorId": 9,
            "OData__CopySource": null,
            "CheckoutUserId": 9,
            "OData__UIVersionString": "1.0",
            "GUID": "903cdabe-7a28-4e96-a55e-c768185d7d9a"
          },
          "CheckInComment": "",
          "CheckOutType": 0,
          "ContentTag": "{16035D61-EDB9-4758-A490-3D13FCD9FDAA},10,8",
          "CustomizedPageStatus": 0,
          "ETag": "\"{16035D61-EDB9-4758-A490-3D13FCD9FDAA},10\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "4708",
          "Level": 255,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 1,
          "MinorVersion": 0,
          "Name": "Nova.aspx",
          "ServerRelativeUrl": "/SitePages/Nova.aspx",
          "TimeCreated": "2018-07-11T19:24:12Z",
          "TimeLastModified": "2018-07-11T19:33:57Z",
          "Title": "Nova",
          "UIVersion": 512,
          "UIVersionLabel": "1.0",
          "UniqueId": "16035d61-edb9-4758-a490-3d13fcd9fdaa"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', section: 1 } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists columns on the modern page (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(apiResponse);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', section: 1 } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "order": 1,
            "factor": 6,
            "controls": 1
          },
          {
            "order": 2,
            "factor": 6,
            "controls": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists columns on the modern page when the specified page name doesn\'t contain extension', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(apiResponse);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home', section: 1 } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "order": 1,
            "factor": 6,
            "controls": 1
          },
          {
            "order": 2,
            "factor": 6,
            "controls": 0
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists all information about columns on the modern page in json output mode', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(apiResponse);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', output: 'json', section: 1 } }, () => {
      try {
        assert.strictEqual(JSON.stringify(log[0]), JSON.stringify([{
          "factor": 6,
          "order": 1,
          "dataVersion": "1.0",
          "jsonData": "&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionFactor&quot;&#58;6,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;",
          "controls": 1
        },
        {
          "factor": 6,
          "order": 2,
          "dataVersion": "1.0",
          "jsonData": "&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionFactor&quot;&#58;6,&quot;sectionIndex&quot;&#58;2,&quot;zoneIndex&quot;&#58;1&#125;&#125;",
          "controls": 0
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error when the specified page is a classic page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve({
          "ListItemAllFields": {
            "CommentsDisabled": false,
            "FileSystemObjectType": 0,
            "Id": 1,
            "ServerRedirectedEmbedUri": null,
            "ServerRedirectedEmbedUrl": "",
            "ContentTypeId": "0x0101080088E2A2ED69D0324A8981DD7FAC103494",
            "FileLeafRef": "Home.aspx",
            "ComplianceAssetId": null,
            "WikiField": "<div class=\"ExternalClass1188FC9011E046D4BED9C05BAD4DA96E\">\r\n            <table id=\"layoutsTable\" style=\"width&#58;100%;\">\r\n                <tbody>\r\n                    <tr style=\"vertical-align&#58;top;\">\r\n            <td colspan=\"2\"><div class=\"ms-rte-layoutszone-outer\" style=\"width&#58;100%;\"><div class=\"ms-rte-layoutszone-inner\" style=\"word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;\"><div class=\"ms-rtestate-read ms-rte-wpbox\"><div class=\"ms-rtestate-read f01b62ca-c190-410c-aef9-2499ab79436e\" id=\"div_f01b62ca-c190-410c-aef9-2499ab79436e\"></div>\n  <div class=\"ms-rtestate-read\" id=\"vid_f01b62ca-c190-410c-aef9-2499ab79436e\" style=\"display&#58;none;\"></div>\n</div>\n</div></div></td>\r\n                    </tr>\r\n                    <tr style=\"vertical-align&#58;top;\">\r\n            <td style=\"width&#58;49.95%;\"><div class=\"ms-rte-layoutszone-outer\" style=\"width&#58;100%;\"><div class=\"ms-rte-layoutszone-inner\" style=\"word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;\"><div class=\"ms-rtestate-read ms-rte-wpbox\"><div class=\"ms-rtestate-read 837b046b-6a02-4770-9a25-3292d955e903\" id=\"div_837b046b-6a02-4770-9a25-3292d955e903\"></div>\n  <div class=\"ms-rtestate-read\" id=\"vid_837b046b-6a02-4770-9a25-3292d955e903\" style=\"display&#58;none;\"></div>\n</div>\n</div></div></td>\r\n            <td class=\"ms-wiki-columnSpacing\" style=\"width&#58;49.95%;\"><div class=\"ms-rte-layoutszone-outer\" style=\"width&#58;100%;\"><div class=\"ms-rte-layoutszone-inner\" style=\"word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;\"><div class=\"ms-rtestate-read ms-rte-wpbox\"><div class=\"ms-rtestate-read f36dd97b-6f2b-437b-a169-26a97962074d\" id=\"div_f36dd97b-6f2b-437b-a169-26a97962074d\"></div>\n  <div class=\"ms-rtestate-read\" id=\"vid_f36dd97b-6f2b-437b-a169-26a97962074d\" style=\"display&#58;none;\"></div>\n</div>\n</div></div></td>\r\n                    </tr>\r\n                </tbody>\r\n            </table>\r\n            <span id=\"layoutsData\" style=\"display&#58;none;\">true,false,2</span></div>",
            "Title": null,
            "ClientSideApplicationId": null,
            "PageLayoutType": null,
            "CanvasContent1": null,
            "BannerImageUrl": null,
            "Description": null,
            "PromotedState": null,
            "FirstPublishedDate": null,
            "LayoutWebpartsContent": null,
            "AuthorsId": null,
            "AuthorsStringId": null,
            "OriginalSourceUrl": null,
            "ID": 1,
            "Created": "2018-03-19T17:52:56",
            "AuthorId": 1073741823,
            "Modified": "2018-03-24T07:14:28",
            "EditorId": 1073741823,
            "OData__CopySource": null,
            "CheckoutUserId": null,
            "OData__UIVersionString": "1.0",
            "GUID": "19ac5510-bba6-427b-9c1b-a3329a3b0cad"
          },
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{8F33F78C-9F39-48E2-B99D-01C2937A56BB},4,1",
          "CustomizedPageStatus": 1,
          "ETag": "\"{8F33F78C-9F39-48E2-B99D-01C2937A56BB},4\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "3356",
          "Level": 1,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 1,
          "MinorVersion": 0,
          "Name": "home.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/home.aspx",
          "TimeCreated": "2018-03-20T00:52:56Z",
          "TimeLastModified": "2018-03-24T14:14:28Z",
          "Title": null,
          "UIVersion": 512,
          "UIVersionLabel": "1.0",
          "UniqueId": "8f33f78c-9f39-48e2-b99d-01c2937a56bb"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', section: 1 } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Page home.aspx is not a modern page.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles page not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-2130575338, Microsoft.SharePoint.SPException",
            "message": {
              "lang": "en-US",
              "value": "The file /sites/team-a/SitePages/home1.aspx does not exist."
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', section: 1 } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The file /sites/team-a/SitePages/home1.aspx does not exist.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when retrieving pages', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', section: 1 } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', name: 'home.aspx', section: 1 } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the section option is not specifed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the section option is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx', section: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL and name is specified and section is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx', section: 1 } });
    assert.strictEqual(actual, true);
  });
});