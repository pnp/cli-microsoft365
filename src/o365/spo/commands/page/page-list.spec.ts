import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./page-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.PAGE_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.PAGE_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.PAGE_LIST);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists all modern pages', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "ListItemAllFields": {
                "FileSystemObjectType": 0,
                "Id": 122,
                "ServerRedirectedEmbedUri": null,
                "ServerRedirectedEmbedUrl": "",
                "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180023536E5F3BB7DA449A374D731B978084",
                "ComplianceAssetId": null,
                "WikiField": null,
                "Title": "page_118",
                "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
                "CanvasContent1": "<div></div>",
                "BannerImageUrl": {
                  "Description": "/_layouts/15/images/sitepagethumbnail.png",
                  "Url": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png"
                },
                "Description": null,
                "PromotedState": 0,
                "FirstPublishedDate": null,
                "LayoutWebpartsContent": null,
                "AuthorsId": null,
                "AuthorsStringId": null,
                "OriginalSourceUrl": null,
                "ID": 122,
                "Created": "2018-03-13T13:18:00",
                "AuthorId": 6,
                "Modified": "2018-03-13T13:18:01",
                "EditorId": 6,
                "OData__CopySource": null,
                "CheckoutUserId": null,
                "OData__UIVersionString": "0.3",
                "GUID": "b8920589-bbed-4e21-a1c1-1f4d93118caf"
              },
              "CheckInComment": "",
              "CheckOutType": 2,
              "ContentTag": "{6707E2AF-14B5-4FF1-A25D-001C6B44EEC2},3,2",
              "CustomizedPageStatus": 2,
              "ETag": "\"{6707E2AF-14B5-4FF1-A25D-001C6B44EEC2},3\"",
              "Exists": true,
              "IrmEnabled": false,
              "Length": "1899",
              "Level": 2,
              "LinkingUri": null,
              "LinkingUrl": "",
              "MajorVersion": 0,
              "MinorVersion": 3,
              "Name": "page_118.aspx",
              "ServerRelativeUrl": "/sites/team-a/SitePages/page_118.aspx",
              "TimeCreated": "2018-03-13T20:18:00Z",
              "TimeLastModified": "2018-03-13T20:18:01Z",
              "Title": "page_118",
              "UIVersion": 3,
              "UIVersionLabel": "0.3",
              "UniqueId": "6707e2af-14b5-4ff1-a25d-001c6b44eec2"
            },
            {
              "ListItemAllFields": {
                "FileSystemObjectType": 0,
                "Id": 723,
                "ServerRedirectedEmbedUri": null,
                "ServerRedirectedEmbedUrl": "",
                "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180023536E5F3BB7DA449A374D731B978084",
                "ComplianceAssetId": null,
                "WikiField": null,
                "Title": "page_719",
                "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
                "CanvasContent1": "<div></div>",
                "BannerImageUrl": {
                  "Description": "/_layouts/15/images/sitepagethumbnail.png",
                  "Url": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png"
                },
                "Description": null,
                "PromotedState": 0,
                "FirstPublishedDate": null,
                "LayoutWebpartsContent": null,
                "AuthorsId": null,
                "AuthorsStringId": null,
                "OriginalSourceUrl": null,
                "ID": 723,
                "Created": "2018-03-13T13:31:43",
                "AuthorId": 6,
                "Modified": "2018-03-13T13:31:44",
                "EditorId": 6,
                "OData__CopySource": null,
                "CheckoutUserId": null,
                "OData__UIVersionString": "0.3",
                "GUID": "e8cd0967-d340-4f48-aec6-b0fb73714f98"
              },
              "CheckInComment": "",
              "CheckOutType": 2,
              "ContentTag": "{3CCC58F9-7892-4132-9B0C-003686AB7C68},3,2",
              "CustomizedPageStatus": 2,
              "ETag": "\"{3CCC58F9-7892-4132-9B0C-003686AB7C68},3\"",
              "Exists": true,
              "IrmEnabled": false,
              "Length": "1899",
              "Level": 2,
              "LinkingUri": null,
              "LinkingUrl": "",
              "MajorVersion": 0,
              "MinorVersion": 3,
              "Name": "page_719.aspx",
              "ServerRelativeUrl": "/sites/team-a/SitePages/page_719.aspx",
              "TimeCreated": "2018-03-13T20:31:43Z",
              "TimeLastModified": "2018-03-13T20:31:44Z",
              "Title": "page_719",
              "UIVersion": 3,
              "UIVersionLabel": "0.3",
              "UniqueId": "3ccc58f9-7892-4132-9b0c-003686ab7c68"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Name: 'page_118.aspx',
            Title: 'page_118'
          },
          {
            Name: 'page_719.aspx',
            Title: 'page_719'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists all modern pages (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "ListItemAllFields": {
                "FileSystemObjectType": 0,
                "Id": 122,
                "ServerRedirectedEmbedUri": null,
                "ServerRedirectedEmbedUrl": "",
                "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180023536E5F3BB7DA449A374D731B978084",
                "ComplianceAssetId": null,
                "WikiField": null,
                "Title": "page_118",
                "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
                "CanvasContent1": "<div></div>",
                "BannerImageUrl": {
                  "Description": "/_layouts/15/images/sitepagethumbnail.png",
                  "Url": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png"
                },
                "Description": null,
                "PromotedState": 0,
                "FirstPublishedDate": null,
                "LayoutWebpartsContent": null,
                "AuthorsId": null,
                "AuthorsStringId": null,
                "OriginalSourceUrl": null,
                "ID": 122,
                "Created": "2018-03-13T13:18:00",
                "AuthorId": 6,
                "Modified": "2018-03-13T13:18:01",
                "EditorId": 6,
                "OData__CopySource": null,
                "CheckoutUserId": null,
                "OData__UIVersionString": "0.3",
                "GUID": "b8920589-bbed-4e21-a1c1-1f4d93118caf"
              },
              "CheckInComment": "",
              "CheckOutType": 2,
              "ContentTag": "{6707E2AF-14B5-4FF1-A25D-001C6B44EEC2},3,2",
              "CustomizedPageStatus": 2,
              "ETag": "\"{6707E2AF-14B5-4FF1-A25D-001C6B44EEC2},3\"",
              "Exists": true,
              "IrmEnabled": false,
              "Length": "1899",
              "Level": 2,
              "LinkingUri": null,
              "LinkingUrl": "",
              "MajorVersion": 0,
              "MinorVersion": 3,
              "Name": "page_118.aspx",
              "ServerRelativeUrl": "/sites/team-a/SitePages/page_118.aspx",
              "TimeCreated": "2018-03-13T20:18:00Z",
              "TimeLastModified": "2018-03-13T20:18:01Z",
              "Title": "page_118",
              "UIVersion": 3,
              "UIVersionLabel": "0.3",
              "UniqueId": "6707e2af-14b5-4ff1-a25d-001c6b44eec2"
            },
            {
              "ListItemAllFields": {
                "FileSystemObjectType": 0,
                "Id": 723,
                "ServerRedirectedEmbedUri": null,
                "ServerRedirectedEmbedUrl": "",
                "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180023536E5F3BB7DA449A374D731B978084",
                "ComplianceAssetId": null,
                "WikiField": null,
                "Title": "page_719",
                "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
                "CanvasContent1": "<div></div>",
                "BannerImageUrl": {
                  "Description": "/_layouts/15/images/sitepagethumbnail.png",
                  "Url": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png"
                },
                "Description": null,
                "PromotedState": 0,
                "FirstPublishedDate": null,
                "LayoutWebpartsContent": null,
                "AuthorsId": null,
                "AuthorsStringId": null,
                "OriginalSourceUrl": null,
                "ID": 723,
                "Created": "2018-03-13T13:31:43",
                "AuthorId": 6,
                "Modified": "2018-03-13T13:31:44",
                "EditorId": 6,
                "OData__CopySource": null,
                "CheckoutUserId": null,
                "OData__UIVersionString": "0.3",
                "GUID": "e8cd0967-d340-4f48-aec6-b0fb73714f98"
              },
              "CheckInComment": "",
              "CheckOutType": 2,
              "ContentTag": "{3CCC58F9-7892-4132-9B0C-003686AB7C68},3,2",
              "CustomizedPageStatus": 2,
              "ETag": "\"{3CCC58F9-7892-4132-9B0C-003686AB7C68},3\"",
              "Exists": true,
              "IrmEnabled": false,
              "Length": "1899",
              "Level": 2,
              "LinkingUri": null,
              "LinkingUrl": "",
              "MajorVersion": 0,
              "MinorVersion": 3,
              "Name": "page_719.aspx",
              "ServerRelativeUrl": "/sites/team-a/SitePages/page_719.aspx",
              "TimeCreated": "2018-03-13T20:31:43Z",
              "TimeLastModified": "2018-03-13T20:31:44Z",
              "Title": "page_719",
              "UIVersion": 3,
              "UIVersionLabel": "0.3",
              "UniqueId": "3ccc58f9-7892-4132-9b0c-003686ab7c68"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Name: 'page_118.aspx',
            Title: 'page_118'
          },
          {
            Name: 'page_719.aspx',
            Title: 'page_719'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists all properties for all modern pages in JSON output mode', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "ListItemAllFields": {
                "FileSystemObjectType": 0,
                "Id": 122,
                "ServerRedirectedEmbedUri": null,
                "ServerRedirectedEmbedUrl": "",
                "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180023536E5F3BB7DA449A374D731B978084",
                "ComplianceAssetId": null,
                "WikiField": null,
                "Title": "page_118",
                "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
                "CanvasContent1": "<div></div>",
                "BannerImageUrl": {
                  "Description": "/_layouts/15/images/sitepagethumbnail.png",
                  "Url": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png"
                },
                "Description": null,
                "PromotedState": 0,
                "FirstPublishedDate": null,
                "LayoutWebpartsContent": null,
                "AuthorsId": null,
                "AuthorsStringId": null,
                "OriginalSourceUrl": null,
                "ID": 122,
                "Created": "2018-03-13T13:18:00",
                "AuthorId": 6,
                "Modified": "2018-03-13T13:18:01",
                "EditorId": 6,
                "OData__CopySource": null,
                "CheckoutUserId": null,
                "OData__UIVersionString": "0.3",
                "GUID": "b8920589-bbed-4e21-a1c1-1f4d93118caf"
              },
              "CheckInComment": "",
              "CheckOutType": 2,
              "ContentTag": "{6707E2AF-14B5-4FF1-A25D-001C6B44EEC2},3,2",
              "CustomizedPageStatus": 2,
              "ETag": "\"{6707E2AF-14B5-4FF1-A25D-001C6B44EEC2},3\"",
              "Exists": true,
              "IrmEnabled": false,
              "Length": "1899",
              "Level": 2,
              "LinkingUri": null,
              "LinkingUrl": "",
              "MajorVersion": 0,
              "MinorVersion": 3,
              "Name": "page_118.aspx",
              "ServerRelativeUrl": "/sites/team-a/SitePages/page_118.aspx",
              "TimeCreated": "2018-03-13T20:18:00Z",
              "TimeLastModified": "2018-03-13T20:18:01Z",
              "Title": "page_118",
              "UIVersion": 3,
              "UIVersionLabel": "0.3",
              "UniqueId": "6707e2af-14b5-4ff1-a25d-001c6b44eec2"
            },
            {
              "ListItemAllFields": {
                "FileSystemObjectType": 0,
                "Id": 723,
                "ServerRedirectedEmbedUri": null,
                "ServerRedirectedEmbedUrl": "",
                "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180023536E5F3BB7DA449A374D731B978084",
                "ComplianceAssetId": null,
                "WikiField": null,
                "Title": "page_719",
                "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
                "CanvasContent1": "<div></div>",
                "BannerImageUrl": {
                  "Description": "/_layouts/15/images/sitepagethumbnail.png",
                  "Url": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png"
                },
                "Description": null,
                "PromotedState": 0,
                "FirstPublishedDate": null,
                "LayoutWebpartsContent": null,
                "AuthorsId": null,
                "AuthorsStringId": null,
                "OriginalSourceUrl": null,
                "ID": 723,
                "Created": "2018-03-13T13:31:43",
                "AuthorId": 6,
                "Modified": "2018-03-13T13:31:44",
                "EditorId": 6,
                "OData__CopySource": null,
                "CheckoutUserId": null,
                "OData__UIVersionString": "0.3",
                "GUID": "e8cd0967-d340-4f48-aec6-b0fb73714f98"
              },
              "CheckInComment": "",
              "CheckOutType": 2,
              "ContentTag": "{3CCC58F9-7892-4132-9B0C-003686AB7C68},3,2",
              "CustomizedPageStatus": 2,
              "ETag": "\"{3CCC58F9-7892-4132-9B0C-003686AB7C68},3\"",
              "Exists": true,
              "IrmEnabled": false,
              "Length": "1899",
              "Level": 2,
              "LinkingUri": null,
              "LinkingUrl": "",
              "MajorVersion": 0,
              "MinorVersion": 3,
              "Name": "page_719.aspx",
              "ServerRelativeUrl": "/sites/team-a/SitePages/page_719.aspx",
              "TimeCreated": "2018-03-13T20:31:43Z",
              "TimeLastModified": "2018-03-13T20:31:44Z",
              "Title": "page_719",
              "UIVersion": 3,
              "UIVersionLabel": "0.3",
              "UniqueId": "3ccc58f9-7892-4132-9b0c-003686ab7c68"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, output: 'json', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ListItemAllFields": {
              "FileSystemObjectType": 0,
              "Id": 122,
              "ServerRedirectedEmbedUri": null,
              "ServerRedirectedEmbedUrl": "",
              "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180023536E5F3BB7DA449A374D731B978084",
              "ComplianceAssetId": null,
              "WikiField": null,
              "Title": "page_118",
              "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
              "CanvasContent1": "<div></div>",
              "BannerImageUrl": {
                "Description": "/_layouts/15/images/sitepagethumbnail.png",
                "Url": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png"
              },
              "Description": null,
              "PromotedState": 0,
              "FirstPublishedDate": null,
              "LayoutWebpartsContent": null,
              "AuthorsId": null,
              "AuthorsStringId": null,
              "OriginalSourceUrl": null,
              "ID": 122,
              "Created": "2018-03-13T13:18:00",
              "AuthorId": 6,
              "Modified": "2018-03-13T13:18:01",
              "EditorId": 6,
              "OData__CopySource": null,
              "CheckoutUserId": null,
              "OData__UIVersionString": "0.3",
              "GUID": "b8920589-bbed-4e21-a1c1-1f4d93118caf"
            },
            "CheckInComment": "",
            "CheckOutType": 2,
            "ContentTag": "{6707E2AF-14B5-4FF1-A25D-001C6B44EEC2},3,2",
            "CustomizedPageStatus": 2,
            "ETag": "\"{6707E2AF-14B5-4FF1-A25D-001C6B44EEC2},3\"",
            "Exists": true,
            "IrmEnabled": false,
            "Length": "1899",
            "Level": 2,
            "LinkingUri": null,
            "LinkingUrl": "",
            "MajorVersion": 0,
            "MinorVersion": 3,
            "Name": "page_118.aspx",
            "ServerRelativeUrl": "/sites/team-a/SitePages/page_118.aspx",
            "TimeCreated": "2018-03-13T20:18:00Z",
            "TimeLastModified": "2018-03-13T20:18:01Z",
            "Title": "page_118",
            "UIVersion": 3,
            "UIVersionLabel": "0.3",
            "UniqueId": "6707e2af-14b5-4ff1-a25d-001c6b44eec2"
          },
          {
            "ListItemAllFields": {
              "FileSystemObjectType": 0,
              "Id": 723,
              "ServerRedirectedEmbedUri": null,
              "ServerRedirectedEmbedUrl": "",
              "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C41180023536E5F3BB7DA449A374D731B978084",
              "ComplianceAssetId": null,
              "WikiField": null,
              "Title": "page_719",
              "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
              "CanvasContent1": "<div></div>",
              "BannerImageUrl": {
                "Description": "/_layouts/15/images/sitepagethumbnail.png",
                "Url": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png"
              },
              "Description": null,
              "PromotedState": 0,
              "FirstPublishedDate": null,
              "LayoutWebpartsContent": null,
              "AuthorsId": null,
              "AuthorsStringId": null,
              "OriginalSourceUrl": null,
              "ID": 723,
              "Created": "2018-03-13T13:31:43",
              "AuthorId": 6,
              "Modified": "2018-03-13T13:31:44",
              "EditorId": 6,
              "OData__CopySource": null,
              "CheckoutUserId": null,
              "OData__UIVersionString": "0.3",
              "GUID": "e8cd0967-d340-4f48-aec6-b0fb73714f98"
            },
            "CheckInComment": "",
            "CheckOutType": 2,
            "ContentTag": "{3CCC58F9-7892-4132-9B0C-003686AB7C68},3,2",
            "CustomizedPageStatus": 2,
            "ETag": "\"{3CCC58F9-7892-4132-9B0C-003686AB7C68},3\"",
            "Exists": true,
            "IrmEnabled": false,
            "Length": "1899",
            "Level": 2,
            "LinkingUri": null,
            "LinkingUrl": "",
            "MajorVersion": 0,
            "MinorVersion": 3,
            "Name": "page_719.aspx",
            "ServerRelativeUrl": "/sites/team-a/SitePages/page_719.aspx",
            "TimeCreated": "2018-03-13T20:31:43Z",
            "TimeLastModified": "2018-03-13T20:31:44Z",
            "Title": "page_719",
            "UIVersion": 3,
            "UIVersionLabel": "0.3",
            "UniqueId": "3ccc58f9-7892-4132-9b0c-003686ab7c68"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no modern pages', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
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

  it('fails validation if the webUrl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.equal(actual, true);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.PAGE_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});