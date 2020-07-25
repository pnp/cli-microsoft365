import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./page-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.PAGE_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
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
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.PAGE_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates new modern page', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(true)') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates new modern page (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(true)') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates new modern page on root of tenant (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(true)') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('automatically appends the .aspx extension', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(true)') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'page', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets page title when specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'My page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(true)') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'page.aspx', title: 'My page', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates new modern page using the Home layout', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Home'
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(true)') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates new modern page and promotes it as NewsPage', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(true)') > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        opts.body.PromotedState === 2 &&
        opts.body.FirstPublishedDate) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', promoteAs: 'NewsPage' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates new modern page and promotes it as Template', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        !opts.body) {
        return Promise.resolve({ Id: '1' });
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsTemplate`) > -1) {
        return Promise.resolve({ Id: '2', BannerImageUrl: 'url', CanvasContent1: 'content1', LayoutWebpartsContent: 'content', UniqueId: 'a4eb92e3-4eae-427f-8f6d-4e2ed907c2c4' });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('a4eb92e3-4eae-427f-8f6d-4e2ed907c2c4')/ListItemAllFields/SetCommentsDisabled`) > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(2)/SavePage`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', promoteAs: 'Template' } }, (res: { Id: string }) => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates new modern page using the Home layout and promotes it as HomePage (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Home'
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(true)') > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/rootfolder') > -1 &&
        opts.body.WelcomePage === 'SitePages/page.aspx') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home', promoteAs: 'HomePage' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates new modern page with comments enabled', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(false)') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', commentsEnabled: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates new modern page and publishes it', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(true)') > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/Publish(\'\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates new modern page and publishes it with a message (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(true)') > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/Publish(\'Initial%20version\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: 'Initial version' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes special characters in user input', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfolderbyserverrelativeurl('/sites/team-a/sitepages')/files/AddTemplateFile`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          urlOfFile: '/sites/team-a/sitepages/page.aspx',
          templateFileType: 3
        })) {
        return Promise.resolve({
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{64201083-46BA-4966-8BC5-B0CB31E3456C},1,0",
          "CustomizedPageStatus": 1,
          "ETag": "\"{64201083-46BA-4966-8BC5-B0CB31E3456C},1\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "780",
          "Level": 2,
          "LinkingUri": null,
          "LinkingUrl": "",
          "MajorVersion": 0,
          "MinorVersion": 1,
          "Name": "page.aspx",
          "ServerRelativeUrl": "/sites/team-a/SitePages/page.aspx",
          "TimeCreated": "2018-03-18T20:44:17Z",
          "TimeLastModified": "2018-03-18T20:44:17Z",
          "Title": null,
          "UIVersion": 1,
          "UIVersionLabel": "0.1",
          "UniqueId": "64201083-46ba-4966-8bc5-b0cb31e3456c"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyid('64201083-46ba-4966-8bc5-b0cb31e3456c')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
          Title: 'page',
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/ListItemAllFields/SetCommentsDisabled(true)') > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/getfilebyid(\'64201083-46ba-4966-8bc5-b0cb31e3456c\')/Publish(\'Don%39t%20tell\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: 'Don\'t tell' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when creating modern page', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, (err?: any) => {
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

  it('supports specifying name', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page layout', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--layoutType') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page promote option', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--promoteAs') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying if comments should be enabled', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--commentsEnabled') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying if page should be published', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publish') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page publish message', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publishMessage') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not an absolute URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'http://foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name has no extension', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page', webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if layout type is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if layout type is Home', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Home' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is Article', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Article' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if promote type is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if promote type is HomePage', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Home' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if promote type is NewsPage', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if promote type is Template', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'Template' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if promote type is HomePage but layout type is not Home', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Article' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if promote type is NewsPage but layout type is not Article', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage', layoutType: 'Home' } });
    assert.notStrictEqual(actual, true);
  });
});
