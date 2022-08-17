import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./sitedesign-add');

describe(commands.SITEDESIGN_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds new site design for a team site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24']
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new site design for a team site (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24']
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new team site site design wilt multiple site script IDs', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24', '449c0c6d-5380-4df2-b84b-622e0ac8ec25']
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24", "449c0c6d-5380-4df2-b84b-622e0ac8ec25"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24, 449c0c6d-5380-4df2-b84b-622e0ac8ec25" } }, () => {
      try {
        assert(loggerLogSpy.calledOnce);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new site design for a communication site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '68',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24']
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 68
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'Contoso', webTemplate: 'CommunicationSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 68
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new team site site design with description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            Description: 'Contoso team site'
          }
        })) {
        return Promise.resolve({
          "Description": "Contoso team site",
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", description: 'Contoso team site' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Description": "Contoso team site",
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new team site site design with previewImageUrl', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            PreviewImageUrl: 'https://contoso.com/assets/team-site-preview.png'
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", previewImageUrl: 'https://contoso.com/assets/team-site-preview.png' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new team site site design with previewImageAltText', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            PreviewImageAltText: 'Contoso team site preview'
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", previewImageAltText: 'Contoso team site preview' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new default team site site design', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            IsDefault: true
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", isDefault: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new team site site design with all options specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            Description: 'Contoso team site',
            PreviewImageUrl: 'https://contoso.com/assets/team-site-preview.png',
            PreviewImageAltText: 'Contoso team site preview',
            IsDefault: true
          }
        })) {
        return Promise.resolve({
          "Description": 'Contoso team site',
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", description: 'Contoso team site', previewImageUrl: 'https://contoso.com/assets/team-site-preview.png', previewImageAltText: 'Contoso team site preview', isDefault: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Description": 'Contoso team site',
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when creating site script', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    command.action(logger, { options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: '449c0c6d-5380-4df2-b84b-622e0ac8ec24' } } as any, (err?: any) => {
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
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying title', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--title') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webTemplate', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webTemplate') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying siteScripts', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--siteScripts') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying description', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying previewImageUrl', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--previewImageUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying previewImageAltText', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--previewImageAltText') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying isDefault', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--isDefault') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if specified webTemplate is invalid', async () => {
    const actual = await command.validate({ options: { title: 'Contoso', webTemplate: 'Invalid', siteScripts: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified siteScripts is not a valid GUID', async () => {
    const actual = await command.validate({ options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the second specified siteScriptId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24,abc" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required parameters are valid', async () => {
    const actual = await command.validate({ options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if all required parameters are valid (multiple siteScripts)', async () => {
    const actual = await command.validate({ options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24,449c0c6d-5380-4df2-b84b-622e0ac8ec25" } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});