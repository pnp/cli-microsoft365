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
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./sitedesign-add');

describe(commands.SITEDESIGN_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
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
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITEDESIGN_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds new site design for a team site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24']
          }
        })) {
        return {
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "ThumbnailUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Contoso",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('adds new site design for a team site (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24']
          }
        })) {
        return {
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "ThumbnailUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Contoso",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('adds new team site site design with multiple site script IDs', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24', '449c0c6d-5380-4df2-b84b-622e0ac8ec25']
          }
        })) {
        return {
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "ThumbnailUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24", "449c0c6d-5380-4df2-b84b-622e0ac8ec25"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24, 449c0c6d-5380-4df2-b84b-622e0ac8ec25" } });
    assert(loggerLogSpy.calledOnce);
  });

  it('adds new site design for a communication site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '68',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24']
          }
        })) {
        return {
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "ThumbnailUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 68
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'Contoso', webTemplate: 'CommunicationSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Contoso",
      "Version": 1,
      "WebTemplate": 68
    }));
  });

  it('adds new team site site design with description', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            Description: 'Contoso team site'
          }
        })) {
        return {
          "Description": "Contoso team site",
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "ThumbnailUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", description: 'Contoso team site' } });
    assert(loggerLogSpy.calledWith({
      "Description": "Contoso team site",
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Contoso",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('adds new team site site design with previewImageUrl', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            PreviewImageUrl: 'https://contoso.com/assets/team-site-preview.png'
          }
        })) {
        return {
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", previewImageUrl: 'https://contoso.com/assets/team-site-preview.png' } });
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
  });

  it('adds new team site site design with ThumbnailUrl', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            ThumbnailUrl: 'https://contoso.com/assets/team-site-thumbnail.png'
          }
        })) {
        return {
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "ThumbnailUrl": "https://contoso.com/assets/team-site-thumbnail.png",
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", thumbnailUrl: 'https://contoso.com/assets/team-site-thumbnail.png' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": "https://contoso.com/assets/team-site-thumbnail.png",
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Contoso",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('adds new team site site design with previewImageAltText', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            PreviewImageAltText: 'Contoso team site preview'
          }
        })) {
        return {
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": null,
          "ThumbnailUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", previewImageAltText: 'Contoso team site preview' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": 'Contoso team site preview',
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Contoso",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('adds new default team site site design', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            IsDefault: true
          }
        })) {
        return {
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "ThumbnailUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", isDefault: true } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": true,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Contoso",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('adds new team site site design with all options specified', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            Description: 'Contoso team site',
            PreviewImageUrl: 'https://contoso.com/assets/team-site-preview.png',
            PreviewImageAltText: 'Contoso team site preview',
            ThumbnailUrl: 'https://contoso.com/assets/team-site-thumbnail.png',
            IsDefault: true
          }
        })) {
        return {
          "Description": 'Contoso team site',
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "ThumbnailUrl": 'https://contoso.com/assets/team-site-thumbnail.png',
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", description: 'Contoso team site', previewImageUrl: 'https://contoso.com/assets/team-site-preview.png', thumbnailUrl: 'https://contoso.com/assets/team-site-thumbnail.png', previewImageAltText: 'Contoso team site preview', isDefault: true } });
    assert(loggerLogSpy.calledWith({
      "Description": 'Contoso team site',
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": true,
      "PreviewImageAltText": "Contoso team site preview",
      "PreviewImageUrl": "https://contoso.com/assets/team-site-preview.png",
      "ThumbnailUrl": "https://contoso.com/assets/team-site-thumbnail.png",
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Contoso",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('correctly handles OData error when creating site script', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: 'Contoso',
        webTemplate: 'TeamSite',
        siteScripts: '449c0c6d-5380-4df2-b84b-622e0ac8ec24'
      }
    } as any), new CommandError('An error has occurred'));
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

  it('supports specifying thumbnailUrl', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--thumbnailUrl') > -1) {
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
