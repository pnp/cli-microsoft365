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
import command from './sitedesign-set.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.SITEDESIGN_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITEDESIGN_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates site design title', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            Title: 'New title'
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
          "Title": "New title",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', title: 'New title' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "New title",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('updates site design web template to TeamSite', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            WebTemplate: '64'
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
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', webTemplate: 'TeamSite' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Title",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('updates site design web template to CommunicationSite', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            WebTemplate: '68'
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
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 68
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', webTemplate: 'CommunicationSite' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Title",
      "Version": 1,
      "WebTemplate": 68
    }));
  });

  it('updates site design site scripts (one script)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
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
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', siteScripts: '449c0c6d-5380-4df2-b84b-622e0ac8ec24' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Title",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('updates site design site scripts (multiple scripts)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
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
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', siteScripts: '449c0c6d-5380-4df2-b84b-622e0ac8ec24, 449c0c6d-5380-4df2-b84b-622e0ac8ec25' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24", "449c0c6d-5380-4df2-b84b-622e0ac8ec25"],
      "Title": "Title",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('updates site design description', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            Description: 'New description'
          }
        })) {
        return {
          "Description": "New description",
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "ThumbnailUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', description: 'New description' } });
    assert(loggerLogSpy.calledWith({
      "Description": "New description",
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Title",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('updates site design previewImageUrl', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            PreviewImageUrl: 'https://contoso.com/image.png'
          }
        })) {
        return {
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": "https://contoso.com/image.png",
          "ThumbnailUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', previewImageUrl: 'https://contoso.com/image.png' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": "https://contoso.com/image.png",
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Title",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('updates site design previewImageAltText', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            PreviewImageAltText: 'Logo image'
          }
        })) {
        return {
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": "Logo image",
          "PreviewImageUrl": null,
          "ThumbnailUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', previewImageAltText: 'Logo image' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": "Logo image",
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Title",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('updates site design thumbnailUrl', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            ThumbnailUrl: 'https://contoso.com/assets/team-site-thumbnail.png'
          }
        })) {
        return {
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageUrl": null,
          "PreviewImageAltText": null,
          "ThumbnailUrl": "https://contoso.com/assets/team-site-thumbnail.png",
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', thumbnailUrl: 'https://contoso.com/assets/team-site-thumbnail.png' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageUrl": null,
      "PreviewImageAltText": null,
      "ThumbnailUrl": "https://contoso.com/assets/team-site-thumbnail.png",
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Title",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('updates site design version', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            Version: 2
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
          "Title": "Title",
          "Version": 2,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', version: 2 } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Title",
      "Version": 2,
      "WebTemplate": 64
    }));
  });

  it('makes site design default', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
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
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', isDefault: true } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": true,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Title",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('makes site design not-default (explicit)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            IsDefault: false
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
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c', isDefault: false } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Title",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('makes site design not-default (implicit)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c'
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
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '2a9f178a-4d1d-449c-9296-df509ab4702c' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "ThumbnailUrl": null,
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Title",
      "Version": 1,
      "WebTemplate": 64
    }));
  });

  it('updates all site design properties (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          "updateInfo": {
            "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
            "Title": "Contoso",
            "Description": "Contoso team site",
            "SiteScriptIds": [
              "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
            ],
            "PreviewImageUrl": "https://contoso.com/assets/team-site-preview.png",
            "PreviewImageAltText": "Contoso team site preview",
            "ThumbnailUrl": "https://contoso.com/assets/team-site-thumbnail.png",
            "WebTemplate": "64",
            "Version": 2,
            "IsDefault": true
          }
        })) {
        return {
          "Description": 'Contoso team site',
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "ThumbnailUrl": "https://contoso.com/assets/team-site-thumbnail.png",
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 2,
          "WebTemplate": 64
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", description: 'Contoso team site', previewImageUrl: 'https://contoso.com/assets/team-site-preview.png', thumbnailUrl: "https://contoso.com/assets/team-site-thumbnail.png", previewImageAltText: 'Contoso team site preview', version: 2, isDefault: true } });
    assert(loggerLogSpy.calledWith({
      "Description": 'Contoso team site',
      "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
      "IsDefault": true,
      "PreviewImageAltText": 'Contoso team site preview',
      "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
      "ThumbnailUrl": "https://contoso.com/assets/team-site-thumbnail.png",
      "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
      "Title": "Contoso",
      "Version": 2,
      "WebTemplate": 64
    }));
  });

  it('correctly handles OData error when updating site design', async () => {
    sinon.stub(request, 'post').rejects({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });

    await assert.rejects(command.action(logger, {
      options: {
        id: '9b142c22-037f-4a7f-9017-e9d8c0e34b98',
        webTemplate: 'TeamSite',
        siteScripts: '449c0c6d-5380-4df2-b84b-622e0ac8ec24'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
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

  it('supports specifying version', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--version') > -1) {
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

  it('fails validation if id specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passed validation if id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified webTemplate is invalid', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if specified webTemplate is CommunicationSite', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'CommunicationSite' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if specified webTemplate is TeamSite', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'TeamSite' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified siteScripts is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'TeamSite', siteScripts: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the second specified siteScriptId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24,abc" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if specified siteScriptId is valid', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if all required parameters are valid (multiple siteScripts)', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24,449c0c6d-5380-4df2-b84b-622e0ac8ec25" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified version is not a number', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', version: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if specified version is a number', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', version: 2 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if specified isDefault value is true', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', isDefault: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if specified isDefault value is false', async () => {
    const actual = await command.validate({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', isDefault: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});