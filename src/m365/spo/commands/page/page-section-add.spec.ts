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
import command, { options } from './page-section-add.js';

describe(commands.PAGE_SECTION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    sinonUtil.restore([request.post, request.get]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_SECTION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('checks out page if not checked out by the current user', async () => {
    let checkedOut = false;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": false,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`) {
        checkedOut = true;
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        pageName: 'home',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      })
    });
    assert.deepEqual(checkedOut, true);
  });

  it('doesn\'t check out page if not checked out by the current user', async () => {
    let checkingOut = false;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`) {
        checkingOut = true;
        return;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      })
    });
    assert.deepEqual(checkingOut, false);
  });

  it('adds a first section to an uncustomized page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a first section to an uncustomized page with order set to 1', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        order: 1
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a first section to an uncustomized page correctly even when CanvasContent1 of returned page is null', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": null
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a first section to the page if no order specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a first section to the page if order 1 specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumnFullWidth',
        order: 1
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":0,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a section to the beginning of the page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnLeft',
        order: 1
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a section to the end of the page when order not specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnRight'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a section to the end of the page when order set to last section', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnRight',
        order: 2
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a section to the end of the page when order is larger than the last section', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnRight',
        order: 5
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a section between two other sections', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'ThreeColumn',
        order: 2
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":3,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":3,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":3,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a section between two other sections (2)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":2,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":3,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumn',
        order: 2
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":2,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2.5,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2.5,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2.5,\"sectionIndex\":3,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":3,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":3,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a Vertical section at the end to an uncustomized page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'Vertical'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":2,\"isLayoutReflowOnTop\":false,\"controlIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a Vertical section at the end with correct zoneEmphasisValue to an uncustomized page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'Vertical',
        zoneEmphasis: 'Neutral'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":2,\"isLayoutReflowOnTop\":false,\"controlIndex\":1},\"emphasis\":{\"zoneEmphasis\":1}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a Vertical section at the end with correct zoneEmphasisValue and isLayoutReflowOnTop values to an uncustomized page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'Vertical',
        zoneEmphasis: 'Neutral',
        isLayoutReflowOnTop: true
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":2,\"isLayoutReflowOnTop\":true,\"controlIndex\":1},\"emphasis\":{\"zoneEmphasis\":1}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a OneColumn section at the end to an uncustomized page with Image zoneEmphasis', async () => {
    let newZoneId = '';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        newZoneId = JSON.parse(opts.data.CanvasContent1)[1].position.zoneId;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        zoneEmphasis: 'Image',
        imageUrl: 'https://contoso.com/image.jpg'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": `[{"controlType":0,"pageSettingsSlice":{"isDefaultDescription":true,"isDefaultThumbnail":true}},{"displayMode":2,"position":{"zoneIndex":1,"sectionIndex":1,"sectionFactor":12,"layoutIndex":1,"zoneId":"${newZoneId}"},"emphasis":{}},{"controlType":14,"webPartData":{"properties":{"zoneBackground":{"${newZoneId}":{"type":"image","imageData":{"source":2,"fileName":"sectionbackground.jpg","height":955,"width":555},"fillMode":0,"useLightText":false,"overlay":{"color":"#FFFFFF","opacity":60}}}},"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"zoneBackground.${newZoneId}.imageData.url":"https://contoso.com/image.jpg"},"links":{}},"dataVersion":"1.0"}}]` });
  });

  it('adds a OneColumn section at the end to an uncustomized page with Gradient zoneEmphasis', async () => {
    let newZoneId = '';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        newZoneId = JSON.parse(opts.data.CanvasContent1)[1].position.zoneId;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        zoneEmphasis: 'Gradient',
        gradientText: 'test gradient'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": `[{"controlType":0,"pageSettingsSlice":{"isDefaultDescription":true,"isDefaultThumbnail":true}},{"displayMode":2,"position":{"zoneIndex":1,"sectionIndex":1,"sectionFactor":12,"layoutIndex":1,"zoneId":"${newZoneId}"},"emphasis":{}},{"controlType":14,"webPartData":{"properties":{"zoneBackground":{"${newZoneId}":{"type":"gradient","gradient":"test gradient","useLightText":false,"overlay":{"color":"#FFFFFF","opacity":60}}}},"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.0"}}]` });
  });

  it('adds a OneColumn section at the end to an uncustomized page with Image zoneEmphasis and all options available', async () => {
    let newZoneId = '';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        newZoneId = JSON.parse(opts.data.CanvasContent1)[1].position.zoneId;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        zoneEmphasis: 'Image',
        imageUrl: 'https://contoso.com/image.jpg',
        imageHeight: 100,
        imageWidth: 200,
        fillMode: 'ScaleToFill',
        useLightText: true,
        overlayColor: '#FF00FF',
        overlayOpacity: 50
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": `[{"controlType":0,"pageSettingsSlice":{"isDefaultDescription":true,"isDefaultThumbnail":true}},{"displayMode":2,"position":{"zoneIndex":1,"sectionIndex":1,"sectionFactor":12,"layoutIndex":1,"zoneId":"${newZoneId}"},"emphasis":{}},{"controlType":14,"webPartData":{"properties":{"zoneBackground":{"${newZoneId}":{"type":"image","imageData":{"source":2,"fileName":"sectionbackground.jpg","height":100,"width":200},"fillMode":0,"useLightText":true,"overlay":{"color":"#FF00FF","opacity":50}}}},"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"zoneBackground.${newZoneId}.imageData.url":"https://contoso.com/image.jpg"},"links":{}},"dataVersion":"1.0"}}]` });
  });

  it('adds a OneColumn section at the end to an uncustomized page with Gradient zoneEmphasis and all options available', async () => {
    let newZoneId = '';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        newZoneId = JSON.parse(opts.data.CanvasContent1)[1].position.zoneId;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        zoneEmphasis: 'Gradient',
        gradientText: 'test gradient',
        overlayColor: '#FF00FF',
        overlayOpacity: 50
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": `[{"controlType":0,"pageSettingsSlice":{"isDefaultDescription":true,"isDefaultThumbnail":true}},{"displayMode":2,"position":{"zoneIndex":1,"sectionIndex":1,"sectionFactor":12,"layoutIndex":1,"zoneId":"${newZoneId}"},"emphasis":{}},{"controlType":14,"webPartData":{"properties":{"zoneBackground":{"${newZoneId}":{"type":"gradient","gradient":"test gradient","useLightText":false,"overlay":{"color":"#FF00FF","opacity":50}}}},"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.0"}}]` });
  });

  it('adds a OneColumn section at the end to a page with background section added with Image zoneEmphasis', async () => {
    let newZoneId = '';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"sectionIndex\":1,\"controlIndex\":1,\"sectionFactor\":6,\"zoneId\":\"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35\"},\"id\":\"emptySection\",\"addedFromPersistedData\":true},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"sectionIndex\":2,\"controlIndex\":1,\"sectionFactor\":6,\"zoneId\":\"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35\"},\"id\":\"emptySection\",\"addedFromPersistedData\":true},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"sectionIndex\":1,\"controlIndex\":1,\"sectionFactor\":12,\"zoneId\":\"931e6d64-c667-4e2e-b678-eab508d511c8\"},\"id\":\"emptySection\",\"addedFromPersistedData\":true},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true,\"globalRichTextStylingVersion\":0,\"rtePageSettings\":{\"contentVersion\":4},\"isEmailReady\":false}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35\":{\"type\":\"gradient\",\"gradient\":\"radial-gradient(53.89% 99.37% at 39.45% -6.02%, rgba(4, 110, 212, 0.8) 0%, rgba(4, 110, 212, 0) 100%),\\n      radial-gradient(47.01% 82.21% at 104.3% 15.51%, rgba(118, 5, 180, 0.5) 0%, rgba(118, 5, 180, 0) 100%),\\n      radial-gradient(56.12% 58.33% at 50% 131.71%, #7605B4 34.7%, rgba(118, 5, 180, 0) 100%),\\n      linear-gradient(0deg, #110739, #110739)\",\"useLightText\":true,\"overlay\":{\"color\":\"#000000\",\"opacity\":35}},\"931e6d64-c667-4e2e-b678-eab508d511c8\":{\"type\":\"image\",\"imageData\":{\"source\":1,\"fileName\":\"sectionbackgroundimagedark3.jpg\",\"height\":955,\"width\":555},\"overlay\":{\"color\":\"#000000\",\"opacity\":60},\"useLightText\":true}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.931e6d64-c667-4e2e-b678-eab508d511c8.imageData.url\":\"/_layouts/15/images/sectionbackgroundimagedark3.jpg\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        newZoneId = JSON.parse(opts.data.CanvasContent1)[4].position.zoneId;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        zoneEmphasis: 'Image',
        imageUrl: 'https://contoso.com/image.jpg'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": `[{"position":{"layoutIndex":1,"zoneIndex":2,"sectionIndex":1,"controlIndex":1,"sectionFactor":6,"zoneId":"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35"},"id":"emptySection","addedFromPersistedData":true},{"position":{"layoutIndex":1,"zoneIndex":2,"sectionIndex":2,"controlIndex":1,"sectionFactor":6,"zoneId":"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35"},"id":"emptySection","addedFromPersistedData":true},{"position":{"layoutIndex":1,"zoneIndex":3,"sectionIndex":1,"controlIndex":1,"sectionFactor":12,"zoneId":"931e6d64-c667-4e2e-b678-eab508d511c8"},"id":"emptySection","addedFromPersistedData":true},{"controlType":0,"pageSettingsSlice":{"isDefaultDescription":true,"isDefaultThumbnail":true,"globalRichTextStylingVersion":0,"rtePageSettings":{"contentVersion":4},"isEmailReady":false}},{"displayMode":2,"position":{"zoneIndex":4,"sectionIndex":1,"sectionFactor":12,"layoutIndex":1,"zoneId":"${newZoneId}"},"emphasis":{}},{"controlType":14,"webPartData":{"properties":{"zoneBackground":{"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35":{"type":"gradient","gradient":"radial-gradient(53.89% 99.37% at 39.45% -6.02%, rgba(4, 110, 212, 0.8) 0%, rgba(4, 110, 212, 0) 100%),\\n      radial-gradient(47.01% 82.21% at 104.3% 15.51%, rgba(118, 5, 180, 0.5) 0%, rgba(118, 5, 180, 0) 100%),\\n      radial-gradient(56.12% 58.33% at 50% 131.71%, #7605B4 34.7%, rgba(118, 5, 180, 0) 100%),\\n      linear-gradient(0deg, #110739, #110739)","useLightText":true,"overlay":{"color":"#000000","opacity":35}},"931e6d64-c667-4e2e-b678-eab508d511c8":{"type":"image","imageData":{"source":1,"fileName":"sectionbackgroundimagedark3.jpg","height":955,"width":555},"overlay":{"color":"#000000","opacity":60},"useLightText":true},"${newZoneId}":{"type":"image","imageData":{"source":2,"fileName":"sectionbackground.jpg","height":955,"width":555},"fillMode":0,"useLightText":false,"overlay":{"color":"#FFFFFF","opacity":60}}}},"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"zoneBackground.931e6d64-c667-4e2e-b678-eab508d511c8.imageData.url":"/_layouts/15/images/sectionbackgroundimagedark3.jpg","zoneBackground.${newZoneId}.imageData.url":"https://contoso.com/image.jpg"},"links":{}},"dataVersion":"1.0"}}]` });
  });

  it('adds a OneColumn section at the end to a page with background section added with Gradient zoneEmphasis', async () => {
    let newZoneId = '';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"sectionIndex\":1,\"controlIndex\":1,\"sectionFactor\":6,\"zoneId\":\"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35\"},\"id\":\"emptySection\",\"addedFromPersistedData\":true},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"sectionIndex\":2,\"controlIndex\":1,\"sectionFactor\":6,\"zoneId\":\"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35\"},\"id\":\"emptySection\",\"addedFromPersistedData\":true},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"sectionIndex\":1,\"controlIndex\":1,\"sectionFactor\":12,\"zoneId\":\"931e6d64-c667-4e2e-b678-eab508d511c8\"},\"id\":\"emptySection\",\"addedFromPersistedData\":true},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true,\"globalRichTextStylingVersion\":0,\"rtePageSettings\":{\"contentVersion\":4},\"isEmailReady\":false}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35\":{\"type\":\"gradient\",\"gradient\":\"radial-gradient(53.89% 99.37% at 39.45% -6.02%, rgba(4, 110, 212, 0.8) 0%, rgba(4, 110, 212, 0) 100%),\\n      radial-gradient(47.01% 82.21% at 104.3% 15.51%, rgba(118, 5, 180, 0.5) 0%, rgba(118, 5, 180, 0) 100%),\\n      radial-gradient(56.12% 58.33% at 50% 131.71%, #7605B4 34.7%, rgba(118, 5, 180, 0) 100%),\\n      linear-gradient(0deg, #110739, #110739)\",\"useLightText\":true,\"overlay\":{\"color\":\"#000000\",\"opacity\":35}},\"931e6d64-c667-4e2e-b678-eab508d511c8\":{\"type\":\"image\",\"imageData\":{\"source\":1,\"fileName\":\"sectionbackgroundimagedark3.jpg\",\"height\":955,\"width\":555},\"overlay\":{\"color\":\"#000000\",\"opacity\":60},\"useLightText\":true}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.931e6d64-c667-4e2e-b678-eab508d511c8.imageData.url\":\"/_layouts/15/images/sectionbackgroundimagedark3.jpg\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        newZoneId = JSON.parse(opts.data.CanvasContent1)[4].position.zoneId;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        zoneEmphasis: 'Gradient',
        gradientText: 'test gradient'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": `[{"position":{"layoutIndex":1,"zoneIndex":2,"sectionIndex":1,"controlIndex":1,"sectionFactor":6,"zoneId":"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35"},"id":"emptySection","addedFromPersistedData":true},{"position":{"layoutIndex":1,"zoneIndex":2,"sectionIndex":2,"controlIndex":1,"sectionFactor":6,"zoneId":"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35"},"id":"emptySection","addedFromPersistedData":true},{"position":{"layoutIndex":1,"zoneIndex":3,"sectionIndex":1,"controlIndex":1,"sectionFactor":12,"zoneId":"931e6d64-c667-4e2e-b678-eab508d511c8"},"id":"emptySection","addedFromPersistedData":true},{"controlType":0,"pageSettingsSlice":{"isDefaultDescription":true,"isDefaultThumbnail":true,"globalRichTextStylingVersion":0,"rtePageSettings":{"contentVersion":4},"isEmailReady":false}},{"displayMode":2,"position":{"zoneIndex":4,"sectionIndex":1,"sectionFactor":12,"layoutIndex":1,"zoneId":"${newZoneId}"},"emphasis":{}},{"controlType":14,"webPartData":{"properties":{"zoneBackground":{"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35":{"type":"gradient","gradient":"radial-gradient(53.89% 99.37% at 39.45% -6.02%, rgba(4, 110, 212, 0.8) 0%, rgba(4, 110, 212, 0) 100%),\\n      radial-gradient(47.01% 82.21% at 104.3% 15.51%, rgba(118, 5, 180, 0.5) 0%, rgba(118, 5, 180, 0) 100%),\\n      radial-gradient(56.12% 58.33% at 50% 131.71%, #7605B4 34.7%, rgba(118, 5, 180, 0) 100%),\\n      linear-gradient(0deg, #110739, #110739)","useLightText":true,"overlay":{"color":"#000000","opacity":35}},"931e6d64-c667-4e2e-b678-eab508d511c8":{"type":"image","imageData":{"source":1,"fileName":"sectionbackgroundimagedark3.jpg","height":955,"width":555},"overlay":{"color":"#000000","opacity":60},"useLightText":true},"${newZoneId}":{"type":"gradient","gradient":"test gradient","useLightText":false,"overlay":{"color":"#FFFFFF","opacity":60}}}},"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{"zoneBackground.931e6d64-c667-4e2e-b678-eab508d511c8.imageData.url":"/_layouts/15/images/sectionbackgroundimagedark3.jpg"},"links":{}},"dataVersion":"1.0"}}]` });
  });

  it('adds a OneColumn section at the end to an uncustomized page with collapsible setting', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        isCollapsibleSection: true,
        iconAlignment: 'Right'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{},\"zoneGroupMetadata\":{\"type\":1,\"isExpanded\":false,\"showDividerLine\":false,\"iconAlignment\":\"right\",\"headingLevel\":2}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a OneColumn section at the end to an uncustomized page with collapsible setting and left iconAlignment', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        isCollapsibleSection: true,
        iconAlignment: 'Left'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{},\"zoneGroupMetadata\":{\"type\":1,\"isExpanded\":false,\"showDividerLine\":false,\"iconAlignment\":\"left\",\"headingLevel\":2}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a OneColumn section at the end to an uncustomized page with collapsible setting and section title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        isCollapsibleSection: true,
        iconAlignment: 'Right',
        collapsibleTitle: 'Collapsible section title'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{},\"zoneGroupMetadata\":{\"type\":1,\"isExpanded\":false,\"showDividerLine\":false,\"iconAlignment\":\"right\",\"displayName\":\"Collapsible section title\",\"headingLevel\":2}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a OneColumn section at the end to an uncustomized page with collapsible setting and headingLevel', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        isCollapsibleSection: true,
        headingLevel: 3
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{},\"zoneGroupMetadata\":{\"type\":1,\"isExpanded\":false,\"showDividerLine\":false,\"iconAlignment\":\"left\",\"headingLevel\":3}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a Flexible section at the end to an uncustomized page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'Flexible'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":100,\"layoutIndex\":1},\"emphasis\":{},\"zoneReflowStrategy\":{\"axis\":0}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('adds a Flexible section at the end to an uncustomized page with zoneHeight and zoneReflowStrategy parameters', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'Flexible',
        zoneHeight: 500,
        zoneReflowStrategy: 'TopToBottom'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":100,\"layoutIndex\":1},\"emphasis\":{},\"zoneReflowStrategy\":{\"axis\":0},\"zoneHeight\":500},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" });
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumn',
        order: 2
      })
    }), new CommandError('An error has occurred'));
  });

  it('fails when trying to add a vertical section while a full-width section is already on the page.', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"sectionIndex\":1,\"controlIndex\":1,\"sectionFactor\":0,\"zoneId\":\"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35\"},\"id\":\"emptySection\",\"addedFromPersistedData\":true},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'Vertical'
      })
    }), new CommandError('A full-width section already exists on the page. A vertical section cannot be added to a page that already has a full-width section.'));
  });

  it('fails when trying to add a full-width section while a vertical section is already on the page.', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"position\":{\"layoutIndex\":2,\"zoneIndex\":1,\"sectionIndex\":1,\"controlIndex\":1,\"sectionFactor\":12,\"zoneId\":\"0158a0e8-20ad-4d8d-9cdc-6e1fde815a35\"},\"id\":\"emptySection\",\"addedFromPersistedData\":true},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/newsletter/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options:
      commandOptionsSchema.parse({
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumnFullWidth'
      })
    }), new CommandError('A vertical section already exists on the page. A full-width section cannot be added to a page that already has a vertical section.'));
  });

  it('fails validation if order has invalid (negative) value', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      order: -1,
      sectionTemplate: 'OneColumn'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'The value of parameter order must be 1 or higher');
  });

  it('fails validation if order has invalid (non number) value', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      order: 'abc',
      sectionTemplate: 'OneColumn'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if sectionTemplate is not valid', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      order: 'abc',
      sectionTemplate: 'OneColumnInvalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if webUrl is not valid', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      order: 1,
      sectionTemplate: 'OneColumn',
      webUrl: 'http://notasharepointurl'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'http://notasharepointurl' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation if zoneEmphasis is not valid', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      zoneEmphasis: 'Invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if isLayoutReflowOnTop is valid but sectionTemplate is not Vertical', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      isLayoutReflowOnTop: true
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if iconAlignment is not valid', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      zoneEmphasis: 'Image',
      imageUrl: 'https://contoso.com/image.jpg',
      iconAlignment: 'Invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if fillMode is not valid', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      zoneEmphasis: 'Image',
      imageUrl: 'https://contoso.com/image.jpg',
      fillMode: 'Invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if imageUrl is specified but zoneEmphasis is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      imageUrl: 'test.png'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if gradientText is specified but zoneEmphasis is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      gradientText: 'test gradient'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if overlayOpacity is not valid', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      zoneEmphasis: 'Image',
      imageUrl: 'https://contoso.com/image.jpg',
      overlayOpacity: 100001
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if overlayColor is not valid', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      zoneEmphasis: 'Image',
      imageUrl: 'https://contoso.com/image.jpg',
      overlayColor: "InvalidColor"
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if overlayColor is specified but is not Image or Gradient zoneEmphasis', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      zoneEmphasis: 'Strong',
      overlayColor: "#FFFFFF"
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if zoneEmphasis is Image and imageUrl is not defined', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      zoneEmphasis: 'Image'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if zoneEmphasis is Gradient and gradientText is not defined', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      zoneEmphasis: 'Gradient'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if zoneReflowStrategy is valid but sectionTemplate is not Flexible', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      zoneReflowStrategy: 'LeftToRight'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if zoneHeight is valid but sectionTemplate is not Flexible', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      zoneHeight: 300
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if zoneReflowStrategy is not valid and sectionTemplate is Flexible', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'Flexible',
      zoneReflowStrategy: 'Invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if zoneHeight is not valid and sectionTemplate is Flexible', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'Flexible',
      zoneHeight: -1
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'The value of parameter zoneHeight must be 34 or higher');
  });

  it('fails validation if headingLevel is not valid', () => {
    const actual = commandOptionsSchema.safeParse({
      pageName: 'page.aspx',
      webUrl: 'https://contoso.sharepoint.com',
      sectionTemplate: 'OneColumn',
      isCollapsibleSection: true,
      headingLevel: 6
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'The value of parameter headingLevel must be 2, 3 or 4');
  });


  it('passes validation if all the parameters are specified for a regular Section', () => {
    const actual = commandOptionsSchema.safeParse({
      order: 1,
      sectionTemplate: 'OneColumn',
      webUrl: 'https://contoso.sharepoint.com',
      pageName: 'Home.aspx',
      zoneEmphasis: 'None'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if all the parameters are specified for Vertical Section', () => {
    const actual = commandOptionsSchema.safeParse({
      order: 1,
      sectionTemplate: 'Vertical',
      webUrl: 'https://contoso.sharepoint.com',
      pageName: 'Home.aspx',
      zoneEmphasis: 'None',
      isLayoutReflowOnTop: false
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if order, zoneEmphasis and isLayoutReflowOnTop are not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      sectionTemplate: 'OneColumn',
      webUrl: 'https://contoso.sharepoint.com',
      pageName: 'Home.aspx'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if order and isLayoutReflowOnTop are not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      sectionTemplate: 'OneColumn',
      webUrl: 'https://contoso.sharepoint.com',
      pageName: 'Home.aspx',
      zoneEmphasis: 'None'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if isLayoutReflowOnTop is specified along with Vertical sectionTemplate', () => {
    const actual = commandOptionsSchema.safeParse({
      sectionTemplate: 'Vertical',
      webUrl: 'https://contoso.sharepoint.com',
      pageName: 'Home.aspx',
      zoneEmphasis: 'None',
      order: 1,
      isLayoutReflowOnTop: false
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if order is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      sectionTemplate: 'Vertical',
      webUrl: 'https://contoso.sharepoint.com',
      pageName: 'Home.aspx',
      isLayoutReflowOnTop: false
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if zoneEmphasis is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      sectionTemplate: 'Vertical',
      webUrl: 'https://contoso.sharepoint.com',
      pageName: 'Home.aspx',
      order: 1,
      isLayoutReflowOnTop: false
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      sectionTemplate: 'OneColumn',
      webUrl: 'https://contoso.sharepoint.com',
      pageName: 'Home.aspx',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
