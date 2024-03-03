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
import command from './page-section-add.js';

describe(commands.PAGE_SECTION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": false,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`) > -1) {
        checkedOut = true;
        return {};
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        pageName: 'home',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      }
    });
    assert.deepEqual(checkedOut, true);
  });

  it('doesn\'t check out page if not checked out by the current user', async () => {
    let checkingOut = false;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`) > -1) {
        checkingOut = true;
        return {};
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      }
    });
    assert.deepEqual(checkingOut, false);
  });

  it('adds a first section to an uncustomized page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a first section to an uncustomized page with order set to 1', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        order: 1
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a first section to an uncustomized page correctly even when CanvasContent1 of returned page is null', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": null
        });
      }

      return Promise.reject('Invalid request');
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a first section to the page if no order specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a first section to the page if order 1 specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumnFullWidth',
        order: 1
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":0,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a section to the beginning of the page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnLeft',
        order: 1
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a section to the end of the page when order not specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnRight'
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a section to the end of the page when order set to last section', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnRight',
        order: 2
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a section to the end of the page when order is larger than the last section', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnRight',
        order: 5
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a section between two other sections', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'ThreeColumn',
        order: 2
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":3,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a section between two other sections (2)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return {
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":2,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":3,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        };
      }

      throw 'Invalid request';
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumn',
        order: 2
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":2,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.75,\"sectionIndex\":1,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.75,\"sectionIndex\":2,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":3,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a Vertical section at the end to an uncustomized page', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'Vertical'
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":2,\"isLayoutReflowOnTop\":false,\"controlIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });


  it('adds a Vertical section at the end with correct zoneEmphasisValue to an uncustomized page', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'Vertical',
        zoneEmphasis: 'Neutral'
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":2,\"isLayoutReflowOnTop\":false,\"controlIndex\":1},\"emphasis\":{\"zoneEmphasis\":1}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('adds a Vertical section at the end with correct zoneEmphasisValue and isLayoutReflowOnTop values to an uncustomized page', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    let data: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        data = JSON.stringify(opts.data);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'Vertical',
        zoneEmphasis: 'Neutral',
        isLayoutReflowOnTop: true
      }
    });
    assert.strictEqual(data, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":2,\"isLayoutReflowOnTop\":true,\"controlIndex\":1},\"emphasis\":{\"zoneEmphasis\":1}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, {
      options:
      {
        pageName: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumn',
        order: 2
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if order has invalid (negative) value', async () => {
    const actual = await command.validate({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        order: -1,
        sectionTemplate: 'OneColumn'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if order has invalid (non number) value', async () => {
    const actual = await command.validate({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        order: 'abc',
        sectionTemplate: 'OneColumn'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if sectionTemplate is not valid', async () => {
    const actual = await command.validate({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        order: 'abc',
        sectionTemplate: 'OneColumnInvalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not valid', async () => {
    const actual = await command.validate({
      options: {
        pageName: 'page.aspx',
        order: 1,
        sectionTemplate: 'OneColumn',
        webUrl: 'http://notasharepointurl'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if zoneEmphasis is not valid', async () => {
    const actual = await command.validate({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        sectionTemplate: 'OneColumn',
        zoneEmphasis: 'Invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });


  it('fails validation if isLayoutReflowOnTop is valid but sectionTemplate is not Vertical', async () => {
    const actual = await command.validate({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        sectionTemplate: 'OneColumn',
        isLayoutReflowOnTop: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all the parameters are specified for a regular Section', async () => {
    const actual = await command.validate({
      options: {
        order: 1,
        sectionTemplate: 'OneColumn',
        webUrl: 'https://contoso.sharepoint.com',
        pageName: 'Home.aspx',
        zoneEmphasis: 'None'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if all the parameters are specified for Vertical Section', async () => {
    const actual = await command.validate({
      options: {
        order: 1,
        sectionTemplate: 'Vertical',
        webUrl: 'https://contoso.sharepoint.com',
        pageName: 'Home.aspx',
        zoneEmphasis: 'None',
        isLayoutReflowOnTop: false
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if order, zoneEmphasis and isLayoutReflowOnTop are not specified', async () => {
    const actual = await command.validate({
      options: {
        sectionTemplate: 'OneColumn',
        webUrl: 'https://contoso.sharepoint.com',
        pageName: 'Home.aspx'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if order and isLayoutReflowOnTop are not specified', async () => {
    const actual = await command.validate({
      options: {
        sectionTemplate: 'OneColumn',
        webUrl: 'https://contoso.sharepoint.com',
        pageName: 'Home.aspx',
        zoneEmphasis: 'None'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if isLayoutReflowOnTop is specified along with Vertical sectionTemplate', async () => {
    const actual = await command.validate({
      options: {
        sectionTemplate: 'Vertical',
        webUrl: 'https://contoso.sharepoint.com',
        pageName: 'Home.aspx',
        zoneEmphasis: 'None',
        order: 1,
        isLayoutReflowOnTop: false
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if order is not specified', async () => {
    const actual = await command.validate({
      options: {
        sectionTemplate: 'Vertical',
        webUrl: 'https://contoso.sharepoint.com',
        pageName: 'Home.aspx',
        isLayoutReflowOnTop: false
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if zoneEmphasis is not specified', async () => {
    const actual = await command.validate({
      options: {
        sectionTemplate: 'Vertical',
        webUrl: 'https://contoso.sharepoint.com',
        pageName: 'Home.aspx',
        order: 1,
        isLayoutReflowOnTop: false
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });


  it('supports specifying page name', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--pageName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying sectionTemplate', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--sectionTemplate') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying order', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--order') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying zoneEmphasis', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--zoneEmphasis') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying isLayoutReflowOnTop', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--isLayoutReflowOnTop') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
