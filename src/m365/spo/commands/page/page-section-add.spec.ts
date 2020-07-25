import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./page-section-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.PAGE_SECTION_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  
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
  });

  afterEach(() => {
    Utils.restore([request.post, request.get]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_SECTION_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('checks out page if not checked out by the current user', (done) => {
    let checkedOut = false;
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": false,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`) > -1) {
        checkedOut = true;
        return Promise.resolve({});
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'home',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
      }
    }, () => {
      try {
        assert.deepEqual(checkedOut, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t check out page if not checked out by the current user', (done) => {
    let checkingOut = false;
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`) > -1) {
        checkingOut = true;
        return Promise.resolve({});
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      }
    }, () => {
      try {
        assert.deepEqual(checkingOut, false);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a first section to an uncustomized page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": null
        });
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        body = JSON.stringify(opts.body);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      }
    }, () => {
      try {
        assert.strictEqual(body, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a first section to an uncustomized page with order set to 1', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": null
        });
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        body = JSON.stringify(opts.body);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn',
        order: 1
      }
    }, () => {
      try {
        assert.strictEqual(body, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a first section to the page if no order specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        body = JSON.stringify(opts.body);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumn'
      }
    }, () => {
      try {
        assert.strictEqual(body, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a first section to the page if order 1 specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        body = JSON.stringify(opts.body);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'OneColumnFullWidth',
        order: 1
      }
    }, () => {
      try {
        assert.strictEqual(body, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":0,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a section to the beginning of the page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        body = JSON.stringify(opts.body);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnLeft',
        order: 1
      }
    }, () => {
      try {
        assert.strictEqual(body, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a section to the end of the page when order not specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        body = JSON.stringify(opts.body);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnRight'
      }
    }, () => {
      try {
        assert.strictEqual(body, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a section to the end of the page when order set to last section', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        body = JSON.stringify(opts.body);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnRight',
        order: 2
      }
    }, () => {
      try {
        assert.strictEqual(body, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a section to the end of the page when order is larger than the last section', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        body = JSON.stringify(opts.body);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumnRight',
        order: 5
      }
    }, () => {
      try {
        assert.strictEqual(body, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a section between two other sections', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        body = JSON.stringify(opts.body);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'ThreeColumn',
        order: 2
      }
    }, () => {
      try {
        assert.strictEqual(body, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":3,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a section between two other sections (2)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":2,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":3,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        body = JSON.stringify(opts.body);
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumn',
        order: 2
      }
    }, () => {
      try {
        assert.strictEqual(body, JSON.stringify({ "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.5,\"sectionIndex\":2,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.75,\"sectionIndex\":1,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":0.75,\"sectionIndex\":2,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1.5,\"sectionIndex\":3,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action({
      options:
      {
        name: 'home.aspx',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        sectionTemplate: 'TwoColumn',
        order: 2
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if order has invalid (negative) value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        name: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        order: -1,
        sectionTemplate: 'OneColumn'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if order has invalid (non number) value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        name: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        order: 'abc',
        sectionTemplate: 'OneColumn'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if sectionTemplate is not valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        name: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        order: 'abc',
        sectionTemplate: 'OneColumnInvalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        name: 'page.aspx',
        order: 1,
        sectionTemplate: 'OneColumn',
        webUrl: 'http://notasharepointurl'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all the parameters are specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        order: 1,
        sectionTemplate: 'OneColumn',
        webUrl: 'https://contoso.sharepoint.com',
        name: 'Home.aspx'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if order is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        sectionTemplate: 'OneColumn',
        webUrl: 'https://contoso.sharepoint.com',
        name: 'Home.aspx'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('supports specifying page name', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--name') > -1) {
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

  it('supports specifying sectionTemplate', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--sectionTemplate') > -1) {
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
});
