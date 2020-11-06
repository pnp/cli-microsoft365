import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import { ClientSidePage } from './clientsidepages';
import { mockPage, mockPageData, mockPageDataFail } from './page-control-set.mock';
const command: Command = require('./page-control-set');

describe(commands.PAGE_CONTROL_SET, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post,
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

  // NAME and DESCRIPTION

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_CONTROL_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  // VALIDATE FUNCTIONALITY

  it('correctly handles control not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(mockPage);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e6' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Control with ID 3ede60d3-dc2c-438b-b5bf-cc40bb2351e6 not found on page home.aspx")));
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles control found and handles error on page checkout error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(mockPage);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } }, (err) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles control found and handles page checkout correctly when no data is provided', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(mockPage);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      const checkOutPostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`;

      if ((opts.url as string).indexOf(checkOutPostUrl) > -1) {
        return Promise.resolve(null);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Page home.aspx information not retrieved with the checkout')));
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles control not found after the page has been checked out', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(mockPage);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      const checkOutPostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`;

      if ((opts.url as string).indexOf(checkOutPostUrl) > -1) {
        return Promise.resolve(mockPageDataFail);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Control with ID ede2ee65-157d-4523-b4ed-87b9b64374a6 not found on page home.aspx')));
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles control found and handles page checkout', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(mockPage);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      const checkOutPostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`;

      if ((opts.url as string).indexOf(checkOutPostUrl) > -1) {
        return Promise.resolve(mockPageData);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } }, (err) => {
      try {
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('correctly page save with webPartData', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(mockPage);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      const checkOutPostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`;
      const savePagePostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`;

      if ((opts.url as string).indexOf(checkOutPostUrl) > -1) {
        return Promise.resolve(mockPageData);
      }

      if ((opts.url as string).indexOf(savePagePostUrl) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6', webPartData: '{}' } }, (err?: any) => {
      try {
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('correctly page save with webPartProperties', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(mockPage);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      const checkOutPostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`;
      const savePagePostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`;

      if ((opts.url as string).indexOf(checkOutPostUrl) > -1) {
        return Promise.resolve(mockPageData);
      }

      if ((opts.url as string).indexOf(savePagePostUrl) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6', webPartProperties: '{}' } }, (err?: any) => {
      try {
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when retrieving pages', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    command.action(logger, { options: { debug: true, id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', webPartData: "{}", webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  // OPTIONS

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  // VALIDATION

  it('fails validation if the specified id is not a valid GUID', () => {
    const actual = command.validate({ options: { id: 'abc', name: 'home.aspx', webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified webPartProperties is not a valid JSON string', () => {
    const actual = command.validate({ options: { webPartProperties: "abc", id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', name: 'home.aspx', webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified webPartData is not a valid JSON string', () => {
    const actual = command.validate({ options: { webPartData: "abc", id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', name: 'home.aspx', webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webPartData and webPartProperties options are provided', () => {
    const actual = command.validate({ options: { webPartProperties: "{}", webPartData: "{}", id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', webUrl: 'foo', name: 'home.aspx' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', webUrl: 'foo', name: 'home.aspx' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when right properties with webPartData are provided', () => {
    const actual = command.validate({ options: { id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', webPartData: "{}", webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when right properties with webPartProperties are provided', () => {
    const actual = command.validate({ options: { id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', webPartProperties: "{}", webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx' } });
    assert.strictEqual(actual, true);
  });
});