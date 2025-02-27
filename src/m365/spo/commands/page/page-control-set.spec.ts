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
import { ClientSidePage } from './clientsidepages.js';
import command from './page-control-set.js';
import { mockCanvasContentStringified, mockPageData, mockPageDataFail } from './page-control-set.mock.js';

describe(commands.PAGE_CONTROL_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
    sinonUtil.restore([
      request.get,
      request.post,
      ClientSidePage.fromHtml
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  // NAME and DESCRIPTION

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_CONTROL_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  // VALIDATE FUNCTIONALITY

  it('correctly handles control not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return { CanvasContent1: mockCanvasContentStringified };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e6' } }), new CommandError("Control with ID 3ede60d3-dc2c-438b-b5bf-cc40bb2351e6 not found on page home.aspx"));
  });

  it('correctly handles control page with no Canvas Control content', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return {};
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e6' } }), new CommandError("Page home.aspx doesn't contain canvas controls."));
  });

  it('correctly handles control found and handles error on page checkout error', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return { CanvasContent1: mockCanvasContentStringified };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } }), new CommandError('An error has occurred'));
  });

  it('correctly handles control found and handles page checkout correctly when no data is provided', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return { CanvasContent1: mockCanvasContentStringified };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      const checkOutPostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`;

      if ((opts.url as string).indexOf(checkOutPostUrl) > -1) {
        return null;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } }), new CommandError('Page home.aspx information not retrieved with the checkout'));
  });

  it('correctly handles control not found after the page has been checked out', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return { CanvasContent1: mockCanvasContentStringified };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      const checkOutPostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`;

      if ((opts.url as string).indexOf(checkOutPostUrl) > -1) {
        return mockPageDataFail;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } }), new CommandError('Control with ID ede2ee65-157d-4523-b4ed-87b9b64374a6 not found on page home.aspx'));
  });

  it('correctly handles control found and handles page checkout', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return { CanvasContent1: mockCanvasContentStringified };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      const checkOutPostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`;

      if ((opts.url as string).indexOf(checkOutPostUrl) > -1) {
        return mockPageData;
      }

      if (opts.url?.endsWith(`_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/SavePageAsDraft`)) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } });
  });

  it('correctly page save with webPartData', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return { CanvasContent1: mockCanvasContentStringified };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      const checkOutPostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`;
      const savePagePostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`;

      if ((opts.url as string).indexOf(checkOutPostUrl) > -1) {
        return mockPageData;
      }

      if ((opts.url as string).indexOf(savePagePostUrl) > -1) {
        return {};
      }

      if (opts.url?.endsWith(`_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/SavePageAsDraft`)) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6', webPartData: '{}' } });
  });

  it('correctly page save with webPartProperties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return { CanvasContent1: mockCanvasContentStringified };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      const checkOutPostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`;
      const savePagePostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`;

      if ((opts.url as string).indexOf(checkOutPostUrl) > -1) {
        return mockPageData;
      }

      if ((opts.url as string).indexOf(savePagePostUrl) > -1) {
        return {};
      }

      if (opts.url?.endsWith(`_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/SavePageAsDraft`)) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6', webPartProperties: '{}' } });
  });

  it('correctly page save when page extension is not provided', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return { CanvasContent1: mockCanvasContentStringified };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      const checkOutPostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`;
      const savePagePostUrl = `_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`;

      if ((opts.url as string).indexOf(checkOutPostUrl) > -1) {
        return mockPageData;
      }

      if ((opts.url as string).indexOf(savePagePostUrl) > -1) {
        return {};
      }

      if (opts.url?.endsWith(`_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/SavePageAsDraft`)) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6', webPartProperties: '{}' } });
  });

  it('correctly handles OData error when retrieving pages', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    await assert.rejects(command.action(logger, { options: { debug: true, id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', webPartData: "{}", webUrl: 'https://contoso.sharepoint.com', pageName: 'home.aspx' } } as any),
      new CommandError('An error has occurred'));
  });

  // VALIDATION

  it('fails validation if the specified id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc', pageName: 'home.aspx', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified webPartProperties is not a valid JSON string', async () => {
    const actual = await command.validate({ options: { webPartProperties: "abc", id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', pageName: 'home.aspx', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified webPartData is not a valid JSON string', async () => {
    const actual = await command.validate({ options: { webPartData: "abc", id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', pageName: 'home.aspx', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webPartData and webPartProperties options are provided', async () => {
    const actual = await command.validate({ options: { webPartProperties: "{}", webPartData: "{}", id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', webUrl: 'foo', pageName: 'home.aspx' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', webUrl: 'foo', pageName: 'home.aspx' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when right properties with webPartData are provided', async () => {
    const actual = await command.validate({ options: { id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', webPartData: "{}", webUrl: 'https://contoso.sharepoint.com', pageName: 'home.aspx' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when right properties with webPartProperties are provided', async () => {
    const actual = await command.validate({ options: { id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e5', webPartProperties: "{}", webUrl: 'https://contoso.sharepoint.com', pageName: 'home.aspx' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
