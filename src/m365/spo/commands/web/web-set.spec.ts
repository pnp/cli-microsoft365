import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./web-set');

describe(commands.WEB_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.WEB_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates site title', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        Title: 'New title'
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', title: 'New title' } });
  });

  it('updates site logo URL', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SiteLogoUrl: 'image.png'
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', siteLogoUrl: 'image.png' } });
  });

  it('unsets the site logo', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SiteLogoUrl: ''
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', siteLogoUrl: '' } });
  });

  it('disables quick launch', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        QuickLaunchEnabled: false
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'false' } });
  });

  it('enables quick launch', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        QuickLaunchEnabled: true
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'true' } });
  });

  it('sets site header to compact', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderLayout: 2
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'compact' } });
  });

  it('sets site header to standard', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderLayout: 1
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'standard' } });
  });

  it('sets site header emphasis to 0', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderEmphasis: 0
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 0 } });
  });

  it('sets site header emphasis to 1', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderEmphasis: 1
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 1 } });
  });

  it('sets site header emphasis to 2', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderEmphasis: 2
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 2 } });
  });

  it('sets site header emphasis to 3', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderEmphasis: 3
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 3 } });
  });

  it('sets site menu mode to megamenu', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        MegaMenuEnabled: true
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'true' } });
  });

  it('sets site menu mode to cascading', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        MegaMenuEnabled: false
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'false' } });
  });

  it('updates all properties', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/web') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', title: 'New title', description: 'New description', siteLogoUrl: 'image.png', quickLaunchEnabled: 'true', headerLayout: 'compact', headerEmphasis: 1, megaMenuEnabled: 'true', footerEnabled: 'true' } });
  });

  it('Update Welcome page', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/web') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, welcomePage: 'SitePages/Home.aspx', url: 'https://contoso.sharepoint.com/sites/team-a' } });
  });

  it('Update Welcome page (debug)', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/web') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, welcomePage: 'SitePages/Home.aspx', url: 'https://contoso.sharepoint.com/sites/team-a' } });
  });

  it('correctly handles error when hub site not found', async () => {
    sinon.stub(request, 'patch').callsFake(() => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-1, Microsoft.SharePoint.Client.ResourceNotFoundException",
            "message": {
              "lang": "en-US",
              "value": "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
            }
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a' } } as any), new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."));
  });

  it('correctly handles error while updating Welcome page', async () => {
    sinon.stub(request, 'patch').callsFake(() => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-1, Microsoft.SharePoint.Client.ResourceNotFoundException",
            "message": {
              "lang": "en-US",
              "value": "The WelcomePage property must be a path that is relative to the folder, and the path cannot contain two consecutive periods (..)."
            }
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: {
      debug: false, 
      welcomePage: 'https://contoso.sharepoint.com/sites/team-a/SitePages/Home.aspx', 
      url: 'https://contoso.sharepoint.com/sites/team-a' } } as any), new CommandError('The WelcomePage property must be a path that is relative to the folder, and the path cannot contain two consecutive periods (..).'));
  });

  it('allows unknown properties', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
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

  it('supports specifying url', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--url') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if url is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the url is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if quickLaunchEnabled is not a valid boolean', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the url is a valid SharePoint URL and quickLaunch set to "true"', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if headerLayout is invalid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if headerLayout is set to standard', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'standard' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if headerLayout is set to compact', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'compact' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if headerEmphasis is not a number', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if headerEmphasis is out of bounds', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 4 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if headerEmphasis is 0', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 0 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if headerEmphasis is 1', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 1 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if headerEmphasis is 2', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 2 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if headerEmphasis is 3', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 3 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if megaMenuEnabled is not a valid boolean', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if megaMenuEnabled is set to true', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if megaMenuEnabled is set to false', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'false' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if footerEnabled is not a valid boolean', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if footerEnabled is set to true', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if footerEnabled is set to false', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'false' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('enables footer', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        FooterEnabled: true
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'true' } });
  });

  it('disables footer', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        FooterEnabled: false
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'false' } });
  });

  it('fails validation if search scope is not valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if search scope is set to defaultscope', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'defaultscope' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if search scope is set to tenant', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'tenant' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if search scope is set to hub', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'hub' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if search scope is set to site', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'site' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation even if search scope is not all lower case', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'DefaultScope' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if search scope passed is a number', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 2 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('sets search scope to default scope', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SearchScope: 0
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'defaultscope' } });
  });

  it('sets search scope to tenant', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SearchScope: 1
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'tenant' } });
  });

  it('sets search scope to hub', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SearchScope: 2
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'hub' } });
  });

  it('sets search scope to site', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SearchScope: 3
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'site' } });
  });

  it('sets search scope even if parameter is not all lower case', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SearchScope: 3
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'Site' } });
  });
});