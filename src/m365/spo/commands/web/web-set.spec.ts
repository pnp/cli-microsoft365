import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./web-set');

describe(commands.WEB_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.WEB_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates site title', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        Title: 'New title'
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', title: 'New title' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site logo URL', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SiteLogoUrl: 'image.png'
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', siteLogoUrl: 'image.png' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('unsets the site logo', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SiteLogoUrl: ''
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', siteLogoUrl: '' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disables quick launch', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        QuickLaunchEnabled: false
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'false' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('enables quick launch', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        QuickLaunchEnabled: true
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'true' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header to compact', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderLayout: 2
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'compact' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header to standard', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderLayout: 1
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'standard' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header emphasis to 0', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderEmphasis: 0
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 0 } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header emphasis to 1', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderEmphasis: 1
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 1 } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header emphasis to 2', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderEmphasis: 2
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 2 } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header emphasis to 3', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        HeaderEmphasis: 3
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 3 } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site menu mode to megamenu', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        MegaMenuEnabled: true
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'true' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site menu mode to cascading', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        MegaMenuEnabled: false
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'false' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates all properties', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        Title: 'New title',
        Description: 'New description',
        SiteLogoUrl: 'image.png',
        QuickLaunchEnabled: true,
        HeaderEmphasis: 2,
        HeaderLayout: 2,
        MegaMenuEnabled: true,
        FooterEnabled: true
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', title: 'New title', description: 'New description', siteLogoUrl: 'image.png', quickLaunchEnabled: 'true', headerLayout: 'compact', headerEmphasis: 1, megaMenuEnabled: 'true', footerEnabled: 'true' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Update Welcome page', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/web') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, welcomePage: 'SitePages/Home.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Update Welcome page (debug)', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/web') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, welcomePage: 'SitePages/Home.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when hub site not found', (done) => {
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

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error while updating Welcome page', (done) => {
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

    command.action(logger, { options: { debug: false, welcomePage: 'https://contoso.sharepoint.com/sites/team-a/SitePages/Home.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("The WelcomePage property must be a path that is relative to the folder, and the path cannot contain two consecutive periods (..).")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('supports specifying webUrl', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if quickLaunchEnabled is not a valid boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL and quickLaunch set to "true"', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if headerLayout is invalid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if headerLayout is set to standard', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'standard' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if headerLayout is set to compact', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'compact' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if headerEmphasis is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if headerEmphasis is out of bounds', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 4 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if headerEmphasis is 0', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 0 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if headerEmphasis is 1', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 1 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if headerEmphasis is 2', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 2 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if headerEmphasis is 3', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 3 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if megaMenuEnabled is not a valid boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if megaMenuEnabled is set to true', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if megaMenuEnabled is set to false', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'false' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if footerEnabled is not a valid boolean', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if footerEnabled is set to true', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if footerEnabled is set to false', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'false' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('enables footer', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        FooterEnabled: true
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'true' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disables footer', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        FooterEnabled: false
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'false' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if search scope is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if search scope is set to defaultscope', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'defaultscope' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if search scope is set to tenant', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'tenant' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if search scope is set to hub', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'hub' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if search scope is set to site', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'site' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation even if search scope is not all lower case', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'DefaultScope' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if search scope passed is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 2 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('sets search scope to default scope', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SearchScope: 0
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'defaultscope' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets search scope to tenant', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SearchScope: 1
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'tenant' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets search scope to hub', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SearchScope: 2
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'hub' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets search scope to site', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SearchScope: 3
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'site' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets search scope even if parameter is not all lower case', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.data) === JSON.stringify({
        SearchScope: 3
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'Site' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});