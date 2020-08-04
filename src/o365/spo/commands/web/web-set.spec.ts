import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./web-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.WEB_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.patch
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.WEB_SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('updates site title', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        Title: 'New title'
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', title: 'New title' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site description', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        Description: 'New description'
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', description: 'New description' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site logo URL', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        SiteLogoUrl: 'image.png'
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', siteLogoUrl: 'image.png' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disables quick launch', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        QuickLaunchEnabled: false
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'false' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('enables quick launch', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        QuickLaunchEnabled: true
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header to compact', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        HeaderLayout: 2
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'compact' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header to standard', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        HeaderLayout: 1
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'standard' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header emphasis to 0', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        HeaderEmphasis: 0
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 0 } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header emphasis to 1', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        HeaderEmphasis: 1
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 1 } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header emphasis to 2', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        HeaderEmphasis: 2
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 2 } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site header emphasis to 3', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        HeaderEmphasis: 3
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 3 } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site menu mode to megamenu', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        MegaMenuEnabled: true
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets site menu mode to cascading', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        MegaMenuEnabled: false
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'false' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates all properties', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', title: 'New title', description: 'New description', siteLogoUrl: 'image.png', quickLaunchEnabled: 'true', headerLayout: 'compact', headerEmphasis: 1, megaMenuEnabled: 'true', footerEnabled: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
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

    cmdInstance.action({ options: { debug: false, welcomePage: 'SitePages/Home.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
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

    cmdInstance.action({ options: { debug: true, welcomePage: 'SitePages/Home.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when hub site not found', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error while updating Welcome page', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
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

    cmdInstance.action({ options: { debug: false, welcomePage: 'https://contoso.sharepoint.com/sites/team-a/SitePages/Home.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a'} }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("The WelcomePage property must be a path that is relative to the folder, and the path cannot contain two consecutive periods (..).")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a' } });
    assert.equal(actual, true);
  });

  it('fails validation if quickLaunchEnabled is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', quickLaunchEnabled: 'true' } });
    assert.equal(actual, true);
  });

  it('fails validation if headerLayout is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if headerLayout is set to standard', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'standard' } });
    assert.equal(actual, true);
  });

  it('passes validation if headerLayout is set to compact', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerLayout: 'compact' } });
    assert.equal(actual, true);
  });

  it('fails validation if headerEmphasis is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if headerEmphasis is out of bounds', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 4 } });
    assert.notEqual(actual, true);
  });

  it('passes validation if headerEmphasis is 0', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 0 } });
    assert.equal(actual, true);
  });

  it('passes validation if headerEmphasis is 1', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 1 } });
    assert.equal(actual, true);
  });

  it('passes validation if headerEmphasis is 2', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 2 } });
    assert.equal(actual, true);
  });

  it('passes validation if headerEmphasis is 3', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', headerEmphasis: 3 } });
    assert.equal(actual, true);
  });

  it('fails validation if megaMenuEnabled is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if megaMenuEnabled is set to true', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'true' } });
    assert.equal(actual, true);
  });

  it('passes validation if megaMenuEnabled is set to false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', megaMenuEnabled: 'false' } });
    assert.equal(actual, true);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.WEB_SET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('fails validation if footerEnabled is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if footerEnabled is set to true', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'true' } });
    assert.equal(actual, true);
  });

  it('passes validation if footerEnabled is set to false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'false' } });
    assert.equal(actual, true);
  });

  it('enables footer', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        FooterEnabled: true
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disables footer', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        FooterEnabled: false
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', footerEnabled: 'false' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if search scope is not valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if search scope is set to defaultscope', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'defaultscope' } });
    assert.equal(actual, true);
  });

  it('passes validation if search scope is set to tenant', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'tenant' } });
    assert.equal(actual, true);
  });

  it('passes validation if search scope is set to hub', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'hub' } });
    assert.equal(actual, true);
  });

  it('passes validation if search scope is set to site', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'site' } });
    assert.equal(actual, true);
  });

  it('passes validation even if search scope is not all lower case', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'DefaultScope' } });
    assert.equal(actual, true);
  });

  it('fails validation if search scope passed is a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 2 } });
    assert.notEqual(actual, true);
  });

  it('sets search scope to default scope', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        SearchScope: 0
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'defaultscope' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets search scope to tenant', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        SearchScope: 1
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'tenant' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets search scope to hub', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        SearchScope: 2
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'hub' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets search scope to site', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        SearchScope: 3
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'site' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets search scope even if parameter is not all lower case', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (JSON.stringify(opts.body) === JSON.stringify({
        SearchScope: 3
      })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', searchScope: 'Site' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});