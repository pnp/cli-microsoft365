import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./site-inplacerecordsmanagement-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.SITE_INPLACERECORDSMANAGEMENT_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  
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
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post
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
    assert.equal(command.name.startsWith(commands.SITE_INPLACERECORDSMANAGEMENT_SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('correctly handles error when in-place records management already activated', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('_api/site/features/add') > -1) {
        return Promise.reject({
          error: {
            "odata.error": {
              "code": "-1, System.Data.DuplicateNameException",
              "message": {
                "lang": "en-US",
                "value": "Feature 'InPlaceRecords' (ID: da2e115b-07e4-49d9-bb2c-35e93bb9fca9) is already activated at scope 'https://contoso.sharepoint.com/sites/team-a'."
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: 'true' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("Feature 'InPlaceRecords' (ID: da2e115b-07e4-49d9-bb2c-35e93bb9fca9) is already activated at scope 'https://contoso.sharepoint.com/sites/team-a'.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when in-place records management already deactivated', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('_api/site/features/remove') > -1) {
        return Promise.reject({
          error: {
            "odata.error": {
              "code": "-1, System.InvalidOperationException",
              "message": {
                "lang": "en-US",
                "value": "Feature 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9' is not activated at this scope."
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: 'false' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("Feature 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9' is not activated at this scope.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should deactivate in-place records management', (done) => {
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('_api/site/features/remove') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: 'false' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/team-a/_api/site/features/remove');
        assert.equal(requestStub.lastCall.args[0].body.featureId, 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9');
        assert.equal(requestStub.lastCall.args[0].body.force, true);

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should activate in-place records management (verbose)', (done) => {
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('_api/site/features/add') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: 'true' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/team-a/_api/site/features/add');
        assert.equal(requestStub.lastCall.args[0].body.featureId, 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9');
        assert.equal(requestStub.lastCall.args[0].body.force, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should activate in-place records management', (done) => {
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('_api/site/features/add') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: 'true' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/team-a/_api/site/features/add');
        assert.equal(requestStub.lastCall.args[0].body.featureId, 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9');
        assert.equal(requestStub.lastCall.args[0].body.force, true);
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

  it('supports specifying siteUrl', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--siteUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying enabled', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--enabled') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if siteUrl not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { enabled: 'true' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if enabled option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if enabled option not "true" or "false"', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if siteUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: 'abc', enabled: 'true' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when the siteUrl is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: 'true' } });
    assert.equal(actual, true);
  });

  it('passes validation when the siteUrl is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: 'false' } });
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
    assert(find.calledWith(commands.SITE_INPLACERECORDSMANAGEMENT_SET));
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
});