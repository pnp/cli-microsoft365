import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./sitedesign-task-get');

describe(commands.SITEDESIGN_TASK_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_TASK_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about site design scheduled for execution', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTask`) > -1) {
        return Promise.resolve({
          "ID": "e40b1c66-0292-4697-b686-f2b05446a588", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, taskId: 'e40b1c66-0292-4697-b686-f2b05446a588' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "ID": "e40b1c66-0292-4697-b686-f2b05446a588", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about site design scheduled for execution (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTask`) > -1) {
        return Promise.resolve({
          "ID": "e40b1c66-0292-4697-b686-f2b05446a588", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, taskId: 'e40b1c66-0292-4697-b686-f2b05446a588' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "ID": "e40b1c66-0292-4697-b686-f2b05446a588", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles site design task not found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTask`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, taskId: 'e40b1c66-0292-4697-b686-f2b05446a588' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when retrieving information about site designs', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('fails validation if taskId is not a valid GUID', () => {
    const actual = command.validate({ options: { taskId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if taskId is valid', () => {
    const actual = command.validate({ options: { taskId: '6ec3ca5b-d04b-4381-b169-61378556d76e' } });
    assert.strictEqual(actual, true);
  });
});