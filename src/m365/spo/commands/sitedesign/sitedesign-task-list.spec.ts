import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./sitedesign-task-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.SITEDESIGN_TASK_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_TASK_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about site designs scheduled for execution on the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTasks`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "ID": "e40b1c66-0292-4697-b686-f2b05446a588", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            },
            {
              "ID": "e40b1c66-0292-4697-b686-f2b05446a589", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76f", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ID": "e40b1c66-0292-4697-b686-f2b05446a588",
            "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
            "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com"
          },
          {
            "ID": "e40b1c66-0292-4697-b686-f2b05446a589",
            "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76f",
            "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about site designs scheduled for execution on the specified site (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTasks`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "ID": "e40b1c66-0292-4697-b686-f2b05446a588", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            },
            {
              "ID": "e40b1c66-0292-4697-b686-f2b05446a589", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76f", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ID": "e40b1c66-0292-4697-b686-f2b05446a588",
            "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
            "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com"
          },
          {
            "ID": "e40b1c66-0292-4697-b686-f2b05446a589",
            "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76f",
            "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs all information in JSON output mode', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTasks`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "ID": "e40b1c66-0292-4697-b686-f2b05446a588", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            },
            {
              "ID": "e40b1c66-0292-4697-b686-f2b05446a589", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76f", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ID": "e40b1c66-0292-4697-b686-f2b05446a588", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
          },
          {
            "ID": "e40b1c66-0292-4697-b686-f2b05446a589", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76f", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
          }
        ]));
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl is valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });
});