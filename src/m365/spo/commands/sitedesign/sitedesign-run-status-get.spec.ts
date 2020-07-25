import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./sitedesign-run-status-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.SITEDESIGN_RUN_STATUS_GET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_RUN_STATUS_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about site designs applied to the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRunStatus`) > -1) {
        return Promise.resolve({
          "value": [
            { "ActionIndex": 0, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Add to hub site", "LastModified": "1548960114000", "OrdinalIndex": 0, "OutcomeCode": 1, "OutcomeText": "One or more of the properties on this action has an invalid type.", "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" },
            { "ActionIndex": 1, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Associate SPFX extension Collab Footer", "LastModified": "1548960114000", "OrdinalIndex": 1, "OutcomeCode": 0, "OutcomeText": null, "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ActionTitle": "Add to hub site",
            "SiteScriptTitle": "Contoso Team Site",
            "OutcomeText": "One or more of the properties on this action has an invalid type."
          },
          {
            "ActionTitle": "Associate SPFX extension Collab Footer",
            "SiteScriptTitle": "Contoso Team Site",
            "OutcomeText": null
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
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRunStatus`) > -1) {
        return Promise.resolve({
          "value": [
            { "ActionIndex": 0, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Add to hub site", "LastModified": "1548960114000", "OrdinalIndex": 0, "OutcomeCode": 1, "OutcomeText": "One or more of the properties on this action has an invalid type.", "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" },
            { "ActionIndex": 1, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Associate SPFX extension Collab Footer", "LastModified": "1548960114000", "OrdinalIndex": 1, "OutcomeCode": 0, "OutcomeText": null, "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8', output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          { "ActionIndex": 0, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Add to hub site", "LastModified": "1548960114000", "OrdinalIndex": 0, "OutcomeCode": 1, "OutcomeText": "One or more of the properties on this action has an invalid type.", "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" },
          { "ActionIndex": 1, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Associate SPFX extension Collab Footer", "LastModified": "1548960114000", "OrdinalIndex": 1, "OutcomeCode": 0, "OutcomeText": null, "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified runId doesn\'t point to a valid run', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'Value does not fall within the expected range' } } } });
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Value does not fall within the expected range')));
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
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'invalid', runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if runId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', runId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl and runId are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', runId: '6ec3ca5b-d04b-4381-b169-61378556d76e' } });
    assert.strictEqual(actual, true);
  });
});