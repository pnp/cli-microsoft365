import commands from '../../commands';
import Command, { CommandValidate, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./userprofile-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.USERPROFILE_GET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  const spoUrl = 'https://contoso.sharepoint.com';
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = spoUrl;
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
      request.get
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
    assert.equal(command.name.startsWith(commands.USERPROFILE_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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
  
  it('retrieves user profile properties by user email with output option json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor') > -1) {
        return Promise.resolve(
          {
            "value": [{
              "AccountName": "i:0#.f|membership|john.doe@contoso.onmicrosoft.com",
              "DisplayName":"john doe",
              "Email": "john.doe@contoso.onmicrosoft.com",
              "ExtendedReports": "['i:0#.f|membership|john.doe@contoso.onmicrosoft.com']",
              "IsFollowed": false,
              "PersonalSiteHostUrl": "https://contoso-my.sharepoint.com:443/",
              "PersonalUrl": "https://contoso-my.sharepoint.com/personal/john.doe_contoso_onmicrosoft_com/",
              "UserProfileProperties":[
                {
                  "Key":"UserProfile_GUID",
                  "Value":"f3f102bb-7ac7-408e-9184-384062abd0d5",
                },
                {
                  "Key":"SID",
                  "Value":"i:0h.f|membership|10032000840f3681@live.com",
                }
               ]
            }]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'json',
        debug: false,
        userName: 'john.doe@contoso.onmicrosoft.com'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          value: [{
            AccountName: "i:0#.f|membership|john.doe@contoso.onmicrosoft.com",
            DisplayName:"john doe",
            Email: "john.doe@contoso.onmicrosoft.com",
            ExtendedReports: "['i:0#.f|membership|john.doe@contoso.onmicrosoft.com']",
            IsFollowed: false,
            PersonalSiteHostUrl: "https://contoso-my.sharepoint.com:443/",
            PersonalUrl: "https://contoso-my.sharepoint.com/personal/john.doe_contoso_onmicrosoft_com/",
            UserProfileProperties:[
              {
                "Key":"UserProfile_GUID",
                "Value":"f3f102bb-7ac7-408e-9184-384062abd0d5",
              },
              {
                "Key":"SID",
                "Value":"i:0h.f|membership|10032000840f3681@live.com",
              }]
          }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves user profile properties by user email with output option json (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor') > -1) {
        return Promise.resolve(
          {
            "odata.null": true
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'json',
        debug: true,
        userName: 'john.doe@contoso.onmicrosoft.com'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the user name option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
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
    assert(find.calledWith(commands.USERPROFILE_GET));
  });

  it('fails validation if the userName option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('passes validation when the input is correct', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
      }
    });
    assert.equal(actual, true);
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