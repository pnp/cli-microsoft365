import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./userprofile-get');

describe(commands.USERPROFILE_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = true;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.USERPROFILE_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets userprofile information about the specified user', (done) => {
    const profile = {
      "AccountName": "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com",
      "DirectReports": [],
      "DisplayName": "Dipen Shah",
      "Email": "dips1802@dev1802.onmicrosoft.com",
      "ExtendedManagers": [],
      "ExtendedReports": [
        "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com"
      ],
      "IsFollowed": false,
      "LatestPost": null,
      "Peers": [],
      "PersonalSiteHostUrl": "https://dev1802-my.sharepoint.com:443/",
      "PersonalUrl": "https://dev1802-my.sharepoint.com/personal/dips1802_dev1802_onmicrosoft_com/",
      "PictureUrl": null,
      "Title": null,
      "UserProfileProperties": [
        {
          "Key": "UserProfile_GUID",
          "Value": "f3f102bb-7ac7-408e-9184-384062abd0d5"
        },
        {
          "Key": "SID",
          "Value": "i:0h.f|membership|10032000840f3681@live.com"
        }]
    };
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor') > -1) {
        // we need to clone the object because it's changed in the command
        // and would skew the comparison in the test outcome
        return Promise.resolve(JSON.parse(JSON.stringify(profile)));
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, {
      options: {
        output: 'text',
        userName: 'john.doe@contoso.onmicrosoft.com'
      }
    } as any, () => {
      try {
        const loggedProfile = JSON.parse(JSON.stringify(profile));
        loggedProfile.UserProfileProperties = JSON.stringify(loggedProfile.UserProfileProperties);
        assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(loggedProfile));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets userprofile information about the specified user output json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor') > -1) {
        return Promise.resolve({
          "AccountName": "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com",
          "DirectReports": [],
          "DisplayName": "Dipen Shah",
          "Email": "dips1802@dev1802.onmicrosoft.com",
          "ExtendedManagers": [],
          "ExtendedReports": [
            "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com"
          ],
          "IsFollowed": false,
          "LatestPost": null,
          "Peers": [],
          "PersonalSiteHostUrl": "https://dev1802-my.sharepoint.com:443/",
          "PersonalUrl": "https://dev1802-my.sharepoint.com/personal/dips1802_dev1802_onmicrosoft_com/",
          "PictureUrl": null,
          "Title": null
        });
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        userName: 'john.doe@contoso.onmicrosoft.com'
      }
    } as any, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "AccountName": "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com",
          "DirectReports": [],
          "DisplayName": "Dipen Shah",
          "Email": "dips1802@dev1802.onmicrosoft.com",
          "ExtendedManagers": [],
          "ExtendedReports": [
            "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com"
          ],
          "IsFollowed": false,
          "LatestPost": null,
          "Peers": [],
          "PersonalSiteHostUrl": "https://dev1802-my.sharepoint.com:443/",
          "PersonalUrl": "https://dev1802-my.sharepoint.com/personal/dips1802_dev1802_onmicrosoft_com/",
          "PictureUrl": null,
          "Title": null
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('supports specifying userName', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--userName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the user principal name is not a valid', async () => {
    const actual = await command.validate({ options: { userName: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the user principal name is a valid', async () => {
    const actual = await command.validate({ options: { userName: 'john.doe@mytenant.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});