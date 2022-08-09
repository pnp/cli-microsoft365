import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./group-user-list');

describe(commands.GROUP_USER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const JSONSPGroupMembersList =
  {
    "value": [
      {
        "Id": 6,
        "IsHiddenInUI": false,
        "LoginName": "i:0#.f|membership|Alex.Wilber@contoso.com.com",
        "Title": "Alex Wilber",
        "PrincipalType": 1,
        "Email": "Alex.Wilber@contoso.com",
        "Expiration": "",
        "IsEmailAuthenticationGuestUser": false,
        "IsShareByEmailGuestUser": false,
        "IsSiteAdmin": true,
        "UserId": {
          "NameId": "10032000afc2e592",
          "NameIdIssuer": "urn:federation:microsoftonline"
        },
        "UserPrincipalName": "Alex.Wilber@contoso.com"
      },
      {
        "Id": 18,
        "IsHiddenInUI": false,
        "LoginName": "i:0#.f|membership|AdeleV@contoso.com",
        "Title": "Adele Vance",
        "PrincipalType": 1,
        "Email": "AdeleV@contoso.com",
        "Expiration": "",
        "IsEmailAuthenticationGuestUser": false,
        "IsShareByEmailGuestUser": false,
        "IsSiteAdmin": false,
        "UserId": {
          "NameId": "10032000b07c0d71",
          "NameIdIssuer": "urn:federation:microsoftonline"
        },
        "UserPrincipalName": "AdeleV@contoso.com"
      }
    ]
  };

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      Cli.executeCommandWithOutput
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
    assert.strictEqual(command.name.startsWith(commands.GROUP_USER_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Title', 'UserPrincipalName', 'Id', 'Email']);
  });

  it('Getting the members of a SharePoint Group using groupId', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return Promise.resolve(JSONSPGroupMembersList);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    command.action(logger, {
      options: {
        debug: false,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Getting the members of a SharePoint Group using groupId (DEBUG)', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return Promise.resolve(JSONSPGroupMembersList);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Getting the members of a SharePoint Group using groupName (DEBUG)', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetByName') > -1) {
        return Promise.resolve(JSONSPGroupMembersList);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Contoso Site Owners"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Correctly Handles Error when listing members of the group', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return Promise.reject('Invalid request');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError('Invalid request')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if webURL is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "InvalidWEBURL", groupId: 3 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupid and groupName is entered', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "4", groupName: "Contoso Site Owners" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither groupId nor groupName is entered', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupId is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "INVALIDGROUP" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all the required options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 3 } }, commandInfo);
    assert.strictEqual(actual, true);
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
});