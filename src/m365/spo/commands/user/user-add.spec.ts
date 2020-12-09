import commands from '../../commands';
import Command, { CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./user-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.USER_ADD, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      request.get
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
    assert.equal(command.name.startsWith(commands.USER_ADD), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  
  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {
      email:"john.doe@contoso.onmicrosoft.com",
      group:"Team Site Members"
    } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the email option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { 
      webUrl: 'https://contoso.sharepoint.com/sites/mysite', 
      group: "Team Site Members" 
    } });
    assert.notEqual(actual, true);
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

  it('correctly retrieves groupId and add user to specified SharePoint group', (done) => {
    let addMemberRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/mysite/_api/web/sitegroups/GetByName('Team%20Site%20Members')`) {
        return Promise.resolve({
          "Id": "7"
        });
      }

      return Promise.reject('Invalid request');
    });
    
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/mysite/_api/web/sitegroups/GetById('7')/users` &&
        JSON.stringify(opts.body) === `{"__metadata":{"type":"SP.User"},"LoginName":"i:0#.f|membership|john.doe@contoso.onmicrosoft.com"}`) {
        addMemberRequestIssued = true;
      }
      return Promise.resolve();
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug:false,webUrl: 'https://contoso.sharepoint.com/sites/mysite', 
        email: "john.doe@contoso.onmicrosoft.com", 
        group: "Team Site Members" } 
      }, () => {
      try {
        assert(addMemberRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the group option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { 
      webUrl: 'https://contoso.sharepoint.com/sites/mysite', 
      email: "john.doe@contoso.onmicrosoft.com" 
    } });
    assert.notEqual(actual, true);
  });

  it('fails validation if url is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'abc' } });
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
    assert(find.calledWith(commands.USER_ADD));
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