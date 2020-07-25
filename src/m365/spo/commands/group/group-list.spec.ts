import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./group-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.GROUP_LIST, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    assert.strictEqual(command.name.startsWith(commands.GROUP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('Lists all the groups within specific web with output option json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups') > -1) {
        return Promise.resolve(
          {
            "value": [{
              "Id": 15,
              "Title": "Contoso Members",
              "LoginName": "Contoso Members",
              "Description": "SharePoint Contoso",
              "IsHiddenInUI": false,
              "PrincipalType": 8
            }]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          value: [{
            Id: 15,
            Title: "Contoso Members",
            LoginName: "Contoso Members",
            Description: "SharePoint Contoso",
            IsHiddenInUI: false,
            PrincipalType: 8
          }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all groups with output option text', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups') > -1) {
        return Promise.resolve(
          {
            "value": [
              {
                "Id": 15,
                "Title": "Contoso Members",
                "LoginName": "Contoso Members",
                "Description": "SharePoint Contoso",
                "IsHiddenInUI": false,
                "PrincipalType": 8
              }
            ]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          [{
            Id: 15,
            Title: "Contoso Members",
            LoginName: "Contoso Members",
            IsHiddenInUI: false,
            PrincipalType: 8
          }]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('command correctly handles group list reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports specifying URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });


  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation url is valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });
}); 