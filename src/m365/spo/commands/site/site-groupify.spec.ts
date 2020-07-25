import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./site-groupify');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.SITE_GROUPIFY, () => {
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
    assert.strictEqual(command.name.startsWith(commands.SITE_GROUPIFY), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('connects site to an Microsoft 365 Group', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.body) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: false,
          optionalParams: {}
        })) {
        return Promise.resolve({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('connects site to an Microsoft 365 Group (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.body) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: false,
          optionalParams: {}
        })) {
        return Promise.resolve({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('connects site to a public Microsoft 365 Group', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.body) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: true,
          optionalParams: {}
        })) {
        return Promise.resolve({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', isPublic: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('setts Microsoft 365 Group description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.body) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: false,
          optionalParams: {
            Description: 'Team A space'
          }
        })) {
        return Promise.resolve({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', description: 'Team A space' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets Microsoft 365 Group classification', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.body) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: false,
          optionalParams: {
            Classification: 'HBI'
          }
        })) {
        return Promise.resolve({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', classification: 'HBI' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('keeps the old home page', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.body) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: false,
          optionalParams: {
            CreationOptions: ["SharePointKeepOldHomepage"]
          }
        })) {
        return Promise.resolve({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', keepOldHomepage: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when a group with the specified alias already exists', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-2147024713, Microsoft.SharePoint.SPException",
            "message": {
              "lang": "en-US",
              "value": "The group alias already exists."
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The group alias already exists.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the specified site already is connected to a group', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-2147024809, System.ArgumentException",
            "message": {
              "lang": "en-US",
              "value": "This site already has an O365 Group attached."
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('This site already has an O365 Group attached.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when creating site script', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    cmdInstance.action({ options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, (err?: any) => {
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
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if siteUrl is not an absolute URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: '/sites/team-a', alias: 'team-a', displayName: 'Team A' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if siteUrl is not a SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: 'http://contoso/sites/team-a', alias: 'team-a', displayName: 'Team A' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required options are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } });
    assert.strictEqual(actual, true);
  });
});