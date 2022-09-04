import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./site-groupify');

describe(commands.SITE_GROUPIFY, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.SITE_GROUPIFY), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('connects site to an Microsoft 365 Group', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
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

    await command.action(logger, { options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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

  it('connects site to an Microsoft 365 Group (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
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

    await command.action(logger, { options: { debug: true, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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

  it('connects site to a public Microsoft 365 Group', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
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

    await command.action(logger, { options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', isPublic: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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

  it('setts Microsoft 365 Group description', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
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

    await command.action(logger, { options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', description: 'Team A space' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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

  it('sets Microsoft 365 Group classification', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
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

    await command.action(logger, { options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', classification: 'HBI' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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

  it('keeps the old home page', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
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

    await command.action(logger, { options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', keepOldHomepage: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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

  it('handles error when a group with the specified alias already exists', async () => {
    sinon.stub(request, 'post').callsFake(() => {
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

    await command.action(logger, { options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The group alias already exists.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the specified site already is connected to a group', async () => {
    sinon.stub(request, 'post').callsFake(() => {
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

    await command.action(logger, { options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('This site already has an O365 Group attached.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when creating site script', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    await command.action(logger, { options: { debug: false, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } } as any, (err?: any) => {
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
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if siteUrl is not an absolute URL', async () => {
    const actual = await command.validate({ options: { siteUrl: '/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if siteUrl is not a SharePoint URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'http://contoso/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required options are specified', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});