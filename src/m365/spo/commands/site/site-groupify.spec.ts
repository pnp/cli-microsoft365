import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './site-groupify.js';

describe(commands.SITE_GROUPIFY, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
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
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_GROUPIFY);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('connects site to an Microsoft 365 Group', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: false,
          optionalParams: {}
        })) {
        return {
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } });
    assert(loggerLogSpy.calledWith({
      "DocumentsUrl": null,
      "ErrorMessage": null,
      "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
      "SiteStatus": 2,
      "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
    }));
  });

  it('connects site to an Microsoft 365 Group (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: false,
          optionalParams: {}
        })) {
        return {
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } });
    assert(loggerLogSpy.calledWith({
      "DocumentsUrl": null,
      "ErrorMessage": null,
      "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
      "SiteStatus": 2,
      "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
    }));
  });

  it('connects site to a public Microsoft 365 Group', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: true,
          optionalParams: {}
        })) {
        return {
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', isPublic: true } });
    assert(loggerLogSpy.calledWith({
      "DocumentsUrl": null,
      "ErrorMessage": null,
      "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
      "SiteStatus": 2,
      "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
    }));
  });

  it('setts Microsoft 365 Group description', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: false,
          optionalParams: {
            Description: 'Team A space'
          }
        })) {
        return {
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', description: 'Team A space' } });
    assert(loggerLogSpy.calledWith({
      "DocumentsUrl": null,
      "ErrorMessage": null,
      "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
      "SiteStatus": 2,
      "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
    }));
  });

  it('sets Microsoft 365 Group classification', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: false,
          optionalParams: {
            Classification: 'HBI'
          }
        })) {
        return {
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', classification: 'HBI' } });
    assert(loggerLogSpy.calledWith({
      "DocumentsUrl": null,
      "ErrorMessage": null,
      "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
      "SiteStatus": 2,
      "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
    }));
  });

  it('keeps the old home page', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/GroupSiteManager/CreateGroupForSite' &&
        JSON.stringify(opts.data) === JSON.stringify({
          displayName: 'Team A',
          alias: 'team-a',
          isPublic: false,
          optionalParams: {
            CreationOptions: ["SharePointKeepOldHomepage"]
          }
        })) {
        return {
          "DocumentsUrl": null,
          "ErrorMessage": null,
          "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
          "SiteStatus": 2,
          "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A', keepOldHomepage: true } });
    assert(loggerLogSpy.calledWith({
      "DocumentsUrl": null,
      "ErrorMessage": null,
      "GroupId": "114e2be8-7e34-4ed1-b528-7f3762d36a6c",
      "SiteStatus": 2,
      "SiteUrl": "https://contoso.sharepoint.com/sites/team-a"
    }));
  });

  it('handles error when a group with the specified alias already exists', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      throw {
        error: {
          "odata.error": {
            "code": "-2147024713, Microsoft.SharePoint.SPException",
            "message": {
              "lang": "en-US",
              "value": "The group alias already exists."
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } } as any), new CommandError('The group alias already exists.'));
  });

  it('handles error when the specified site already is connected to a group', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      throw {
        error: {
          "odata.error": {
            "code": "-2147024809, System.ArgumentException",
            "message": {
              "lang": "en-US",
              "value": "This site already has an M365 Group attached."
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } } as any), new CommandError('This site already has an M365 Group attached.'));
  });

  it('correctly handles OData error when creating site script', async () => {
    sinon.stub(request, 'post').rejects({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if url is not an absolute URL', async () => {
    const actual = await command.validate({ options: { url: '/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if url is not a SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'http://contoso/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required options are specified', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a', alias: 'team-a', displayName: 'Team A' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
