import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './tenant-site-membership-list.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.TENANT_SITE_MEMBERSHIP_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const membershipList = [
    {
      email: 'user1Email@email.com',
      loginName: 'i:0#.f|membership|user1loginName@email.com',
      name: 'user1DisplayName',
      userPrincipalName: 'user1loginName'
    },
    {
      email: 'user2Email@email.com',
      loginName: 'i:0#.f|membership|user2loginName@email.com',
      name: 'user2DisplayName',
      userPrincipalName: 'user2loginName'
    }
  ];
  const adminUrl = 'https://contoso-admin.sharepoint.com';
  const siteUrl = 'https://contoso.sharepoint.com/sites/site';
  const siteId = '00000000-0000-0000-0000-000000000010';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
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
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SITE_MEMBERSHIP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the siteUrl option is not a valid SharePoint URL', () => {
    const actual = command.validate({ options: { siteUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the role option is not a valid role', () => {
    const actual = command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', role: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('handles client.svc promise error', async () => {
    // get tenant app catalog
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        throw 'An error has occurred';
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('lists all site membership groups', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId=%27${siteId}%27&userGroupIds=[0,1,2]`) {
        return { value: [{ userGroup: membershipList }, { userGroup: membershipList }, { userGroup: membershipList }] };
      };
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl } });
    assert(loggerLogSpy.calledWith({
      AssociatedOwnerGroup: membershipList,
      AssociatedMemberGroup: membershipList,
      AssociatedVisitorGroup: membershipList
    }));
  });

  it('lists all site membership groups - just Owners group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId=%27${siteId}%27&userGroupIds=[0]`) {
        return { value: [{ userGroup: membershipList }] };
      };
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, role: "Owner" } });
    assert(loggerLogSpy.calledWith({
      AssociatedOwnerGroup: membershipList
    }));
  });

  it('lists all site membership groups - just Members group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId=%27${siteId}%27&userGroupIds=[1]`) {
        return { value: [{ userGroup: membershipList }] };
      };
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, role: "Member" } });
    assert(loggerLogSpy.calledWith({
      AssociatedMemberGroup: membershipList
    }));
  });

  it('lists all site membership groups - just Visitors group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId=%27${siteId}%27&userGroupIds=[2]`) {
        return { value: [{ userGroup: membershipList }] };
      };
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${siteId})/` }
            ]
          }
        ]
        );
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, role: "Visitor" } });
    assert(loggerLogSpy.calledWith({
      AssociatedVisitorGroup: membershipList
    }));
  });

  it('correctly handles error when site is not found for specified site URL', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": []
          }
        ]
        );
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl } }), new CommandError(`Failed to obtain site Id from the provided site URL.`));
  });

  it('correctly handles incorrect site Id guid', async () => {
    const incorrectSiteId = 'foo';

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }, 38, {
            "IsNull": false
          }, 40, {
            "IsNull": false
          }, 42, {
            "IsNull": false
          }, 44, {
            "IsNull": false
          }, 46, {
            "IsNull": false
          }, 48, {
            "_Child_Items_": [
              { SiteId: `/Guid(${incorrectSiteId})/` }
            ]
          }
        ]
        );
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl } }), new CommandError('Failed to obtain site Id from the provided site URL.'));
  });

  it('correctly handles error when site id is not found for specified site URL', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": {
              "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "d2d0389e-a040-4000-b24b-d16b0546a03c", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
            }, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
          }
        ]
        );
      }
      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl } }), new CommandError('Unknown Error'));
  });
});