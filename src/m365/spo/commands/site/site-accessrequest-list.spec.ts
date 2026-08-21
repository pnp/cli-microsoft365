import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { z } from 'zod';
import commands from '../../commands.js';
import command from './site-accessrequest-list.js';

describe(commands.SITE_ACCESSREQUEST_LIST, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let loggerLogSpy: sinon.SinonSpy;

  const siteUrl = 'https://contoso.sharepoint.com/sites/Marketing';

  const restItemsResponse = [
    {
      Id: 1,
      Title: null,
      RequestDate: '2024-09-03T22:07:04Z',
      Status: 0,
      PermissionLevelRequested: 5,
      PermissionType: null,
      IsInvitation: false,
      Conversation: null,
      RequestedObjectUrl: 'https://contoso.sharepoint.com/sites/Marketing',
      RequestedObjectTitle: null,
      RequestedByDisplayName: null,
      RequestedForDisplayName: 'John Doe'
    },
    {
      Id: 2,
      Title: null,
      RequestDate: '2024-09-04T10:30:00Z',
      Status: 1,
      PermissionLevelRequested: 3,
      PermissionType: null,
      IsInvitation: false,
      Conversation: null,
      RequestedObjectUrl: null,
      RequestedObjectTitle: null,
      RequestedByDisplayName: null,
      RequestedForDisplayName: 'Jane Smith'
    },
    {
      Id: 3,
      Title: null,
      RequestDate: '2024-09-05T08:15:00Z',
      Status: 3,
      PermissionLevelRequested: 5,
      PermissionType: null,
      IsInvitation: false,
      Conversation: null,
      RequestedObjectUrl: null,
      RequestedObjectTitle: null,
      RequestedByDisplayName: null,
      RequestedForDisplayName: 'Bob Johnson'
    }
  ];

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    auth.connection.active = true;
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
      odata.getAllItems
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_ACCESSREQUEST_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'RequestDate', 'RequestedForDisplayName', 'PermissionLevelRequested', 'StatusLabel']);
  });

  it('fails validation if siteUrl is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if state is not a valid value', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: siteUrl, state: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with only siteUrl', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: siteUrl });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with siteUrl and state', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: siteUrl, state: 'pending' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with state set to approved', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: siteUrl, state: 'approved' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with state set to declined', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: siteUrl, state: 'declined' });
    assert.strictEqual(actual.success, true);
  });

  it('lists all access requests', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${siteUrl}/_api/web/AccessRequestsList/Items?$select=*,Status`) {
        return restItemsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, verbose: true } });
    assert(loggerLogSpy.calledWith([
      { Id: 1, Title: null, RequestDate: '2024-09-03T22:07:04Z', Status: 0, PermissionLevelRequested: 5, PermissionType: null, IsInvitation: false, Conversation: null, RequestedObjectUrl: 'https://contoso.sharepoint.com/sites/Marketing', RequestedObjectTitle: null, RequestedByDisplayName: null, RequestedForDisplayName: 'John Doe', StatusLabel: 'pending' },
      { Id: 2, Title: null, RequestDate: '2024-09-04T10:30:00Z', Status: 1, PermissionLevelRequested: 3, PermissionType: null, IsInvitation: false, Conversation: null, RequestedObjectUrl: null, RequestedObjectTitle: null, RequestedByDisplayName: null, RequestedForDisplayName: 'Jane Smith', StatusLabel: 'approved' },
      { Id: 3, Title: null, RequestDate: '2024-09-05T08:15:00Z', Status: 3, PermissionLevelRequested: 5, PermissionType: null, IsInvitation: false, Conversation: null, RequestedObjectUrl: null, RequestedObjectTitle: null, RequestedByDisplayName: null, RequestedForDisplayName: 'Bob Johnson', StatusLabel: 'declined' }
    ]));
  });

  it('lists only pending access requests', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${siteUrl}/_api/web/AccessRequestsList/Items?$select=*,Status&$filter=Status eq 0`) {
        return [restItemsResponse[0]];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, state: 'pending' } });
    assert(loggerLogSpy.calledWith([
      { Id: 1, Title: null, RequestDate: '2024-09-03T22:07:04Z', Status: 0, PermissionLevelRequested: 5, PermissionType: null, IsInvitation: false, Conversation: null, RequestedObjectUrl: 'https://contoso.sharepoint.com/sites/Marketing', RequestedObjectTitle: null, RequestedByDisplayName: null, RequestedForDisplayName: 'John Doe', StatusLabel: 'pending' }
    ]));
  });

  it('lists only approved access requests', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${siteUrl}/_api/web/AccessRequestsList/Items?$select=*,Status&$filter=Status eq 1`) {
        return [restItemsResponse[1]];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, state: 'approved' } });
    assert(loggerLogSpy.calledWith([
      { Id: 2, Title: null, RequestDate: '2024-09-04T10:30:00Z', Status: 1, PermissionLevelRequested: 3, PermissionType: null, IsInvitation: false, Conversation: null, RequestedObjectUrl: null, RequestedObjectTitle: null, RequestedByDisplayName: null, RequestedForDisplayName: 'Jane Smith', StatusLabel: 'approved' }
    ]));
  });

  it('lists only declined access requests', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${siteUrl}/_api/web/AccessRequestsList/Items?$select=*,Status&$filter=Status eq 3`) {
        return [restItemsResponse[2]];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl, state: 'declined' } });
    assert(loggerLogSpy.calledWith([
      { Id: 3, Title: null, RequestDate: '2024-09-05T08:15:00Z', Status: 3, PermissionLevelRequested: 5, PermissionType: null, IsInvitation: false, Conversation: null, RequestedObjectUrl: null, RequestedObjectTitle: null, RequestedByDisplayName: null, RequestedForDisplayName: 'Bob Johnson', StatusLabel: 'declined' }
    ]));
  });

  it('handles empty response', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${siteUrl}/_api/web/AccessRequestsList/Items?$select=*,Status`) {
        return [];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('returns empty array when the access requests list does not exist yet', async () => {
    sinon.stub(odata, 'getAllItems').rejects({ error: { error: { code: '-1, Microsoft.SharePoint.Client.ResourceNotFoundException', message: 'Cannot find resource for the request SP.RequestContext.current/web/AccessRequestsList/.' } } });

    await command.action(logger, { options: { siteUrl: siteUrl } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('correctly handles error when retrieving access requests', async () => {
    sinon.stub(odata, 'getAllItems').rejects({ error: { 'odata.error': { message: { value: 'An error has occurred.' } } } });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl } }),
      new CommandError('An error has occurred.'));
  });

  it('sets StatusLabel to unknown for unrecognized status values', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${siteUrl}/_api/web/AccessRequestsList/Items?$select=*,Status`) {
        return [{ ...restItemsResponse[0], Status: 99 }];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: siteUrl } });
    assert(loggerLogSpy.calledWith([
      { Id: 1, Title: null, RequestDate: '2024-09-03T22:07:04Z', Status: 99, PermissionLevelRequested: 5, PermissionType: null, IsInvitation: false, Conversation: null, RequestedObjectUrl: 'https://contoso.sharepoint.com/sites/Marketing', RequestedObjectTitle: null, RequestedByDisplayName: null, RequestedForDisplayName: 'John Doe', StatusLabel: 'unknown' }
    ]));
  });
});
