import assert from 'assert';
import sinon from 'sinon';
import auth from '../Auth.js';
import { Logger } from '../cli/Logger.js';
import request from '../request.js';
import { sinonUtil } from '../utils/sinonUtil.js';
import { Drive } from '@microsoft/microsoft-graph-types';
import { drive as driveItem } from '../utils/drive.js';

describe('utils/drive', () => {
  let logger: Logger;
  let log: string[];

  const webUrl = 'https://contoso.sharepoint.com/sites/sales';
  const siteId = '0f9b8f4f-0e8e-4630-bb0a-501442db9b64';
  const itemId = 'b!T4-bD44OMEa7ClAUQtubZID9tc40pGJKpguycvELod_Gx-lo4ZQiRJ7vylonTufG';
  const folderUrl: URL = new URL('https://contoso.sharepoint.com/sites/sales/shared%20documents/');
  const driveId = '013TMHP6UOOSLON57HT5GLKEU7R5UGWZVK';
  const drive: Drive = {
    id: driveId,
    webUrl: `${webUrl}/Shared%20Documents`
  };

  before(() => {
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      auth.storeConnectionInfo,
      driveItem.getDrive,
      driveItem.getDriveItemId,
      global.setTimeout
    ]);
    auth.connection.spoUrl = undefined;
    auth.connection.spoTenantId = undefined;
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('correctly gets drive using getDrive (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`) {
        return {
          value: [
            {
              "id": "013TMHP6UOOSLON57HT5GLKEU7R5UGWZVK",
              "webUrl": `${webUrl}/Shared%20Documents`
            }
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await driveItem.getDrive(siteId, folderUrl, logger, true);
    assert.deepStrictEqual(actual, drive);
  });

  it('correctly gets drive item ID using getDriveItemId (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      const relativeItemUrl = folderUrl.href.replace(new RegExp(`${drive.webUrl}`, 'i'), '').replace(/\/+$/, '');
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/root${relativeItemUrl ? `:${relativeItemUrl}` : ''}?$select=id`) {
        return { id: itemId };
      }

      return 'Invalid Request';
    });

    const actual = await driveItem.getDriveItemId(drive, folderUrl, logger, true);
    assert.strictEqual(actual, itemId);
  });

  it('correctly gets drive item ID using getDriveItemId when relativeItemUrl matches drive webUrl', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/root?$select=id`) {
        return { id: itemId };
      }

      return 'Invalid Request';
    });

    const actual = await driveItem.getDriveItemId(drive, folderUrl, logger, true);
    assert.strictEqual(actual, itemId);
  });

  it('handles when drive not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    try {
      await driveItem.getDrive(siteId, folderUrl, logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, `Drive '${folderUrl.href}' not found`);
    }
  });

  it('throws error when drive not found by url', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`) {
        throw `Drive '${folderUrl.href}' not found`;
      }

      throw 'Invalid request';
    });

    try {
      await driveItem.getDrive(siteId, folderUrl, logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, `Drive '${folderUrl.href}' not found`);
    }
  });
});