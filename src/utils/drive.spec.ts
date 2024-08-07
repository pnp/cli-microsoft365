import assert from 'assert';
import sinon from 'sinon';
import { Logger } from '../cli/Logger.js';
import { sinonUtil } from "./sinonUtil.js";
import request from "../request.js";
import { Drive } from '@microsoft/microsoft-graph-types';
import { drive } from './drive.js';

describe('utils/drive', () => {
  let logger: Logger;
  let log: string[];
  const siteId = '0f9b8f4f-0e8e-4630-bb0a-501442db9b64';
  const driveId = '013TMHP6UOOSLON57HT5GLKEU7R5UGWZVK';
  const itemId = 'b!T4-bD44OMEa7ClAUQtubZID9tc40pGJKpguycvELod_Gx-lo4ZQiRJ7vylonTufG';
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const folderUrl: URL = new URL('https://contoso.sharepoint.com/sites/project-x/shared%20documents/');

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
      request.get
    ]);
  });

  const getDriveResponse: any = {
    value: [
      {
        "id": driveId,
        "webUrl": `${webUrl}/Shared%20Documents`
      }
    ]
  };

  const driveDetails: Drive = {
    id: driveId,
    webUrl: `${webUrl}/Shared%20Documents`
  };

  it('correctly gets drive by URL using getDriveByUrl', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`) {
        return getDriveResponse;
      }

      return 'Invalid Request';
    });

    const actual = await drive.getDriveByUrl(siteId, folderUrl, logger, true);
    assert.deepStrictEqual(actual, driveDetails);
  });

  it('correctly gets drive item ID using getDriveItemId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      const relativeItemUrl = folderUrl.href.replace(new RegExp(`${driveDetails.webUrl}`, 'i'), '').replace(/\/+$/, '');
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/root${relativeItemUrl ? `:${relativeItemUrl}` : ''}?$select=id`) {
        return { id: itemId };
      }

      return 'Invalid Request';
    });

    const actual = await drive.getDriveItemId(driveDetails, folderUrl, logger, true);
    assert.strictEqual(actual, itemId);
  });

  it('correctly gets drive item ID for a specific item using getDriveItemId', async () => {
    const folderUrl: URL = new URL('https://contoso.sharepoint.com/sites/project-x/shared%20documents/folder1');
    sinon.stub(request, 'get').callsFake(async opts => {
      const relativeItemUrl = folderUrl.href.replace(new RegExp(`${driveDetails.webUrl}`, 'i'), '').replace(/\/+$/, '');
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/root${relativeItemUrl ? `:${relativeItemUrl}` : ''}?$select=id`) {
        return { id: itemId };
      }

      return 'Invalid Request';
    });

    const actual = await drive.getDriveItemId(driveDetails, folderUrl, logger, true);
    assert.strictEqual(actual, itemId);
  });

  it('throws error when drive not found by url', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(
      drive.getDriveByUrl(siteId, folderUrl, logger, true),
      Error(`Drive '${folderUrl}' not found`));
  });
});