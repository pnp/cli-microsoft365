
import assert from 'assert';
import sinon from 'sinon';
import { spp } from './spp.js';
import { sinonUtil } from './sinonUtil.js';
import request from '../request.js';

describe('utils/spp', () => {
  const siteUrl = 'https://contoso.sharepoint.com';
  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('calls api correctly and throw an error when site is not a content center using assertSiteIsContentCenter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'SITEPAGEPUBLISHING#0'
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(spp.assertSiteIsContentCenter(siteUrl), Error('https://contoso.sharepoint.com is not a content site.'));
  });

  it('calls api correctly and does not throw an error when site is a content center using assertSiteIsContentCenter', async () => {
    const stubGet = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'CONTENTCTR#0'
        };
      }

      throw 'Invalid request';
    });

    await spp.assertSiteIsContentCenter(siteUrl);
    assert(stubGet.calledOnce);
  });
});