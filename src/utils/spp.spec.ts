
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

  it('calls api correctly and returns false if site is not a content center', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if ((opts.url as string).indexOf(`/_api/web?$select=WebTemplateConfiguration`) > -1) {
        return {
          WebTemplateConfiguration: 'SITEPAGEPUBLISHING#0'
        };
      }

      throw 'Invalid request';
    });

    const actual = await spp.isContentCenter(siteUrl);
    assert.strictEqual(actual, false);
  });

  it('calls api correctly and returns true if site is a content center', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if ((opts.url as string).indexOf(`/_api/web?$select=WebTemplateConfiguration`) > -1) {
        return {
          WebTemplateConfiguration: 'CONTENTCTR#0'
        };
      }

      throw 'Invalid request';
    });

    const actual = await spp.isContentCenter(siteUrl);
    assert.strictEqual(actual, true);
  });
});