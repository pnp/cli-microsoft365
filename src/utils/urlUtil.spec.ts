import assert from 'assert';
import { urlUtil } from '../utils/urlUtil.js';

describe('urlUtil/urlUtil', () => {
  it('should get server relative path when https://contoso.sharepoint.com/sites/team1', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/sites/team1', '');
    assert.strictEqual(actual, '/sites/team1');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/team1/', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/sites/team1/', '');
    assert.strictEqual(actual, '/sites/team1');
  });

  it('should get server relative path when https://contoso.sharepoint.com/', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/', '');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path when domain only', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com', '');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path when /sites/team1 relative path passed as param', () => {
    const actual = urlUtil.getServerRelativePath('/sites/team1', '');
    assert.strictEqual(actual, '/sites/team1');
  });

  it('should get server relative path when /sites/team1/ relative path passed as param', () => {
    const actual = urlUtil.getServerRelativePath('/sites/team1/', '');
    assert.strictEqual(actual, '/sites/team1');
  });

  it('should get server relative path when / relative path passed as param', () => {
    const actual = urlUtil.getServerRelativePath('/', '');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path for https://contoso.sharepoint.com/sites/team1 and Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/sites/team1', 'Shared Documents');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/team1/ and /Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/sites/team1/', '/Shared Documents');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/ and /Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when domain only and Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when /sites/team1 and /Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('/sites/team1', '/Shared Documents');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when /sites/team1 and /Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('/sites/team1', '/Shared Documents/');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when /sites/team1/ and /Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('/sites/team1/', '/Shared Documents/');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when sites/team1/ and Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('sites/team1/', 'Shared Documents');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when / and Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('/', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when / and /Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('/', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when / and /Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('/', '/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when "" and ""', () => {
    const actual = urlUtil.getServerRelativePath('', '');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path when / and /', () => {
    const actual = urlUtil.getServerRelativePath('/', '/');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path when "" and /', () => {
    const actual = urlUtil.getServerRelativePath('', '/');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path when "" and Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('https://contoso.sharepoint.com/sites/site1/', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('/sites/site1/', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('/sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('/sites/site1/', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and sites/site1/Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('/sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and sites/site1/Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('/sites/site1/', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('/sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('/sites/site1/', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('/sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when uppercase in web url e.g. sites/Site1 and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('sites/Site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/Site1/Shared Documents');
  });

  it('should get server relative path when uppercase in folder url e.g. sites/site1 and /sites/Site1/Shared Documents', () => {
    const actual = urlUtil.getServerRelativePath('sites/site1', '/sites/Site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sub folder present url e.g. sites/site1 and /sites/Site1/Shared Documents/MyFolder', () => {
    const actual = urlUtil.getServerRelativePath('sites/site1', '/sites/Site1/Shared Documents/MyFolder');
    assert.strictEqual(actual, '/sites/site1/Shared Documents/MyFolder');
  });

  it('should get server relative path when https://CONTOSO.sharepoint.com/sites/team1', () => {
    const actual = urlUtil.getServerRelativePath('https://CONTOSO.sharepoint.com/sites/team1', '');
    assert.strictEqual(actual, '/sites/team1');
  });

  it('should get web relative path when / relative path passed as param', () => {
    const actual = urlUtil.getWebRelativePath('/', '/');
    assert.strictEqual(actual, '/');
  });

  it('should get web relative path for https://contoso.sharepoint.com/sites/team1 and Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('https://contoso.sharepoint.com/sites/team1', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/team1/ and /Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('https://contoso.sharepoint.com/sites/team1/', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/ and /Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('https://contoso.sharepoint.com/', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when domain only and Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('https://contoso.sharepoint.com', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/team1 and /Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('/sites/team1', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/team1 and /Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('/sites/team1', '/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/team1/ and /Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('/sites/team1/', '/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/team1/ and Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('sites/team1/', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/team1 and Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('sites/team1', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when / and Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('/', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when / and /Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('/', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when / and /Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('/', '/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when "" and ""', () => {
    const actual = urlUtil.getWebRelativePath('', '');
    assert.strictEqual(actual, '/');
  });

  it('should get web relative path when / and /', () => {
    const actual = urlUtil.getWebRelativePath('/', '/');
    assert.strictEqual(actual, '/');
  });

  it('should get web relative path when "" and /', () => {
    const actual = urlUtil.getWebRelativePath('', '/');
    assert.strictEqual(actual, '/');
  });

  it('should get web relative path when "" and Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('https://contoso.sharepoint.com/sites/site1/', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('/sites/site1/', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('/sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('/sites/site1/', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and sites/site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('/sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and sites/site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('/sites/site1/', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('/sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('/sites/site1/', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('/sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = urlUtil.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when uppercase in web url e.g. sites/Site1 and /sites/site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('sites/Site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when uppercase in folder url e.g. sites/site1 and /sites/Site1/Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('sites/site1', '/sites/Site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sub folder present url e.g. sites/site1 and /sites/Site1/Shared Documents/MyFolder', () => {
    const actual = urlUtil.getWebRelativePath('sites/site1', '/sites/Site1/Shared Documents/MyFolder');
    assert.strictEqual(actual, '/Shared Documents/MyFolder');
  });

  it('should get web relative path for https://CONTOSO.sharepoint.com/sites/team1 and Shared Documents', () => {
    const actual = urlUtil.getWebRelativePath('https://CONTOSO.sharepoint.com/sites/team1', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get absolute URL of a folder using webUrl and the folder server relative url', () => {
    const actual = urlUtil.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1', '/sites/team1/Shared Documents/MyFolder');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('should handle the server relative url starting by / or not while getting absolute URL of a folder', () => {
    const actual = urlUtil.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1', 'sites/team1/Shared Documents/MyFolder');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('should handle the presence of an ending / on the web url while getting absolute URL of a folder', () => {
    const actual = urlUtil.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1/', 'sites/team1/Shared Documents/MyFolder');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('should properly concatenate URL parts even with ending and starting / to each while getting absolute URL of a folder', () => {
    const actual = urlUtil.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1/', '/sites/team1/Shared Documents/MyFolder');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('combines URL successfully ending on a slash', () => {
    const actual = urlUtil.urlCombine('https://contoso.sharepoint.com/sites/team1/', '/sites/team1/Shared Documents/MyFolder');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/team1/sites/team1/Shared Documents/MyFolder');
  });

  it('combines URL successfully when second part ends with a slash', () => {
    const actual = urlUtil.urlCombine('https://contoso.sharepoint.com/sites/team1', '/sites/team1/Shared Documents/MyFolder/');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/team1/sites/team1/Shared Documents/MyFolder');
  });

  it('should return the correct target site absolute URL when the targetUrl is relative and starts with "/"', () => {
    const actual = urlUtil.getTargetSiteAbsoluteUrl('https://contoso.sharepoint.com', '/teams/sales/Shared Documents/temp/123/234');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com/teams/sales');
  });

  it('should return the correct target site absolute URL when the targetUrl is an absolute SharePoint URL', () => {
    const actual = urlUtil.getTargetSiteAbsoluteUrl('https://contoso.sharepoint.com', 'https://contoso-my.sharepoint.com/personal/john_contoso_onmicrosoft_com/Documents/123');
    assert.strictEqual(actual, 'https://contoso-my.sharepoint.com/personal/john_contoso_onmicrosoft_com');
  });

  it('should return the correct target site absolute URL when the webUrl is not the root site', () => {
    const actual = urlUtil.getTargetSiteAbsoluteUrl('https://contoso.sharepoint.com/teams/sales', '/Shared Documents/temp');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com');
  });

  it('should return the correct target site absolute URL when the targetUrl does not match SharePoint URL pattern', () => {
    const actual = urlUtil.getTargetSiteAbsoluteUrl('https://contoso.sharepoint.com', 'https://example.com/some/path');
    assert.strictEqual(actual, 'https://example.com');
  });

  it('correctly removes leading slashes from the URL', () => {
    const actual = urlUtil.removeLeadingSlashes('/Shared Documents/MyFolder');
    assert.strictEqual(actual, 'Shared Documents/MyFolder');
  });

  it('correctly removes trailing slashes from the URL', () => {
    const actual = urlUtil.removeTrailingSlashes('/Shared Documents/MyFolder/');
    assert.strictEqual(actual, '/Shared Documents/MyFolder');
  });
});