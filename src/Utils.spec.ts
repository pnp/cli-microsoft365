import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from './Utils';
import appInsights from './appInsights';
import auth from './Auth';

describe('Utils', () => {
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('isValidISODate returns true if value is in ISO Date format with - separator', () => {
    const result = Utils.isValidISODate("2019-03-22");
    assert.strictEqual(result, true);
  });

  it('isValidISODate returns true if value is in ISO Date format with . separator', () => {
    const result = Utils.isValidISODate("2019.03.22");
    assert.strictEqual(result, true);
  });

  it('isValidISODate returns true if value is in ISO Date format with / separator', () => {
    const result = Utils.isValidISODate("2019/03/22");
    assert.strictEqual(result, true);
  });

  it('isValidISODate returns false if value is blank', () => {
    const result = Utils.isValidISODate("");
    assert.strictEqual(result, false);
  });

  it('isValidISODate returns false if value is not in ISO Date format', () => {
    const result = Utils.isValidISODate("22-03-2019");
    assert.strictEqual(result, false);
  });

  it('isValidISODate returns false if alpha characters are passed', () => {
    const result = Utils.isValidISODate("sharing is caring");
    assert.strictEqual(result, false);
  });

  it('isValidISODateDashOnly returns true if value is in ISO Date format with - separator', () => {
    const result = Utils.isValidISODateDashOnly("2019-03-22");
    assert.strictEqual(result, true);
  });

  it('isValidISODateDashOnly returns false if value is in ISO Date format with . separator', () => {
    const result = Utils.isValidISODateDashOnly("2019.03.22");
    assert.strictEqual(result, false);
  });

  it('isValidISODateDashOnly returns false if value is in ISO Date format with / separator', () => {
    const result = Utils.isValidISODateDashOnly("2019/03/22");
    assert.strictEqual(result, false);
  });

  it('isValidISODateDashOnly returns false if value is blank', () => {
    const result = Utils.isValidISODateDashOnly("");
    assert.strictEqual(result, false);
  });

  it('isValidISODate returns false if value is not in ISO Date format', () => {
    const result = Utils.isValidISODate("22-03-2019");
    assert.strictEqual(result, false);
  });

  it('isValidISODateDashOnly returns false if alpha characters are passed', () => {
    const result = Utils.isValidISODateDashOnly("sharing is caring");
    assert.strictEqual(result, false);
  });

  it('isDateInRange returns true if date within monthOffset is passed', () => {
    let d: Date = new Date()
    d.setMonth(d.getMonth() - 1);
    const result = Utils.isDateInRange(d.toISOString(), 2);
    assert.strictEqual(result, true);
  });

  it('isDateInRange returns false if date prior to monthOffset is passed', () => {
    let d: Date = new Date()
    d.setMonth(d.getMonth() - 2);
    const result = Utils.isDateInRange(d.toISOString(), 1);
    assert.strictEqual(result, false);
  });

  it('isDateInRange returns false if alpha characters are passed', () => {
    const result = Utils.isDateInRange("sharing is caring", 1);
    assert.strictEqual(result, false);
  });

  it('should validate a valid date without time is passed', () => {
    const result = Utils.isValidISODateTime("2019-01-01");
    assert.strictEqual(result, true);
  });

  it('should validate a valid date with only hours-precision time is passed', () => {
    const result = Utils.isValidISODateTime("2019-01-01T01Z");
    assert.strictEqual(result, true);
  });

  it('should validate a valid date with only minutes-precision time is passed', () => {
    const result = Utils.isValidISODateTime("2019-01-01T01:01Z");
    assert.strictEqual(result, true);
  });

  it('should validate a valid date with only seconds-precision time is passed', () => {
    const result = Utils.isValidISODateTime("2019-01-01T01:01:01Z");
    assert.strictEqual(result, true);
  });

  it('should validate a valid date with milliseconds-precision time is passed', () => {
    const result = Utils.isValidISODateTime("2019-01-01T01:01:01.123Z");
    assert.strictEqual(result, true);
  });

  it('isValidGuid returns true if valid guid', () => {
    const result = Utils.isValidGuid('b2307a39-e878-458b-bc90-03bc578531d6');
    assert.strictEqual(result, true);
  });

  it('isValidGuid returns false if invalid guid', () => {
    const result = Utils.isValidGuid('b2307a39-e878-458b-bc90-03bc578531dw');
    assert(result == false);
  });

  it('isValidTeamsChannelId returns true if valid channelId (all numbers)', () => {
    const result = Utils.isValidTeamsChannelId('19:0000000000000000000000000000000@thread.skype');
    assert.strictEqual(result, true);
  });

  it('isValidTeamsChannelId returns true if valid channelId (numbers and letters)', () => {
    const result = Utils.isValidTeamsChannelId('19:ABZTZ000000000000000000000rstfv@thread.skype');
    assert.strictEqual(result, true);
  });

  it('isValidTeamsChannelId returns true if valid channelId with new tacv2 format', () => {
    const result = Utils.isValidTeamsChannelId('19:ABZTZ000000000000000000000rstfv@thread.tacv2');
    assert.strictEqual(result, true);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (missing colon)', () => {
    const result = Utils.isValidTeamsChannelId('190000000000000000000000000000000@thread.skype');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (starting with one digit)', () => {
    const result = Utils.isValidTeamsChannelId('1:0000000000000000000000000000000@thread.skype');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (starting with two digits but not 19)', () => {
    const result = Utils.isValidTeamsChannelId('18:0000000000000000000000000000000@thread.skype');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (missing @)', () => {
    const result = Utils.isValidTeamsChannelId('19:0000000000000000000000000000000thread.skype');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (doesn\'t end with skype)', () => {
    const result = Utils.isValidTeamsChannelId('19:0000000000000000000000000000000@thread.skype1');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (no . between thread and skype)', () => {
    const result = Utils.isValidTeamsChannelId('19:0000000000000000000000000000000@threadskype');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (doesn\'t end with thread.skype)', () => {
    const result = Utils.isValidTeamsChannelId('19:0000000000000000000000000000000@threadaskype');
    assert.strictEqual(result, false);
  });

  it('isValidBoolean returns true if valid boolean', () => {
    const result = Utils.isValidBoolean('true');
    assert.strictEqual(result, true);
  });

  it('isValidBoolean returns false if invalid boolean', () => {
    const result = Utils.isValidBoolean('foo');
    assert(result == false);
  });

  it('doesn\'t fail when restoring stub if the passed object is undefined', () => {
    Utils.restore(undefined);
    assert(true);
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/team1', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/team1', '');
    assert.strictEqual(actual, '/sites/team1');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/team1/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/team1/', '');
    assert.strictEqual(actual, '/sites/team1');
  });

  it('should get server relative path when https://contoso.sharepoint.com/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/', '');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path when domain only', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com', '');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path when /sites/team1 relative path passed as param', () => {
    const actual = Utils.getServerRelativePath('/sites/team1', '');
    assert.strictEqual(actual, '/sites/team1');
  });

  it('should get server relative path when /sites/team1/ relative path passed as param', () => {
    const actual = Utils.getServerRelativePath('/sites/team1/', '');
    assert.strictEqual(actual, '/sites/team1');
  });

  it('should get server relative path when / relative path passed as param', () => {
    const actual = Utils.getServerRelativePath('/', '');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path for https://contoso.sharepoint.com/sites/team1 and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/team1', 'Shared Documents');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/ and /Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/team1/', '/Shared Documents');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/ and /Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when domain only and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when /sites/team1 and /Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/sites/team1', '/Shared Documents');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when /sites/team1 and /Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/team1', '/Shared Documents/');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when /sites/team1/ and /Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/team1/', '/Shared Documents/');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when sites/team1/ and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/team1/', 'Shared Documents');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when sites/team1/ and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/team1', 'Shared Documents');
    assert.strictEqual(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when / and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when / and /Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when / and /Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/', '/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when "" and ""', () => {
    const actual = Utils.getServerRelativePath('', '');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path when / and /', () => {
    const actual = Utils.getServerRelativePath('/', '/');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path when "" and /', () => {
    const actual = Utils.getServerRelativePath('', '/');
    assert.strictEqual(actual, '/');
  });

  it('should get server relative path when "" and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1/', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/sites/site1/', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/sites/site1/', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/site1/', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/site1/', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when uppercase in web url e.g. sites/Site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/Site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/sites/Site1/Shared Documents');
  });

  it('should get server relative path when uppercase in folder url e.g. sites/site1 and /sites/Site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/Site1/Shared Documents');
    assert.strictEqual(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sub folder present url e.g. sites/site1 and /sites/Site1/Shared Documents/MyFolder', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/Site1/Shared Documents/MyFolder');
    assert.strictEqual(actual, '/sites/site1/Shared Documents/MyFolder');
  });

  it('should get web relative path when / relative path passed as param', () => {
    const actual = Utils.getWebRelativePath('/', '/');
    assert.strictEqual(actual, '/');
  });

  it('should get web relative path for https://contoso.sharepoint.com/sites/team1 and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/team1', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/ and /Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/team1/', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/ and /Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when domain only and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/team1 and /Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/sites/team1', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/team1 and /Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/team1', '/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/team1/ and /Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/team1/', '/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/team1/ and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/team1/', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/team1/ and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/team1', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when / and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when / and /Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/', '/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when / and /Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/', '/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when "" and ""', () => {
    const actual = Utils.getWebRelativePath('', '');
    assert.strictEqual(actual, '/');
  });

  it('should get web relative path when / and /', () => {
    const actual = Utils.getWebRelativePath('/', '/');
    assert.strictEqual(actual, '/');
  });

  it('should get web relative path when "" and /', () => {
    const actual = Utils.getWebRelativePath('', '/');
    assert.strictEqual(actual, '/');
  });

  it('should get web relative path when "" and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('', 'Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1/', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/sites/site1/', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/sites/site1/', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/site1/', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/site1/', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when uppercase in web url e.g. sites/Site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/Site1', '/sites/site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when uppercase in folder url e.g. sites/site1 and /sites/Site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/Site1/Shared Documents');
    assert.strictEqual(actual, '/Shared Documents');
  });

  it('should get web relative path when sub folder present url e.g. sites/site1 and /sites/Site1/Shared Documents/MyFolder', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/Site1/Shared Documents/MyFolder');
    assert.strictEqual(actual, '/Shared Documents/MyFolder');
  });

  it('should get absolute URL of a folder using webUrl and the folder server relative url', () => {
    const actual = Utils.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1', '/sites/team1/Shared Documents/MyFolder');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('should handle the server relative url starting by / or not while getting absolute URL of a folder', () => {
    const actual = Utils.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1', 'sites/team1/Shared Documents/MyFolder');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('should handle the presence of an ending / on the web url while getting absolute URL of a folder', () => {
    const actual = Utils.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1/', 'sites/team1/Shared Documents/MyFolder');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('should properly concatenate URL parts even with ending and starting / to each while getting absolute URL of a folder', () => {
    const actual = Utils.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1/', '/sites/team1/Shared Documents/MyFolder');
    assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('shows app display name as connected-as for app-only auth', () => {
    const jwt = JSON.stringify({
      app_displayname: 'CLI for Microsoft 365 Contoso'
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;
    const actual = Utils.getUserNameFromAccessToken(accessToken);
    assert.strictEqual(actual, 'CLI for Microsoft 365 Contoso');
  });

  it('returns empty user name when access token is undefined', () => {
    const actual = Utils.getUserNameFromAccessToken(undefined as any);
    assert.strictEqual(actual, '');
  });

  it('returns empty user name when empty access token passed', () => {
    const actual = Utils.getUserNameFromAccessToken('');
    assert.strictEqual(actual, '');
  });

  it('returns empty user name when invalid access token passed', () => {
    const actual = Utils.getUserNameFromAccessToken('abc.def.ghi');
    assert.strictEqual(actual, '');
  });

  it('shows tenant id from valid access token', () => {
    const jwt = JSON.stringify({
      tid: 'de349bc7-1aeb-4506-8cb3-98db021cadb4'
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;
    const actual = Utils.getTenantIdFromAccessToken(accessToken);
    assert.strictEqual(actual, 'de349bc7-1aeb-4506-8cb3-98db021cadb4');
  });

  it('returns empty tenant id when access token is undefined', () => {
    const actual = Utils.getTenantIdFromAccessToken(undefined as any);
    assert.strictEqual(actual, '');
  });

  it('returns empty tenant id when empty access token passed', () => {
    const actual = Utils.getTenantIdFromAccessToken('');
    assert.strictEqual(actual, '');
  });

  it('returns empty tenant id when invalid access token passed', () => {
    const actual = Utils.getTenantIdFromAccessToken('abc.def.ghi');
    assert.strictEqual(actual, '');
  });

  it('isJavascriptReservedWord returns true if value equals a JavaScript Reserved Word (eg. onload)', () => {
    const result = Utils.isJavascriptReservedWord('onload');
    assert.strictEqual(result, true);
  });

  it('isJavascriptReservedWord returns false if value doesn\'t equal a JavaScript Reserved Word due to case sensitivity (eg. onLoad)', () => {
    const result = Utils.isJavascriptReservedWord('onLoad');
    assert.strictEqual(result, false);
  });

  it('isJavascriptReservedWord returns false if value doesn\'t equal a JavaScript Reserved Word', () => {
    const result = Utils.isJavascriptReservedWord('exampleword');
    assert.strictEqual(result, false);
  });

  it('isJavascriptReservedWord returns false if value contains but doesn\'t equal a JavaScript Reserved Word (eg. encodeURIComponent)', () => {
    const result = Utils.isJavascriptReservedWord('examplewordencodeURIComponent');
    assert.strictEqual(result, false);
  });

  it('isJavascriptReservedWord returns true if any part of a value, when split on dot, equals a JavaScript Reserved Word (eg. innerHeight)', () => {
    const result = Utils.isJavascriptReservedWord('exampleword.innerHeight.anotherpart');
    assert.strictEqual(result, true);
  });

  it('isJavascriptReservedWord returns false if any part of a value, when split on dot, doesn\'t equal a JavaScript Reserved Word', () => {
    const result = Utils.isJavascriptReservedWord('exampleword.secondsection.anotherpart');
    assert.strictEqual(result, false);
  });

  it('isJavascriptReservedWord returns false if any part of a value, when split on dot, contains but doesn\'t equal a JavaScript Reserved Word (eg. layer)', () => {
    const result = Utils.isJavascriptReservedWord('exampleword.layersecondsection.anotherpart');
    assert.strictEqual(result, false);
  });

  it('should get safe filename when file\'name.txt', () => {
    const result = Utils.getSafeFileName('file\'name.txt');
    assert.strictEqual(result, 'file\'\'name.txt');
  });

  it('isValidTheme returns true when valid theme is passed', () => {
    const theme = `{
        "themePrimary": "#d81e05",
        "themeLighterAlt": "#fdf5f4",
        "themeLighter": "#f9d6d2",
        "themeLight": "#f4b4ac",
        "themeTertiary": "#e87060",
        "themeSecondary": "#dd351e",
        "themeDarkAlt": "#c31a04",
        "themeDark": "#a51603",
        "themeDarker": "#791002",
        "neutralLighterAlt": "#eeeeee",
        "neutralLighter": "#f5f5f5",
        "neutralLight": "#e1e1e1",
        "neutralQuaternaryAlt": "#d1d1d1",
        "neutralQuaternary": "#c8c8c8",
        "neutralTertiaryAlt": "#c0c0c0",
        "neutralTertiary": "#c2c2c2",
        "neutralSecondary": "#858585",
        "neutralPrimaryAlt": "#4b4b4b",
        "neutralPrimary": "#333333",
        "neutralDark": "#272727",
        "black": "#1d1d1d",
        "white": "#f5f5f5"
    }`;
    const actual = Utils.isValidTheme(theme);
    const expected = true;
    assert.strictEqual(actual, expected);
  });

  it('isValidTheme returns false when theme passed is not valid json', () => {
    const theme = `{ not valid }`;
    const actual = Utils.isValidTheme(theme);
    const expected = false;
    assert.strictEqual(actual, expected);
  });

  it('isValidTheme returns false when theme passed is not a json object', () => {
    const theme = `[{
        "themePrimary": "#d81e05",
        "themeLighterAlt": "#fdf5f4",
        "themeLighter": "#f9d6d2",
        "themeLight": "#f4b4ac",
        "themeTertiary": "#e87060",
        "themeSecondary": "#dd351e",
        "themeDarkAlt": "#c31a04",
        "themeDark": "#a51603",
        "themeDarker": "#791002",
        "neutralLighterAlt": "#eeeeee",
        "neutralLighter": "#f5f5f5",
        "neutralLight": "#e1e1e1",
        "neutralQuaternaryAlt": "#d1d1d1",
        "neutralQuaternary": "#c8c8c8",
        "neutralTertiaryAlt": "#c0c0c0",
        "neutralTertiary": "#c2c2c2",
        "neutralSecondary": "#858585",
        "neutralPrimaryAlt": "#4b4b4b",
        "neutralPrimary": "#333333",
        "neutralDark": "#272727",
        "black": "#1d1d1d",
        "white": "#f5f5f5"
    }]`;
    const actual = Utils.isValidTheme(theme);
    const expected = false;
    assert.strictEqual(actual, expected);
  });

  it('isValidTheme returns false when theme passed does not contain all valid properties', () => {
    const theme = `{
        "themeLighterAlt": "#fdf5f4",
        "themeLighter": "#f9d6d2",
        "themeLight": "#f4b4ac",
        "themeTertiary": "#e87060",
        "themeSecondary": "#dd351e",
        "themeDarkAlt": "#c31a04",
        "themeDark": "#a51603",
        "themeDarker": "#791002",
        "neutralLighterAlt": "#eeeeee",
        "neutralLighter": "#f5f5f5",
        "neutralLight": "#e1e1e1",
        "neutralQuaternaryAlt": "#d1d1d1",
        "neutralQuaternary": "#c8c8c8",
        "neutralTertiaryAlt": "#c0c0c0",
        "neutralTertiary": "#c2c2c2",
        "neutralSecondary": "#858585",
        "neutralPrimaryAlt": "#4b4b4b",
        "neutralPrimary": "#333333",
        "neutralDark": "#272727",
        "black": "#1d1d1d",
        "white": "#f5f5f5"
    }`;
    const actual = Utils.isValidTheme(theme);
    const expected = false;
    assert.strictEqual(actual, expected);
  });

  it('isValidTheme returns false when theme passed contains additional properties', () => {
    const theme = `{
        "additionalProperty": "#aaabbb",
        "themeLighterAlt": "#fdf5f4",
        "themeLighter": "#f9d6d2",
        "themeLight": "#f4b4ac",
        "themeTertiary": "#e87060",
        "themeSecondary": "#dd351e",
        "themeDarkAlt": "#c31a04",
        "themeDark": "#a51603",
        "themeDarker": "#791002",
        "neutralLighterAlt": "#eeeeee",
        "neutralLighter": "#f5f5f5",
        "neutralLight": "#e1e1e1",
        "neutralQuaternaryAlt": "#d1d1d1",
        "neutralQuaternary": "#c8c8c8",
        "neutralTertiaryAlt": "#c0c0c0",
        "neutralTertiary": "#c2c2c2",
        "neutralSecondary": "#858585",
        "neutralPrimaryAlt": "#4b4b4b",
        "neutralPrimary": "#333333",
        "neutralDark": "#272727",
        "black": "#1d1d1d",
        "white": "#f5f5f5"
    }`;
    const actual = Utils.isValidTheme(theme);
    const expected = false;
    assert.strictEqual(actual, expected);
  });

  it('isValidTheme returns false when theme passed does not contain valid hex color value', () => {
    const theme = `{
        "themePrimary": "d81e05",
        "themeLighterAlt": "#fdf5f4",
        "themeLighter": "#f9d6d2",
        "themeLight": "#f4b4ac",
        "themeTertiary": "#e87060",
        "themeSecondary": "#dd351e",
        "themeDarkAlt": "#c31a04",
        "themeDark": "#a51603",
        "themeDarker": "#791002",
        "neutralLighterAlt": "#eeeeee",
        "neutralLighter": "#f5f5f5",
        "neutralLight": "#e1e1e1",
        "neutralQuaternaryAlt": "#d1d1d1",
        "neutralQuaternary": "#c8c8c8",
        "neutralTertiaryAlt": "#c0c0c0",
        "neutralTertiary": "#c2c2c2",
        "neutralSecondary": "#858585",
        "neutralPrimaryAlt": "#4b4b4b",
        "neutralPrimary": "#333333",
        "neutralDark": "#272727",
        "black": "#1d1d1d",
        "white": "#f5f5f5"
    }`;
    const actual = Utils.isValidTheme(theme);
    const expected = false;
    assert.strictEqual(actual, expected);
  });

  it('isValidTheme returns false when theme passed is not valid (issue #1463)', () => {
    const theme = `{
      "Palette": {
        "themePrimary": "#d81e05",
        "themeLighterAlt": "#fdf5f4",
        "themeLighter": "#f9d6d2",
        "themeLight": "#f4b4ac",
        "themeTertiary": "#e87060",
        "themeSecondary": "#dd351e",
        "themeDarkAlt": "#c31a04",
        "themeDark": "#a51603",
        "themeDarker": "#791002",
        "neutralLighterAlt": "#eeeeee",
        "neutralLighter": "#f5f5f5",
        "neutralLight": "#e1e1e1",
        "neutralQuaternaryAlt": "#d1d1d1",
        "neutralQuaternary": "#c8c8c8",
        "neutralTertiaryAlt": "#c0c0c0",
        "neutralTertiary": "#c2c2c2",
        "neutralSecondary": "#858585",
        "neutralPrimaryAlt": "#4b4b4b",
        "neutralPrimary": "#333333",
        "neutralDark": "#272727",
        "black": "#1d1d1d",
        "white": "#f5f5f5"
      }
    }`;
    const actual = Utils.isValidTheme(theme);
    const expected = false;
    assert.strictEqual(actual, expected);
  });
});