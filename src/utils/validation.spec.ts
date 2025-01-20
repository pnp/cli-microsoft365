import assert from 'assert';
import { validation } from '../utils/validation.js';

describe('validation/validation', () => {
  it('isValidISODate returns true if value is in ISO Date format with - separator', () => {
    const result = validation.isValidISODate("2019-03-22");
    assert.strictEqual(result, true);
  });

  it('isValidISODate returns true if value is in ISO Date format with . separator', () => {
    const result = validation.isValidISODate("2019.03.22");
    assert.strictEqual(result, true);
  });

  it('isValidISODate returns true if value is in ISO Date format with / separator', () => {
    const result = validation.isValidISODate("2019/03/22");
    assert.strictEqual(result, true);
  });

  it('isValidISODate returns false if value is blank', () => {
    const result = validation.isValidISODate("");
    assert.strictEqual(result, false);
  });

  it('isValidISODate returns false if value is not in ISO Date format', () => {
    const result = validation.isValidISODate("22-03-2019");
    assert.strictEqual(result, false);
  });

  it('isValidISODate returns false if alpha characters are passed', () => {
    const result = validation.isValidISODate("sharing is caring");
    assert.strictEqual(result, false);
  });

  it('isValidISODateDashOnly returns true if value is in ISO Date format with - separator', () => {
    const result = validation.isValidISODateDashOnly("2019-03-22");
    assert.strictEqual(result, true);
  });

  it('isValidISODateDashOnly returns false if value is in ISO Date format with . separator', () => {
    const result = validation.isValidISODateDashOnly("2019.03.22");
    assert.strictEqual(result, false);
  });

  it('isValidISODateDashOnly returns false if value is in ISO Date format with / separator', () => {
    const result = validation.isValidISODateDashOnly("2019/03/22");
    assert.strictEqual(result, false);
  });

  it('isValidISODateDashOnly returns false if value is blank', () => {
    const result = validation.isValidISODateDashOnly("");
    assert.strictEqual(result, false);
  });

  it('isValidISODateDashOnly returns false if alpha characters are passed', () => {
    const result = validation.isValidISODateDashOnly("sharing is caring");
    assert.strictEqual(result, false);
  });

  it('isDateInRange returns true if date within monthOffset is passed', () => {
    const d: Date = new Date();
    d.setMonth(d.getMonth() - 1);
    const result = validation.isDateInRange(d.toISOString(), 2);
    assert.strictEqual(result, true);
  });

  it('isDateInRange returns false if date prior to monthOffset is passed', () => {
    const d: Date = new Date();
    d.setMonth(d.getMonth() - 2);
    const result = validation.isDateInRange(d.toISOString(), 1);
    assert.strictEqual(result, false);
  });

  it('isDateInRange returns false if alpha characters are passed', () => {
    const result = validation.isDateInRange("sharing is caring", 1);
    assert.strictEqual(result, false);
  });

  it('should validate a valid date without time is passed', () => {
    const result = validation.isValidISODateTime("2019-01-01");
    assert.strictEqual(result, true);
  });

  it('should validate a valid date with only hours-precision time is passed', () => {
    const result = validation.isValidISODateTime("2019-01-01T01Z");
    assert.strictEqual(result, true);
  });

  it('should validate a valid date with only minutes-precision time is passed', () => {
    const result = validation.isValidISODateTime("2019-01-01T01:01Z");
    assert.strictEqual(result, true);
  });

  it('should validate a valid date with only seconds-precision time is passed', () => {
    const result = validation.isValidISODateTime("2019-01-01T01:01:01Z");
    assert.strictEqual(result, true);
  });

  it('should validate a valid date with milliseconds-precision time is passed (short)', () => {
    const result = validation.isValidISODateTime("2019-01-01T01:01:01.123Z");
    assert.strictEqual(result, true);
  });

  it('should validate a valid date with milliseconds-precision time is passed (long)', () => {
    const result = validation.isValidISODateTime("2019-01-01T01:01:01.1234567Z");
    assert.strictEqual(result, true);
  });

  it('isValidGuidArray returns true if guids are valid', () => {
    const result = validation.isValidGuidArray('16c578ea-5785-492e-ac22-cad3cd9ca1fa,16cd5c6b-e9e9-4364-b71e-1a1664f81b98,7c9a1059-a725-424c-9dd0-788e66a5338e,02e83c70-f05f-4e63-b9af-73a8e44fdb32,5a53c7d7-2a26-4645-a938-b3e4d08b4a18');
    assert.strictEqual(result, true);
  });

  it('isValidGuidArray returns guids that are not valid', () => {
    const result = validation.isValidGuidArray('000,16c578ea-5785-492e-ac22-cad3cd9ca1fz,16cd5c6b-e9e9-4364-b71e-1a1664f81b98,7c9a1059-a725-424c-9dd0-788e66a5338e,02e83c70-f05f-4e63-b9af-73a8e44fdb32,5a53c7d7-2a26-4645-a938-b3e4d08b4a18');
    assert.strictEqual(result, "000, 16c578ea-5785-492e-ac22-cad3cd9ca1fz");
  });

  it('isValidGuid returns true if valid guid', () => {
    const result = validation.isValidGuid('b2307a39-e878-458b-bc90-03bc578531d6');
    assert.strictEqual(result, true);
  });

  it('isValidGuid returns false if guid is empty', () => {
    const result = validation.isValidGuid(undefined);
    assert.strictEqual(result, false);
  });

  it('isValidGuid returns false if invalid guid', () => {
    const result = validation.isValidGuid('b2307a39-e878-458b-bc90-03bc578531dw');
    assert(result === false);
  });

  it('isValidGuid returns true with @meid token', () => {
    const result = validation.isValidGuid('@meid');
    assert.strictEqual(result, true);
  });

  it('isValidGuid returns true with @meid token and spaces', () => {
    const result = validation.isValidGuid('@meid ');
    assert.strictEqual(result, true);
  });

  it('isValidGuid returns true with @meId (case sensitive)', () => {
    const result = validation.isValidGuid('@meId ');
    assert.strictEqual(result, true);
  });

  it('isValidUserPrincipalNameArray returns true if valid username array', () => {
    const result = validation.isValidUserPrincipalNameArray('john.doe@contoso.com,adele@contoso.onmicrosoft.com');
    assert.strictEqual(result, true);
  });

  it('isValidUserPrincipalNameArray returns falsy value of invalid username array', () => {
    const result = validation.isValidUserPrincipalNameArray('john.doe@contoso.com,foo');
    assert.strictEqual(result, 'foo');
  });

  it('isValidUserPrincipalNameArray returns true with @meusername token', () => {
    const result = validation.isValidUserPrincipalNameArray('john.doe@contoso.com,@meusername');
    assert.strictEqual(result, true);
  });

  it('isValidUserPrincipalNameArray returns true with @meusername token and spaces', () => {
    const result = validation.isValidUserPrincipalNameArray('john.doe@contoso.com,@meusername ');
    assert.strictEqual(result, true);
  });

  it('isValidUserPrincipalName returns true if valid username', () => {
    const result = validation.isValidUserPrincipalName('John@Contoso.com');
    assert.strictEqual(result, true);
  });

  it('isValidUserPrincipalName returns false if invalid username', () => {
    const result = validation.isValidUserPrincipalName('foo');
    assert.strictEqual(result, false);
  });

  it('isValidUserPrincipalName returns true with @meusername token', () => {
    const result = validation.isValidUserPrincipalName('@meusername');
    assert.strictEqual(result, true);
  });

  it('isValidUserPrincipalName returns true with @meusername token and spaces', () => {
    const result = validation.isValidUserPrincipalName('@meusername ');
    assert.strictEqual(result, true);
  });


  it('isValidUserPrincipalName returns true with @meusername (case sensitive)', () => {
    const result = validation.isValidUserPrincipalName('@meUsername ');
    assert.strictEqual(result, true);
  });

  it('isValidPositiveInteger returns true if valid integer as string', () => {
    const result = validation.isValidPositiveInteger('1');
    assert.strictEqual(result, true);
  });

  it('isValidPositiveInteger returns true if valid integer as number', () => {
    const result = validation.isValidPositiveInteger(1);
    assert.strictEqual(result, true);
  });

  it('isValidPositiveInteger returns error message of invalid integer when input is not a number', () => {
    const result = validation.isValidPositiveInteger('foo');
    assert.strictEqual(result, false);
  });

  it('isValidPositiveInteger returns error message of invalid integer when number not positive', () => {
    const result = validation.isValidPositiveInteger(-5);
    assert.strictEqual(result, false);
  });

  it('isValidPositiveIntegerArray returns true if valid integer array', () => {
    const result = validation.isValidPositiveIntegerArray('1, 2, 3, 4, 5');
    assert.strictEqual(result, true);
  });

  it('isValidPositiveIntegerArray returns error message of invalid integer when input is not a number', () => {
    const result = validation.isValidPositiveIntegerArray('1, 2, foo, 4, bar');
    assert.strictEqual(result, 'foo, bar');
  });

  it('isValidPositiveIntegerArray returns error message of invalid integer when number not positive', () => {
    const result = validation.isValidPositiveIntegerArray('0, 1, 2, 3, 4, -5');
    assert.strictEqual(result, '0, -5');
  });

  it('isValidTeamsChannelId returns true if valid channelId (all numbers)', () => {
    const result = validation.isValidTeamsChannelId('19:0000000000000000000000000000000@thread.skype');
    assert.strictEqual(result, true);
  });

  it('isValidTeamsChannelId returns true if valid channelId (numbers and letters)', () => {
    const result = validation.isValidTeamsChannelId('19:ABZTZ000000000000000000000rstfv@thread.skype');
    assert.strictEqual(result, true);
  });

  it('isValidTeamsChannelId returns true if valid channelId with new tacv2 format', () => {
    const result = validation.isValidTeamsChannelId('19:ABZTZ000000000000000000000rstfv@thread.tacv2');
    assert.strictEqual(result, true);
  });

  it('isValidTeamsChannelId returns true if channelId contains -', () => {
    const result = validation.isValidTeamsChannelId('19:ABZTZ00000000-0000000000000rstfv@thread.tacv2');
    assert.strictEqual(result, true);
  });

  it('isValidTeamsChannelId returns true if channelId contains _', () => {
    const result = validation.isValidTeamsChannelId('19:ABZTZ00000000_0000000000000rstfv@thread.tacv2');
    assert.strictEqual(result, true);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (missing colon)', () => {
    const result = validation.isValidTeamsChannelId('190000000000000000000000000000000@thread.skype');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (starting with one digit)', () => {
    const result = validation.isValidTeamsChannelId('1:0000000000000000000000000000000@thread.skype');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (starting with two digits but not 19)', () => {
    const result = validation.isValidTeamsChannelId('18:0000000000000000000000000000000@thread.skype');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (missing @)', () => {
    const result = validation.isValidTeamsChannelId('19:0000000000000000000000000000000thread.skype');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (doesn\'t end with skype)', () => {
    const result = validation.isValidTeamsChannelId('19:0000000000000000000000000000000@thread.skype1');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (no . between thread and skype)', () => {
    const result = validation.isValidTeamsChannelId('19:0000000000000000000000000000000@threadskype');
    assert.strictEqual(result, false);
  });

  it('isValidTeamsChannelId returns false if invalid channelId (doesn\'t end with thread.skype)', () => {
    const result = validation.isValidTeamsChannelId('19:0000000000000000000000000000000@threadaskype');
    assert.strictEqual(result, false);
  });

  it('isValidBoolean returns true if valid boolean', () => {
    const result = validation.isValidBoolean('true');
    assert.strictEqual(result, true);
  });

  it('isValidBoolean returns false if invalid boolean', () => {
    const result = validation.isValidBoolean('foo');
    assert(result === false);
  });

  it('isJavaScriptReservedWord returns true if value equals a JavaScript Reserved Word (eg. onload)', () => {
    const result = validation.isJavaScriptReservedWord('onload');
    assert.strictEqual(result, true);
  });

  it('isJavaScriptReservedWord returns false if value doesn\'t equal a JavaScript Reserved Word due to case sensitivity (eg. onLoad)', () => {
    const result = validation.isJavaScriptReservedWord('onLoad');
    assert.strictEqual(result, false);
  });

  it('isJavaScriptReservedWord returns false if value doesn\'t equal a JavaScript Reserved Word', () => {
    const result = validation.isJavaScriptReservedWord('exampleword');
    assert.strictEqual(result, false);
  });

  it('isJavaScriptReservedWord returns false if value contains but doesn\'t equal a JavaScript Reserved Word (eg. encodeURIComponent)', () => {
    const result = validation.isJavaScriptReservedWord('examplewordencodeURIComponent');
    assert.strictEqual(result, false);
  });

  it('isJavaScriptReservedWord returns true if any part of a value, when split on dot, equals a JavaScript Reserved Word (eg. innerHeight)', () => {
    const result = validation.isJavaScriptReservedWord('exampleword.innerHeight.anotherpart');
    assert.strictEqual(result, true);
  });

  it('isJavaScriptReservedWord returns false if any part of a value, when split on dot, doesn\'t equal a JavaScript Reserved Word', () => {
    const result = validation.isJavaScriptReservedWord('exampleword.secondsection.anotherpart');
    assert.strictEqual(result, false);
  });

  it('isJavaScriptReservedWord returns false if any part of a value, when split on dot, contains but doesn\'t equal a JavaScript Reserved Word (eg. layer)', () => {
    const result = validation.isJavaScriptReservedWord('exampleword.layersecondsection.anotherpart');
    assert.strictEqual(result, false);
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
    const actual = validation.isValidTheme(theme);
    const expected = true;
    assert.strictEqual(actual, expected);
  });

  it('isValidTheme returns false when theme passed is not valid json', () => {
    const theme = `{ not valid }`;
    const actual = validation.isValidTheme(theme);
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
    const actual = validation.isValidTheme(theme);
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
    const actual = validation.isValidTheme(theme);
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
    const actual = validation.isValidTheme(theme);
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
    const actual = validation.isValidTheme(theme);
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
    const actual = validation.isValidTheme(theme);
    const expected = false;
    assert.strictEqual(actual, expected);
  });

  it('isValidMailNickname returns true when mailNickname is valid', () => {
    const result = validation.isValidMailNickname('nickname');
    assert.strictEqual(result, true);
  });

  it('isValidMailNickname returns false when mailNickname contains \\', () => {
    const result = validation.isValidMailNickname('nicknam\\e');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains <', () => {
    const result = validation.isValidMailNickname('<nickname');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains >', () => {
    const result = validation.isValidMailNickname('nickname>');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains (', () => {
    const result = validation.isValidMailNickname('(nickname');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains )', () => {
    const result = validation.isValidMailNickname('nickname)');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains [', () => {
    const result = validation.isValidMailNickname('[nickname');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains ]', () => {
    const result = validation.isValidMailNickname('nickname]');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains @', () => {
    const result = validation.isValidMailNickname('nick@name');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains space', () => {
    const result = validation.isValidMailNickname('nick name');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains "', () => {
    const result = validation.isValidMailNickname('nick"name');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains ;', () => {
    const result = validation.isValidMailNickname('nick;name');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains :', () => {
    const result = validation.isValidMailNickname('nick:name');
    assert.strictEqual(result, false);
  });

  it('isValidMailNickname returns false when mailNickname contains ,', () => {
    const result = validation.isValidMailNickname('nick,name');
    assert.strictEqual(result, false);
  });

  it('isValidISODuration returns true if duration in years is valid', () => {
    const actual = validation.isValidISODuration("P2Y");
    const expected = true;
    assert.strictEqual(actual, expected);
  });

  it('isValidISODuration returns true if duration in months is valid', () => {
    const actual = validation.isValidISODuration("P2M");
    const expected = true;
    assert.strictEqual(actual, expected);
  });

  it('isValidISODuration returns true if duration in weeks is valid', () => {
    const actual = validation.isValidISODuration("P2W");
    const expected = true;
    assert.strictEqual(actual, expected);
  });

  it('isValidISODuration returns true if duration in days is valid', () => {
    const actual = validation.isValidISODuration("P2D");
    const expected = true;
    assert.strictEqual(actual, expected);
  });

  it('isValidISODuration returns true if duration in hours is valid', () => {
    const actual = validation.isValidISODuration("PT2H");
    const expected = true;
    assert.strictEqual(actual, expected);
  });

  it('isValidISODuration returns true if duration in minutes is valid', () => {
    const actual = validation.isValidISODuration("PT2M");
    const expected = true;
    assert.strictEqual(actual, expected);
  });

  it('isValidISODuration returns true if duration in seconds is valid', () => {
    const actual = validation.isValidISODuration("PT2S");
    const expected = true;
    assert.strictEqual(actual, expected);
  });

  it('isValidISODuration returns true if duration is valid', () => {
    const actual = validation.isValidISODuration("P3Y6M4DT12H30M5S");
    const expected = true;
    assert.strictEqual(actual, expected);
  });

  it('isValidISODuration returns false if duration in years is not valid', () => {
    const actual = validation.isValidISODuration("PY");
    const expected = true;
    assert.notStrictEqual(actual, expected);
  });

  it('isValidISODuration returns false if duration in months is not valid', () => {
    const actual = validation.isValidISODuration("PM");
    const expected = true;
    assert.notStrictEqual(actual, expected);
  });

  it('isValidISODuration returns false if duration in weeks is not valid', () => {
    const actual = validation.isValidISODuration("PW");
    const expected = true;
    assert.notStrictEqual(actual, expected);
  });

  it('isValidISODuration returns false if duration in days is not valid', () => {
    const actual = validation.isValidISODuration("PD");
    const expected = true;
    assert.notStrictEqual(actual, expected);
  });

  it('isValidISODuration returns false if duration in hours is not valid', () => {
    const actual = validation.isValidISODuration("PTH");
    const expected = true;
    assert.notStrictEqual(actual, expected);
  });

  it('isValidISODuration returns false if duration in minutes is not valid', () => {
    const actual = validation.isValidISODuration("PTM");
    const expected = true;
    assert.notStrictEqual(actual, expected);
  });

  it('isValidISODuration returns false if duration in seconds is not valid', () => {
    const actual = validation.isValidISODuration("PTS");
    const expected = true;
    assert.notStrictEqual(actual, expected);
  });

  it('isValidISODuration returns false if duration is not valid', () => {
    const actual = validation.isValidISODuration("P3Y6MDT12H30M5S");
    const expected = true;
    assert.notStrictEqual(actual, expected);
  });

  it('isValidPowerPagesUrl returns false if url is not a valid Power Pages website', () => {
    const actual = validation.isValidPowerPagesUrl("https://site-0uaq9.contoso.com");
    const expected = true;
    assert.notStrictEqual(actual, expected);
  });

  it('isValidPowerPagesUrl returns true if url is a valid Power Pages website', () => {
    const actual = validation.isValidPowerPagesUrl("https://site-0uaq9.powerappsportals.com");
    const expected = true;
    assert.strictEqual(actual, expected);
  });
});