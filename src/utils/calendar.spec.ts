import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../cli/cli.js';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { calendar } from './calendar.js';
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';

describe('utils/calendar', () => {
  const userId = '729827e3-9c14-49f7-bb1b-9608f156bbb8';
  const calendarId = 'AAMkAGI2TGuLAAA';
  const calendarName = 'My Calendar';
  const invalidCalendarName = 'M Calnedar';
  const calendarGroupId = 'AQMkADIxYjJiYgEzLTFmN_F8AAAIBBgAA_F8AAAJjIQAAAA==';
  const calendarResponse = {
    "id": "AAMkAGI2TGuLAAA=",
    "name": "Calendar",
    "color": "auto",
    "isDefaultCalendar": true,
    "changeKey": "nfZyf7VcrEKLNoU37KWlkQAAA0x0+w==",
    "canShare": true,
    "canViewPrivateItems": true,
    "hexColor": "",
    "canEdit": true,
    "allowedOnlineMeetingProviders": [
      "teamsForBusiness"
    ],
    "defaultOnlineMeetingProvider": "teamsForBusiness",
    "isTallyingResponses": true,
    "isRemovable": false,
    "owner": {
      "name": "John Doe",
      "address": "john.doe@contoso.com"
    }
  };
  const anotherCalendarResponse = {
    "id": "AAMkAGI2TGuLBBB=",
    "name": "Vacation",
    "color": "auto",
    "isDefaultCalendar": false,
    "changeKey": "abcdf7VcrEKLNoU37KWlkQAAA0x0+w==",
    "canShare": false,
    "canViewPrivateItems": true,
    "hexColor": "",
    "canEdit": true,
    "allowedOnlineMeetingProviders": [
    ],
    "defaultOnlineMeetingProvider": "none",
    "isTallyingResponses": true,
    "isRemovable": false,
    "owner": {
      "name": "John Doe",
      "address": "john.doe@contoso.com"
    }
  };
  const calendarLimitedResponse = {
    "id": "AAMkAGI2TGuLAAA=",
    "name": "Calendar",
    "color": "auto"
  };

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get single calendar by name using getUserCalendarByName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars?$filter=name eq '${formatting.encodeQueryParameter(calendarName)}'`) {
        return {
          value: [
            calendarResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    const actual = await calendar.getUserCalendarByName(userId, calendarName);
    assert.deepStrictEqual(actual, calendarResponse);
  });

  it('correctly get single calendar by name from a calendar group using getUserCalendarByName with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}/calendars?$filter=name eq '${formatting.encodeQueryParameter(calendarName)}'&$select=id,name`) {
        return {
          value: [
            calendarLimitedResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    const actual = await calendar.getUserCalendarByName(userId, calendarName, calendarGroupId, 'id,name');
    assert.deepStrictEqual(actual, calendarLimitedResponse);
  });

  it('handles selecting single calendar when multiple calendars with the specified name found using getUserCalendarByName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars?$filter=name eq '${formatting.encodeQueryParameter(calendarName)}'`) {
        return {
          value: [
            calendarResponse,
            anotherCalendarResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(calendarResponse);

    const actual = await calendar.getUserCalendarByName(userId, calendarName);
    assert.deepStrictEqual(actual, calendarResponse);
  });

  it('throws error message when no calendar was found using getUserCalendarByName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars?$filter=name eq '${formatting.encodeQueryParameter(invalidCalendarName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(calendar.getUserCalendarByName(userId, invalidCalendarName),
      new Error(`The specified calendar '${invalidCalendarName}' does not exist.`));
  });

  it('throws error message when multiple calendars were found using getUserCalendarByName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars?$filter=name eq '${formatting.encodeQueryParameter(calendarName)}'`) {
        return {
          value: [
            calendarResponse,
            anotherCalendarResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(calendar.getUserCalendarByName(userId, calendarName),
      Error(`Multiple calendars with name '${calendarName}' found. Found: ${calendarResponse.id}, ${anotherCalendarResponse.id}.`));
  });

  it('correctly get single calendar by id using getUserCalendarById', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}`) {
        return calendarResponse;
      }

      throw 'Invalid Request';
    });

    const actual = await calendar.getUserCalendarById(userId, calendarId);
    assert.deepStrictEqual(actual, calendarResponse);
  });

  it('correctly get single calendar by id from a calendar group using getUserCalendarById with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}/calendars/${calendarId}?$select=id,displayName`) {
        return calendarLimitedResponse;
      }

      throw 'Invalid Request';
    });

    const actual = await calendar.getUserCalendarById(userId, calendarId, calendarGroupId, 'id,displayName');
    assert.deepStrictEqual(actual, calendarLimitedResponse);
  });
});