import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../cli/cli.js';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { calendarGroup } from './calendarGroup.js';
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';

describe('utils/calendarGroup', () => {
  const userId = '729827e3-9c14-49f7-bb1b-9608f156bbb8';
  const groupName = 'My Calendars';
  const invalidGroupName = 'M Calnedar';
  const calendarGroupResponse = {
    "name": "My Calendars",
    "classId": "0006f0b7-0000-0000-c000-000000000046",
    "changeKey": "NreqLYgxdE2DpHBBId74XwAAAAAGZw==",
    "id": "AQMkADIxYjJiYgEzLTFmN_F8AAAIBBgAA_F8AAAJjIQAAAA=="
  };
  const anotherCalendarGroupResponse = {
    "name": "My Calendars",
    "classId": "0006f0b7-0000-0000-c000-000000000047",
    "changeKey": "MreqLYgxdE2DpHBBId74XwAAAAAGZw==",
    "id": "AQMkADIxYjJiYgEzLTFmN_F8AAAIBBgAA_F8AAAJjIQBBB=="
  };
  const calendarGroupLimitedResponse = {
    "name": "My Calendars",
    "id": "AQMkADIxYjJiYgEzLTFmN_F8AAAIBBgAA_F8AAAJjIQAAAA=="
  };

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get single calendar group by name using getUserCalendarGroupByName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups?$filter=name eq '${formatting.encodeQueryParameter(groupName)}'`) {
        return {
          value: [
            calendarGroupResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    const actual = await calendarGroup.getUserCalendarGroupByName(userId, groupName);
    assert.deepStrictEqual(actual, calendarGroupResponse);
  });

  it('correctly get single calendar group by name using getUserCalendarGroupByName with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups?$filter=name eq '${formatting.encodeQueryParameter(groupName)}'&$select=id,name`) {
        return {
          value: [
            calendarGroupLimitedResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    const actual = await calendarGroup.getUserCalendarGroupByName(userId, groupName, 'id,name');
    assert.deepStrictEqual(actual, calendarGroupLimitedResponse);
  });

  it('handles selecting single calendar group when multiple calendar groups with the specified name found using getUserCalendarGroupByName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups?$filter=name eq '${formatting.encodeQueryParameter(groupName)}'`) {
        return {
          value: [
            calendarGroupResponse,
            anotherCalendarGroupResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(calendarGroupResponse);

    const actual = await calendarGroup.getUserCalendarGroupByName(userId, groupName);
    assert.deepStrictEqual(actual, calendarGroupResponse);
  });

  it('throws error message when no calendar group was found using getUserCalendarGroupByName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups?$filter=name eq '${formatting.encodeQueryParameter(invalidGroupName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(calendarGroup.getUserCalendarGroupByName(userId, invalidGroupName),
      new Error(`The specified calendar group '${invalidGroupName}' does not exist.`));
  });

  it('throws error message when multiple calendar groups were found using getUserCalendarGroupByName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups?$filter=name eq '${formatting.encodeQueryParameter(groupName)}'`) {
        return {
          value: [
            calendarGroupResponse,
            anotherCalendarGroupResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(calendarGroup.getUserCalendarGroupByName(userId, groupName),
      Error(`Multiple calendar groups with name '${groupName}' found. Found: ${calendarGroupResponse.id}, ${anotherCalendarGroupResponse.id}.`));
  });
});