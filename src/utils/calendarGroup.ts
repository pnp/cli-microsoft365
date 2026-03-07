import { CalendarGroup } from '@microsoft/microsoft-graph-types';
import { odata } from './odata.js';
import { formatting } from './formatting.js';
import { cli } from '../cli/cli.js';

export const calendarGroup = {
  async getUserCalendarGroupByName(userId: string, displayName: string, properties?: string): Promise<CalendarGroup> {
    let url = `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups?$filter=name eq '${formatting.encodeQueryParameter(displayName)}'`;

    if (properties) {
      url += `&$select=${properties}`;
    }

    const calendarGroups = await odata.getAllItems<CalendarGroup>(url);

    if (calendarGroups.length === 0) {
      throw new Error(`The specified calendar group '${displayName}' does not exist.`);
    }

    if (calendarGroups.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', calendarGroups);
      const selectedCalendarGroup = await cli.handleMultipleResultsFound<CalendarGroup>(`Multiple calendar groups with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedCalendarGroup;
    }

    return calendarGroups[0];
  }
};
