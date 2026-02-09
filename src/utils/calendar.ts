import { Calendar } from '@microsoft/microsoft-graph-types';
import { odata } from './odata.js';
import { formatting } from './formatting.js';
import { cli } from '../cli/cli.js';
import request, { CliRequestOptions } from '../request.js';

export const calendar = {
  async getUserCalendarById(userId: string, calendarId:string, calendarGroupId?: string, properties?: string): Promise<Calendar> {
    let url = `https://graph.microsoft.com/v1.0/users('${userId}')/${calendarGroupId ? `calendarGroups/${calendarGroupId}/` : ''}calendars/${calendarId}`;

    if (properties) {
      url += `?$select=${properties}`;
    }

    const requestOptions: CliRequestOptions = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return await request.get<Calendar>(requestOptions);
  },

  async getUserCalendarByName(userId: string, name: string, calendarGroupId?: string, properties?: string): Promise<Calendar> {
    let url = `https://graph.microsoft.com/v1.0/users('${userId}')/${calendarGroupId ? `calendarGroups/${calendarGroupId}/` : ''}calendars?$filter=name eq '${formatting.encodeQueryParameter(name)}'`;

    if (properties) {
      url += `&$select=${properties}`;
    }

    const calendars = await odata.getAllItems<Calendar>(url);

    if (calendars.length === 0) {
      throw new Error(`The specified calendar '${name}' does not exist.`);
    }

    if (calendars.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', calendars );
      const selectedCalendar = await cli.handleMultipleResultsFound<Calendar>(`Multiple calendars with name '${name}' found.`, resultAsKeyValuePair);
      return selectedCalendar;
    }

    return calendars[0];
  }
};
