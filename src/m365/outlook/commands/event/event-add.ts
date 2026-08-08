import { Event } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { calendar } from '../../../../utils/calendar.js';
import fs from 'fs';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

const bodyContentTypes = ['Text', 'HTML'] as const;
const importances = ['low', 'normal', 'high'] as const;
const onlineMeetingProviders = ['teamsForBusiness', 'skypeForBusinnes', 'skypeForConsumer'] as const;
const sensitivities = ['normal', 'personal', 'private', 'confidential'] as const;
const showAsStatuses = ['free', 'tentative', 'busy', 'oof', 'workingElsewhere'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  subject: z.string(),
  start: z.string().refine(date => validation.isValidGraphDateTime(date), {
    error: e => `'${e.input}' is not a valid date.`
  }),
  end: z.string().refine(date => validation.isValidGraphDateTime(date), {
    error: e => `'${e.input}' is not a valid date.`
  }),
  userId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID.`
  }).optional(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid UPN.`
  }).optional(),
  calendarId: z.string().optional(),
  calendarName: z.string().optional(),
  allowNewTimeProposals: z.boolean().optional().default(true),
  bodyContents: z.string().optional(),
  bodyContentType: z.preprocess(val => {
    const target = String(val).toLowerCase();
    return bodyContentTypes.find(t => t.toLowerCase() === target) ?? val;
  }, z.enum(bodyContentTypes)).optional(),
  categories: z.string().transform((value) => value.split(',').map(String)).optional(),
  hideAttendees: z.boolean().optional().default(false),
  importance: z.preprocess(val => {
    const target = String(val).toLowerCase();
    return importances.find(t => t.toLowerCase() === target) ?? val;
  }, z.enum(importances)).optional(),
  isAllDay: z.boolean().optional().default(false),
  isOnlineMeeting: z.boolean().optional().default(false),
  isReminderOn: z.boolean().optional().default(true),
  location: z.string().optional(),
  locationEmailAddress: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid email address.`
  }).optional(),
  locations: z.string().transform((value) => value.split(',').map(String)).optional(),
  onlineMeetingProvider: z.preprocess(val => {
    const target = String(val).toLowerCase();
    return onlineMeetingProviders.find(t => t.toLowerCase() === target) ?? val;
  }, z.enum(onlineMeetingProviders)).optional(),
  optionalAttendees: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `The following attendees names are invalid for the option 'optionalAttendees': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    }).transform((value) => value.split(',').map(String)).optional(),
  recurrence: z.string().optional(),
  reminderMinutesBeforeStart: z.number().refine(minutes => minutes >= 0, {
    error: () => 'The number of reminder minutes must be a positive number or 0'
  }).optional(),
  requiredAttendees: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `The following attendees names are invalid for the option 'requiredAttendees': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    }).transform((value) => value.split(',').map(String)).optional(),
  resources: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `The following attendees names are invalid for the option 'resources': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    }).transform((value) => value.split(',').map(String)).optional(),
  responseRequested: z.boolean().optional().default(true),
  sensitivity: z.preprocess(val => {
    const target = String(val).toLowerCase();
    return sensitivities.find(t => t.toLowerCase() === target) ?? val;
  }, z.enum(sensitivities)).optional(),
  showAs: z.preprocess(val => {
    const target = String(val).toLowerCase();
    return showAsStatuses.find(t => t.toLowerCase() === target) ?? val;
  }, z.enum(showAsStatuses)).optional(),
  timeZone: z.string().optional().default('UTC'),
  transactionId: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookEventAddCommand extends GraphCommand {
  public get name(): string {
    return commands.EVENT_ADD;
  }

  public get description(): string {
    return `Create an event in the default calendar or a specific calendar of a user`;
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !(options.calendarId && options.calendarName), {
        error: 'Specify either calendarId or calendarName, but not both.'
      })
      .refine(options => !(options.location && options.locations), {
        error: 'Specify either location or locations, but not both.'
      })
      .refine(options => !(options.userId && options.userName), {
        error: 'Specify either userId or userName, but not both.'
      })
      .refine(options => !(options.isAllDay && (options.start.endsWith('T00:00:00') || options.end.endsWith('T00:00:00'))), {
        error: 'When isAllDay is true, start and end must be set to midnight.'
      })
      .refine(options => !(options.reminderMinutesBeforeStart && !options.isReminderOn), {
        error: 'When reminderMinutesBeforeStart is specified, isReminderOn must be true.'
      })
      .refine(options => !(options.locationEmailAddress && !options.location), {
        error: 'When locationEmailAddress is specified, location must be also specified.'
      })
      .refine(options => new Date(options.start).getTime() < new Date(options.end).getTime(), {
        error: 'Start date must be before end date.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
    let principalUrl = '';

    const token = auth.connection.accessTokens[auth.defaultResource].accessToken;

    if (isAppOnlyAccessToken) {
      if (!args.options.userId && !args.options.userName) {
        throw `The option 'userId' or 'userName' is required when creating an event using application permissions.`;
      }
    }
    else {
      if (args.options.userId) {
        const currentUserId = accessToken.getUserIdFromAccessToken(token);
        if (args.options.userId !== currentUserId) {
          throw `You can only create your own events when using delegated permissions. The specified userId '${args.options.userId}' does not match the current user '${currentUserId}'.`;
        }
      }

      if (args.options.userName) {
        const currentUserName = accessToken.getUserNameFromAccessToken(token);
        if (args.options.userName.toLowerCase() !== currentUserName.toLowerCase()) {
          throw `You can only create your own events when using delegated permissions. The specified userName '${args.options.userName}' does not match the current user '${currentUserName}'.`;
        }
      }
    }

    let userIdentifier: string | undefined;
    if (args.options.userId || args.options.userName) {
      userIdentifier = args.options.userId ?? args.options.userName;
      principalUrl += `users('${userIdentifier}')`;
    }
    else {
      userIdentifier = accessToken.getUserNameFromAccessToken(token);
      principalUrl += 'me';
    }

    if (this.verbose) {
      await logger.logToStderr(`Creating event for user ${userIdentifier}...`);
    }

    let calendarId = args.options.calendarId;
    if (args.options.calendarName) {
      calendarId = (await calendar.getUserCalendarByName(userIdentifier!, args.options.calendarName))!.id;
    }

    let requestUrl: string = `${this.resource}/v1.0/${principalUrl}`;

    if (calendarId) {
      requestUrl += `/calendars/${calendarId}/events`;
    }
    else {
      requestUrl += '/events';
    }

    const body : any = {};
    body['subject'] = args.options.subject;
    body['start'] = {
      dateTime: args.options.start
    };
    body['end'] = {
      dateTime: args.options.end
    };

    if (args.options.bodyContentType || args.options.bodyContents) {
      body['body'] = {};

      if (args.options.bodyContentType) {
        body['body']['contentType'] = args.options.bodyContentType;
      }

      if (args.options.bodyContents) {
        if (args.options.bodyContents.startsWith('@')) {
          const fileBodyContents: string = fs.readFileSync(args.options.bodyContents.replace('@', ''), 'utf8');
          if (fileBodyContents) {
            body['body']['content'] = fileBodyContents;
          }
        }
        else {
          body['body']['content'] = args.options.bodyContents;
        }
      }
    }

    if (!args.options.allowNewTimeProposals) {
      body['allowNewTimeProposals'] = args.options.allowNewTimeProposals;
    }

    if (args.options.categories) {
      body['categories'] = args.options.categories;
    }

    if (args.options.hideAttendees) {
      body['hideAttendees'] = args.options.hideAttendees;
    }

    if (args.options.importance) {
      body['importance'] = args.options.importance;
    }

    if (args.options.isAllDay) {
      body['isAllDay'] = args.options.isAllDay;
    }

    if (args.options.isOnlineMeeting) {
      body['isOnlineMeeting'] = args.options.isOnlineMeeting;
    }

    if (!args.options.isReminderOn) {
      body['isReminderOn'] = false;
    }

    if (args.options.location || args.options.locationEmailAddress) {
      body['location'] = {};

      if (args.options.location) {
        body['location']['displayName'] = args.options.location;
      }

      if (args.options.locationEmailAddress) {
        body['location']['locationEmailAddress'] = args.options.locationEmailAddress;
      }
    }

    if (args.options.locations) {
      const locations: Array<any> = [];

      args.options.locations.forEach(displayName => locations.push({
        displayName: displayName
      }));

      body['locations'] = locations;
    }

    if (args.options.onlineMeetingProvider) {
      body['onlineMeetingProvider'] = args.options.onlineMeetingProvider;
    }

    if (args.options.optionalAttendees || args.options.requiredAttendees || args.options.resources) {
      body['attendees'] = [];
      const attendees = body['attendees'] as Array<any>;

      if (args.options.optionalAttendees) {
        args.options.optionalAttendees.forEach(value =>
          attendees.push({
            emailAddress: value,
            type: 'optional'
          }));
      }

      if (args.options.requiredAttendees) {
        args.options.requiredAttendees.forEach(value =>
          attendees.push({
            emailAddress: value,
            type: 'required'
          }));
      }

      if (args.options.resources) {
        args.options.resources.forEach(value =>
          attendees.push({
            emailAddress: value,
            type: 'resource'
          }));
      }
    }

    if (args.options.recurrence) {
      if (args.options.recurrence.startsWith('@')) {
        const fileRecurrence: string = fs.readFileSync(args.options.recurrence.replace('@', ''), 'utf8');
        if (fileRecurrence) {
          body['recurrence'] = JSON.parse(fileRecurrence);
        }
      }
      else {
        body['recurrence'] = JSON.parse(args.options.recurrence);
      }
    }

    if (args.options.isReminderOn && args.options.reminderMinutesBeforeStart) {
      body['reminderMinutesBeforeStart'] = args.options.reminderMinutesBeforeStart;
    }

    if (!args.options.responseRequested) {
      body['responseRequested'] = false;
    }

    if (args.options.sensitivity) {
      body['sensitivity'] = args.options.sensitivity;
    }

    if (args.options.showAs) {
      body['showAs'] = args.options.showAs;
    }

    if (args.options.timeZone) {
      body['start']['timeZone'] = args.options.timeZone;
      body['end']['timeZone'] = args.options.timeZone;
    }

    if (args.options.transactionId) {
      body['transactionId'] = args.options.transactionId;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: body
    };

    try {
      const result = await request.post<Event>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookEventAddCommand();
