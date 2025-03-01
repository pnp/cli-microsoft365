import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { MailboxSettings } from '@microsoft/microsoft-graph-types';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

const options = globalOptionsZod
  .extend({
    userId: zod.alias('i', z.string().refine(id => validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    })).optional()),
    userName: zod.alias('n', z.string().refine(name => validation.isValidUserPrincipalName(name), name => ({
      message: `'${name}' is not a valid UPN.`
    })).optional()),
    dateFormat: z.string().optional(),
    timeFormat: z.string().optional(),
    timeZone: z.string().optional(),
    language: z.string().optional(),
    delegateMeetingMessageDeliveryOptions: z.enum(['sendToDelegateAndInformationToPrincipal', 'sendToDelegateAndPrincipal', 'sendToDelegateOnly']).optional(),
    workingDays: z.string().transform((value) => value.split(',')).pipe(z.enum(['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']).array()).optional(),
    workingHoursStartTime: z.string().optional(),
    workingHoursEndTime: z.string().optional(),
    workingHoursTimeZone: z.string().optional(),
    autoReplyExternalAudience: z.enum(['none', 'all', 'contactsOnly']).optional(),
    autoReplyExternalMessage: z.string().optional(),
    autoReplyInternalMessage: z.string().optional(),
    autoReplyStartDateTime: z.string().optional(),
    autoReplyStartTimeZone: z.string().optional(),
    autoReplyEndDateTime: z.string().optional(),
    autoReplyEndTimeZone: z.string().optional(),
    autoReplyStatus: z.enum(['disabled', 'scheduled', 'alwaysEnabled']).optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookMailboxSettingsSetCommand extends GraphCommand {
  public get name(): string {
    return commands.MAILBOX_SETTINGS_SET;
  }

  public get description(): string {
    return 'Updates user mailbox settings';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !(options.userId && options.userName), {
        message: 'Specify either userId or userName, but not both'
      })
      .refine(options => [options.workingDays, options.workingHoursStartTime, options.workingHoursEndTime, options.workingHoursTimeZone,
        options.autoReplyStatus, options.autoReplyExternalAudience, options.autoReplyExternalMessage, options.autoReplyInternalMessage,
        options.autoReplyStartDateTime, options.autoReplyStartTimeZone, options.autoReplyEndDateTime, options.autoReplyEndTimeZone,
        options.timeFormat, options.timeZone, options.dateFormat, options.delegateMeetingMessageDeliveryOptions, options.language].filter(o => o !== undefined).length > 0, {
        message: 'Specify at least one of the following options: workingDays, workingHoursStartTime, workingHoursEndTime, workingHoursTimeZone, autoReplyStatus, autoReplyExternalAudience, autoReplyExternalMessage, autoReplyInternalMessage, autoReplyStartDateTime, autoReplyStartTimeZone, autoReplyEndDateTime, autoReplyEndTimeZone, timeFormat, timeZone, dateFormat, delegateMeetingMessageDeliveryOptions, or language'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);

    let requestUrl = `${this.resource}/v1.0/me/mailboxSettings`;

    if (isAppOnlyAccessToken) {
      if (!args.options.userId && !args.options.userName) {
        throw 'When running with application permissions either userId or userName is required';
      }

      const userIdentifier = args.options.userId ?? args.options.userName;

      if (this.verbose) {
        await logger.logToStderr(`Updating mailbox settings for user ${userIdentifier}...`);
      }

      requestUrl = `${this.resource}/v1.0/users('${userIdentifier}')/mailboxSettings`;
    }
    else {
      if (args.options.userId || args.options.userName) {
        throw 'You can update mailbox settings of other users only if CLI is authenticated in app-only mode';
      }
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: this.createRequestBody(args)
    };

    try {
      const result = await request.patch<MailboxSettings>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private createRequestBody(args: CommandArgs): any {
    const data: any = {
    };

    if (args.options.dateFormat) {
      data.dateFormat = args.options.dateFormat;
    }

    if (args.options.timeFormat) {
      data.timeFormat = args.options.timeFormat;
    }

    if (args.options.timeZone) {
      data.timeZone = args.options.timeZone;
    }

    if (args.options.delegateMeetingMessageDeliveryOptions) {
      data.delegateMeetingMessageDeliveryOptions = args.options.delegateMeetingMessageDeliveryOptions;
    }

    if (args.options.language) {
      data.language = {
        locale: args.options.language
      };
    }

    if (args.options.workingDays || args.options.workingHoursStartTime || args.options.workingHoursEndTime || args.options.workingHoursTimeZone) {
      data['workingHours'] = {};
    }

    if (args.options.workingDays) {
      data.workingHours.daysOfWeek = args.options.workingDays;
    }

    if (args.options.workingHoursStartTime) {
      data.workingHours.startTime = args.options.workingHoursStartTime;
    }

    if (args.options.workingHoursEndTime) {
      data.workingHours.endTime = args.options.workingHoursEndTime;
    }

    if (args.options.workingHoursTimeZone) {
      data.workingHours.timeZone = {
        name: args.options.workingHoursTimeZone
      };
    }

    if (args.options.autoReplyStatus || args.options.autoReplyExternalAudience || args.options.autoReplyExternalMessage || args.options.autoReplyInternalMessage || args.options.autoReplyStartDateTime || args.options.autoReplyStartTimeZone || args.options.autoReplyEndDateTime || args.options.autoReplyEndTimeZone) {
      data['automaticRepliesSetting'] = {};
    }

    if (args.options.autoReplyStatus) {
      data.automaticRepliesSetting['status'] = args.options.autoReplyStatus;
    }

    if (args.options.autoReplyExternalAudience) {
      data.automaticRepliesSetting['externalAudience'] = args.options.autoReplyExternalAudience;
    }

    if (args.options.autoReplyExternalMessage) {
      data.automaticRepliesSetting['externalReplyMessage'] = args.options.autoReplyExternalMessage;
    }

    if (args.options.autoReplyInternalMessage) {
      data.automaticRepliesSetting['internalReplyMessage'] = args.options.autoReplyInternalMessage;
    }

    if (args.options.autoReplyStartDateTime || args.options.autoReplyStartTimeZone) {
      data.automaticRepliesSetting['scheduledStartDateTime'] = {};
    }

    if (args.options.autoReplyStartDateTime) {
      data.automaticRepliesSetting['scheduledStartDateTime']['dateTime'] = args.options.autoReplyStartDateTime;
    }

    if (args.options.autoReplyStartTimeZone) {
      data.automaticRepliesSetting['scheduledStartDateTime']['timeZone'] = args.options.autoReplyStartTimeZone;
    }

    if (args.options.autoReplyEndDateTime || args.options.autoReplyEndTimeZone) {
      data.automaticRepliesSetting['scheduledEndDateTime'] = {};
    }

    if (args.options.autoReplyEndDateTime) {
      data.automaticRepliesSetting['scheduledEndDateTime']['dateTime'] = args.options.autoReplyEndDateTime;
    }

    if (args.options.autoReplyEndTimeZone) {
      data.automaticRepliesSetting['scheduledEndDateTime']['timeZone'] = args.options.autoReplyEndTimeZone;
    }

    return data;
  }
}

export default new OutlookMailboxSettingsSetCommand();