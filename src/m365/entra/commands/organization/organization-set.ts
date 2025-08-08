import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Organization } from '@microsoft/microsoft-graph-types';
import { odata } from '../../../../utils/odata.js';

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string().refine(id => validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    })).optional()),
    displayName: zod.alias('d', z.string().optional()),
    marketingNotificationEmails: z.string().refine(emails => validation.isValidUserPrincipalNameArray(emails) === true, invalidEmails => ({
      message: `The following marketing notification emails are invalid: ${invalidEmails}.`
    })).transform((value) => value.split(',')).optional(),
    securityComplianceNotificationMails: z.string().refine(emails => validation.isValidUserPrincipalNameArray(emails) === true, invalidEmails => ({
      message: `The following security compliance notification emails are invalid: ${invalidEmails}.`
    })).transform((value) => value.split(',')).optional(),
    securityComplianceNotificationPhones: z.string().transform((value) => value.split(',')).optional(),
    technicalNotificationMails: z.string().refine(emails => validation.isValidUserPrincipalNameArray(emails) === true, invalidEmails => ({
      message: `The following technical notification emails are invalid: ${invalidEmails}.`
    })).transform((value) => value.split(',')).optional(),
    contactEmail: z.string().refine(id => validation.isValidUserPrincipalName(id), id => ({
      message: `'${id}' is not a valid email.`
    })).optional(),
    statementUrl: z.string().optional()
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraOrganizationSetCommand extends GraphCommand {
  public get name(): string {
    return commands.ORGANIZATION_SET;
  }

  public get description(): string {
    return 'Updates info about the organization';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !(options.id && options.displayName), {
        message: 'Specify either id or displayName, but not both'
      })
      .refine(options => options.id || options.displayName, {
        message: 'Specify either id or displayName'
      })
      .refine(options => [options.contactEmail, options.marketingNotificationEmails, options.securityComplianceNotificationMails, options.securityComplianceNotificationPhones,
        options.statementUrl, options.technicalNotificationMails].filter(o => o !== undefined).length > 0, {
        message: 'Specify at least one of the following options: contactEmail, marketingNotificationEmails, securityComplianceNotificationMails, securityComplianceNotificationPhones, statementUrl, or technicalNotificationMails'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let organizationId = args.options.id;

      if (args.options.displayName) {
        organizationId = await this.getOrganizationIdByDisplayName(args.options.displayName);
      }

      if (args.options.verbose) {
        await logger.logToStderr(`Updating organization with ID ${organizationId}...`);
      }

      const data: Organization = {
        marketingNotificationEmails: args.options.marketingNotificationEmails,
        securityComplianceNotificationMails: args.options.securityComplianceNotificationMails,
        securityComplianceNotificationPhones: args.options.securityComplianceNotificationPhones,
        technicalNotificationMails: args.options.technicalNotificationMails
      };

      if (args.options.contactEmail || args.options.statementUrl) {
        data.privacyProfile = {};
      }

      if (args.options.contactEmail) {
        data.privacyProfile!.contactEmail = args.options.contactEmail;
      }

      if (args.options.statementUrl) {
        data.privacyProfile!.statementUrl = args.options.statementUrl;
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/organization/${organizationId}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        data: data,
        responseType: 'json'
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  async getOrganizationIdByDisplayName(displayName: string): Promise<string> {
    const url = `${this.resource}/v1.0/organization?$select=id,displayName`;

    // the endpoint always returns one item
    const organizations = await odata.getAllItems<Organization>(url);

    if (organizations[0].displayName !== displayName) {
      throw `The specified organization '${displayName}' does not exist.`;
    }

    return organizations[0].id!;
  }
}

export default new EntraOrganizationSetCommand();