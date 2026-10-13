import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { AssociatedGroupPropertiesCollection } from './AssociatedGroupPropertiesCollection.js';
import { GroupProperties } from './GroupProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  associatedGroupsOnly: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoGroupListCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_LIST;
  }

  public get description(): string {
    return 'Lists all the groups within specific web';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Title', 'LoginName', 'IsHiddenInUI', 'PrincipalType', 'Type'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of groups for specified web at ${args.options.webUrl}...`);
    }

    const baseUrl = `${args.options.webUrl}/_api/web`;

    try {
      if (!args.options.associatedGroupsOnly) {
        await this.getSiteGroups(baseUrl, logger);
      }
      else {
        await this.getAssociatedGroups(baseUrl, args.options, logger);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSiteGroups(baseUrl: string, logger: Logger): Promise<void> {
    const groupProperties = await odata.getAllItems<GroupProperties>(`${baseUrl}/sitegroups`);
    await logger.log(groupProperties);
  }

  private async getAssociatedGroups(baseUrl: string, options: Options, logger: Logger): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: baseUrl + '?$expand=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup&$select=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const groupProperties = await request.get<AssociatedGroupPropertiesCollection>(requestOptions);

    if (!options.output || !cli.shouldTrimOutput(options.output)) {
      await logger.log(groupProperties);
    }
    else {
      //converted to text friendly output
      const output = Object.getOwnPropertyNames(groupProperties).map(prop => ({ Type: prop, ...(groupProperties as any)[prop] }));
      await logger.log(output);
    }
  }
}

export default new SpoGroupListCommand();