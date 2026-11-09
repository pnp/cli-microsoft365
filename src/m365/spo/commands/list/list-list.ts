import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListInstance } from './ListInstance.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  properties: z.string().optional().alias('p'),
  filter: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface FieldProperties {
  selectProperties: string[];
  expandProperties: string[];
}

class SpoListListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_LIST;
  }

  public get description(): string {
    return 'Lists all available list in the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Id'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving all lists in site at ${args.options.webUrl}...`);
    }

    try {
      const fieldProperties = this.formatSelectProperties(args.options.properties);
      const queryParams = [`$expand=${fieldProperties.expandProperties.join(',')}`, `$select=${fieldProperties.selectProperties.join(',')}`];

      if (args.options.filter) {
        queryParams.push(`$filter=${args.options.filter}`);
      }

      const listInstances = await odata.getAllItems<ListInstance>(`${args.options.webUrl}/_api/web/lists?${queryParams.join('&')}`);
      await logger.log(listInstances);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private formatSelectProperties(fields: string | undefined): FieldProperties {
    const selectProperties: any[] = ['RootFolder/ServerRelativeUrl'];
    const expandProperties: any[] = ['RootFolder'];

    if (!fields) {
      selectProperties.push('*');
    }

    if (fields) {
      fields.split(',').forEach((field) => {
        const subparts = field.trim().split('/');
        if (subparts.length > 1) {
          expandProperties.push(subparts[0]);
        }
        selectProperties.push(field.trim());
      });
    }

    return {
      selectProperties: [...new Set(selectProperties)],
      expandProperties: [...new Set(expandProperties)]
    };
  }
}

export default new SpoListListCommand();