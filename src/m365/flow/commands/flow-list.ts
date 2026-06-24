import { z } from 'zod';
import { Logger } from '../../../cli/Logger.js';
import { globalOptionsZod } from '../../../Command.js';
import { formatting } from '../../../utils/formatting.js';
import { odata } from '../../../utils/odata.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';
import commands from '../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  sharingStatus: z.enum(['all', 'personal', 'ownedByMe', 'sharedWithMe']).optional(),
  withSolutions: z.boolean().optional(),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface PowerAutomateFlow {
  name: string;
  id: string;
  displayName: string;
  properties: {
    displayName: string;
  }
}

class FlowListCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return 'Lists Power Automate flows in the given environment';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(options => !(options.asAdmin && options.sharingStatus), {
      error: 'The options asAdmin and sharingStatus cannot be specified together.'
    });
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Getting Power Automate flows${args.options.asAdmin ? ' as admin' : ''} in environment '${args.options.environmentName}'...`);
    }

    try {
      const {
        environmentName,
        asAdmin,
        sharingStatus,
        withSolutions
      } = args.options;

      let items: PowerAutomateFlow[] = [];

      if (sharingStatus === 'personal') {
        const url = this.getApiUrl(environmentName, asAdmin, withSolutions, 'personal');
        items = await odata.getAllItems<PowerAutomateFlow>(url);
      }
      else if (sharingStatus === 'sharedWithMe') {
        const url = this.getApiUrl(environmentName, asAdmin, withSolutions, 'team');
        items = await odata.getAllItems<PowerAutomateFlow>(url);
      }
      else if (sharingStatus === 'all') {
        let url = this.getApiUrl(environmentName, asAdmin, withSolutions, 'personal');
        items = await odata.getAllItems<PowerAutomateFlow>(url);

        url = this.getApiUrl(environmentName, asAdmin, withSolutions, 'team');
        const teamFlows = await odata.getAllItems<PowerAutomateFlow>(url);
        items = items.concat(teamFlows);
      }
      else {
        const url = this.getApiUrl(environmentName, asAdmin, withSolutions);
        items = await odata.getAllItems<PowerAutomateFlow>(url);
      }

      // Remove duplicates
      items = items.filter((flow, index, self) =>
        index === self.findIndex(f => f.id === flow.id)
      );

      if (args.options.output && args.options.output !== 'json') {
        items.forEach(flow => {
          flow.displayName = flow.properties.displayName;
        });
      }
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getApiUrl(environmentName: string, asAdmin?: boolean, includeSolutionFlows?: boolean, filter?: 'personal' | 'team'): string {
    const baseEndpoint = `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple`;
    const environmentSegment = `/environments/${formatting.encodeQueryParameter(environmentName)}`;
    const adminSegment = `/scopes/admin${environmentSegment}/v2`;
    const flowsEndpoint = '/flows?api-version=2016-11-01';
    const filterQuery = filter === 'personal' || filter === 'team' ? `&$filter=search('${filter}')` : '';
    const includeQuery = includeSolutionFlows ? '&include=includeSolutionCloudFlows' : '';

    return `${baseEndpoint}${asAdmin ? adminSegment : environmentSegment}${flowsEndpoint}${filterQuery}${includeQuery}`;
  }
}

export default new FlowListCommand();