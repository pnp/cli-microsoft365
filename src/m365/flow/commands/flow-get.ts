import { z } from 'zod';
import { Logger } from '../../../cli/Logger.js';
import { globalOptionsZod } from '../../../Command.js';
import request, { CliRequestOptions } from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';
import commands from '../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  name: z.string().alias('n'),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface Trigger {
  type: string;
  kind?: string;
}

interface Action {
  type: string;
  swaggerOperationId?: string;
}

interface Flow {
  actions?: string;
  description?: string;
  displayName?: string;
  properties: {
    displayName: string;
    definitionSummary: {
      actions: Action[];
      description: string;
      triggers: Trigger[];
    };
  },
  triggers: string;
}

class FlowGetCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Flow';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Flow ${args.options.name}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01&$expand=swagger,properties.connectionreferences.apidefinition,properties.definitionsummary.operations.apioperation,operationDefinition,plan,properties.throttleData,properties.estimatedsuspensiondata,properties.licenseData`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<Flow>(requestOptions);

      res.displayName = res.properties.displayName;
      res.description = res.properties.definitionSummary.description || '';
      res.triggers = res.properties.definitionSummary.triggers.map((t: Trigger) => (t.type + (t.kind ? "-" + t.kind : '')) as string).join(', ');
      res.actions = res.properties.definitionSummary.actions.map((a: Action) => (a.type + (a.swaggerOperationId ? "-" + a.swaggerOperationId : '')) as string).join(', ');

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowGetCommand();
