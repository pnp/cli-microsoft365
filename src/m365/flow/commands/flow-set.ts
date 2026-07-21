import { z } from 'zod';
import { cli } from '../../../cli/cli.js';
import { CommandError, globalOptionsZod } from '../../../Command.js';
import { Logger } from '../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';
import commands from '../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.uuid().alias('n'),
  environmentName: z.string().alias('e'),
  definition: z.string(),
  force: z.boolean().optional().alias('f')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface FlowCheckResult {
  errorDescription?: string;
  operationName?: string;
  ruleId?: string;
}

class FlowSetCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.SET;
  }

  public get description(): string {
    return 'Sets the specified Power Automate flow';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => this.isJsonObject(formatting.tryParseJson(opts.definition)), {
        error: 'The specified definition is not a valid JSON object',
        path: ['definition']
      })
      .refine(opts => {
        const definition = formatting.tryParseJson(opts.definition);
        return !this.isJsonObject(definition) || this.isJsonObject(definition.properties?.definition);
      }, {
        error: 'The specified definition does not contain the properties.definition object',
        path: ['definition']
      });
  }

  private isJsonObject(value: any): boolean {
    return typeof value === 'object' && value !== null && !Array.isArray(value);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating Power Automate flow ${args.options.name}...`);
    }

    const baseUrl = `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}`;
    const definition = this.sanitizeDefinition(JSON.parse(args.options.definition));

    const updateFlow = async (): Promise<void> => {
      const requestOptions: CliRequestOptions = {
        url: `${baseUrl}?api-version=2016-11-01`,
        headers: {
          accept: 'application/json',
          'content-type': 'application/json'
        },
        data: definition,
        responseType: 'json'
      };

      try {
        await request.patch(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    const errorsRequestOptions: CliRequestOptions = {
      url: `${baseUrl}/checkFlowErrors?api-version=2016-11-01`,
      headers: {
        accept: 'application/json',
        'content-type': 'application/json'
      },
      data: definition,
      responseType: 'json'
    };

    const warningsRequestOptions: CliRequestOptions = {
      url: `${baseUrl}/checkFlowWarnings?api-version=2016-11-01`,
      headers: {
        accept: 'application/json',
        'content-type': 'application/json'
      },
      data: definition,
      responseType: 'json'
    };

    let errors: FlowCheckResult[];
    let warnings: FlowCheckResult[];
    try {
      [errors, warnings] = await Promise.all([
        request.post<FlowCheckResult[]>(errorsRequestOptions),
        request.post<FlowCheckResult[]>(warningsRequestOptions)
      ]);
    }
    catch (err: any) {
      return this.handleRejectedODataJsonPromise(err);
    }

    if (errors.length > 0) {
      const errorDetails = args.options.output === 'json'
        ? JSON.stringify(errors, null, 2)
        : errors.map(e => `  - ${e.errorDescription ?? JSON.stringify(e)}`).join('\n');
      throw new CommandError(`The flow definition has the following errors:\n${errorDetails}`);
    }

    if (warnings.length > 0) {
      const warningDetails = args.options.output === 'json'
        ? JSON.stringify(warnings, null, 2)
        : warnings.map(w => `  - ${w.errorDescription ?? JSON.stringify(w)}`).join('\n');

      if (args.options.force) {
        await logger.logToStderr(`The flow definition has the following warnings:\n\n${warningDetails}`);
        await updateFlow();
        return;
      }

      const confirmed = await cli.promptForConfirmation({
        message: `The flow definition has the following warnings:\n\n${warningDetails}\n\nDo you want to proceed with the update?`
      });

      if (confirmed) {
        await updateFlow();
      }
    }
    else {
      await updateFlow();
    }
  }

  private sanitizeDefinition(definition: any): any {
    delete definition.displayName;
    delete definition.description;
    delete definition.triggers;
    delete definition.actions;

    if (definition.properties?.connectionReferences) {
      for (const key of Object.keys(definition.properties.connectionReferences)) {
        delete definition.properties.connectionReferences[key].operationDefinitions;
        delete definition.properties.connectionReferences[key].apiDefinition;
      }
    }

    return definition;
  }
}

export default new FlowSetCommand();
