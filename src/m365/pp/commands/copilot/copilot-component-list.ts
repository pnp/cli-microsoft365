import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { CommandError, globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import ppCopilotGetCommand from './copilot-get.js';

const componentTypes = {
  topic: 0,
  skill: 1,
  variable: 2,
  entity: 3,
  dialog: 4,
  trigger: 5,
  languageUnderstanding: 6,
  languageGeneration: 7,
  dialogSchema: 8,
  topicV2: 9,
  translationsV2: 10,
  entityV2: 11,
  variableV2: 12,
  skillV2: 13,
  fileAttachment: 14,
  customGpt: 15,
  knowledgeSource: 16,
  externalTrigger: 17,
  settings: 18,
  testCase: 19,
  customMetricDefinition: 20
} as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  copilotId: z.string().refine(val => validation.isValidGuid(val), {
    error: 'The value must be a valid GUID.'
  }).optional().alias('i'),
  copilotName: z.string().optional().alias('n'),
  type: z.enum(Object.keys(componentTypes) as [keyof typeof componentTypes, ...(keyof typeof componentTypes)[]]).optional().alias('t'),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpCopilotComponentListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.COPILOT_COMPONENT_LIST;
  }

  public get description(): string {
    return 'Lists the components of the specified copilot';
  }

  public defaultProperties(): string[] | undefined {
    return ['botcomponentid', 'name', 'componenttype', 'schemaname', 'ismanaged', 'statecode'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.copilotId, opts.copilotName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'copilotId' or 'copilotName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['copilotId', 'copilotName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving components of copilot '${args.options.copilotId || args.options.copilotName}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
      const copilotId = await this.getCopilotId(args);

      let filter = `_parentbotid_value eq ${copilotId}`;
      if (args.options.type) {
        filter += ` and componenttype eq ${componentTypes[args.options.type]}`;
      }

      const select = 'botcomponentid,name,componenttype,schemaname,ismanaged,statecode';
      const components = await odata.getAllItems<any>(`${dynamicsApiUrl}/api/data/v9.1/botcomponents?$select=${select}&$filter=${filter}`);

      await logger.log(components);
    }
    catch (err: any) {
      if (err instanceof CommandError) {
        throw err;
      }

      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getCopilotId(args: CommandArgs): Promise<string> {
    if (args.options.copilotId) {
      return args.options.copilotId;
    }

    const options = {
      environmentName: args.options.environmentName,
      name: args.options.copilotName,
      asAdmin: args.options.asAdmin,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await cli.executeCommandWithOutput(ppCopilotGetCommand as Command, { options: { ...options, _: [] } });
    const copilot = JSON.parse(output.stdout);
    return copilot.botid;
  }
}

export default new PpCopilotComponentListCommand();
