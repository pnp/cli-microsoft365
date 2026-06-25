import { z } from 'zod';
import { AuthType } from "../../../../Auth.js";
import { cli } from "../../../../cli/cli.js";
import { Logger } from "../../../../cli/Logger.js";
import { globalOptionsZod } from "../../../../Command.js";
import { settingsNames } from "../../../../settingsNames.js";
import { validation } from "../../../../utils/validation.js";
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import commands from "../../commands.js";

const settingNameValues = Object.getOwnPropertyNames(settingsNames) as [string, ...string[]];

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  key: z.enum(settingNameValues).alias('k'),
  value: z.string().alias('v')
});
type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class CliConfigSetCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CONFIG_SET;
  }

  public get description(): string {
    return 'Sets CLI for Microsoft 365 configuration options';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => {
        if (opts.key === settingsNames.output) {
          return ['text', 'json', 'csv', 'md', 'none'].includes(opts.value);
        }
        return true;
      }, {
        error: `The value is not valid for the option ${settingsNames.output}. Allowed values: text, json, csv, md, none`,
        path: ['value']
      })
      .refine(opts => {
        if (opts.key === settingsNames.errorOutput) {
          return ['stdout', 'stderr'].includes(opts.value);
        }
        return true;
      }, {
        error: `The value is not valid for the option ${settingsNames.errorOutput}. Allowed values: stdout, stderr`,
        path: ['value']
      })
      .refine(opts => {
        if (opts.key === settingsNames.promptListPageSize) {
          const num = Number(opts.value);
          return !isNaN(num) && num > 0;
        }
        return true;
      }, {
        error: `The value is not valid for the option ${settingsNames.promptListPageSize}. The value has to be a number higher than 0.`,
        path: ['value']
      })
      .refine(opts => {
        if (opts.key === settingsNames.helpMode) {
          return cli.helpModes.includes(opts.value);
        }
        return true;
      }, {
        error: `The value is not valid for the option ${settingsNames.helpMode}. Allowed values: ${cli.helpModes.join(', ')}`,
        path: ['value']
      })
      .refine(opts => {
        if (opts.key === settingsNames.authType) {
          return Object.values(AuthType).map(String).includes(opts.value);
        }
        return true;
      }, {
        error: `The value is not valid for the option ${settingsNames.authType}. Allowed values: ${Object.values(AuthType).join(', ')}`,
        path: ['value']
      })
      .refine(opts => {
        if (opts.key === settingsNames.helpTarget) {
          return cli.helpTargets.includes(opts.value);
        }
        return true;
      }, {
        error: `The value is not valid for the option ${settingsNames.helpTarget}. Allowed values: ${cli.helpTargets.join(', ')}`,
        path: ['value']
      })
      .refine(opts => {
        if (opts.key === settingsNames.clientId) {
          return validation.isValidGuid(opts.value);
        }
        return true;
      }, {
        error: `The value is not valid for the option ${settingsNames.clientId}. The value has to be a valid GUID.`,
        path: ['value']
      })
      .refine(opts => {
        if (opts.key === settingsNames.tenantId) {
          return opts.value === 'common' || validation.isValidGuid(opts.value);
        }
        return true;
      }, {
        error: `The value is not valid for the option ${settingsNames.tenantId}. The value has to be a valid GUID or 'common'.`,
        path: ['value']
      }) as any;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let value: any;

    switch (args.options.key) {
      case settingsNames.autoOpenLinksInBrowser:
      case settingsNames.copyDeviceCodeToClipboard:
      case settingsNames.csvHeader:
      case settingsNames.csvQuoted:
      case settingsNames.csvQuotedEmpty:
      case settingsNames.disableTelemetry:
      case settingsNames.printErrorsAsPlainText:
      case settingsNames.prompt:
      case settingsNames.showHelpOnFailure:
        value = args.options.value === 'true';
        break;
      default:
        value = args.options.value;
        break;
    }

    cli.getConfig().set(args.options.key, value);
  }
}

export default new CliConfigSetCommand();