import { AuthType } from "../../../../Auth.js";
import { cli } from "../../../../cli/cli.js";
import { Logger } from "../../../../cli/Logger.js";
import GlobalOptions from "../../../../GlobalOptions.js";
import { settingsNames } from "../../../../settingsNames.js";
import { validation } from "../../../../utils/validation.js";
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import commands from "../../commands.js";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  key: string;
  value: string;
}

class CliConfigSetCommand extends AnonymousCommand {
  private static readonly optionNames: string[] = Object.getOwnPropertyNames(settingsNames);

  public get name(): string {
    return commands.CONFIG_SET;
  }

  public get description(): string {
    return 'Manage global configuration settings about the CLI for Microsoft 365';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      this.telemetryProperties[args.options.key] = args.options.value;
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-k, --key <key>',
        autocomplete: CliConfigSetCommand.optionNames
      },
      {
        option: '-v, --value <value>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (CliConfigSetCommand.optionNames.indexOf(args.options.key) < 0) {
          return `${args.options.key} is not a valid setting. Allowed values: ${CliConfigSetCommand.optionNames.join(', ')}`;
        }

        const allowedOutputs = ['text', 'json', 'csv', 'md', 'none'];
        if (args.options.key === settingsNames.output &&
          allowedOutputs.indexOf(args.options.value) === -1) {
          return `${args.options.value} is not a valid value for the option ${args.options.key}. Allowed values: ${allowedOutputs.join(', ')}`;
        }

        const allowedErrorOutputs = ['stdout', 'stderr'];
        if (args.options.key === settingsNames.errorOutput &&
          allowedErrorOutputs.indexOf(args.options.value) === -1) {
          return `${args.options.value} is not a valid value for the option ${args.options.key}. Allowed values: ${allowedErrorOutputs.join(', ')}`;
        }

        if (args.options.key === settingsNames.promptListPageSize &&
          typeof args.options.value !== 'number') {
          return `${args.options.value} is not a valid value for the option ${args.options.key}. The value has to be a number.`;
        }

        if (args.options.key === settingsNames.promptListPageSize &&
          (args.options.value as unknown as number) <= 0) {
          return `${args.options.value} is not a valid value for the option ${args.options.key}. The number has to be higher than 0.`;
        }

        if (args.options.key === settingsNames.helpMode &&
          cli.helpModes.indexOf(args.options.value) === -1) {
          return `${args.options.value} is not a valid value for the option ${args.options.key}. Allowed values: ${cli.helpModes.join(', ')}`;
        }

        if (args.options.key === settingsNames.authType &&
          !Object.values(AuthType).map(String).includes(args.options.value)) {
          return `${args.options.value} is not a valid value for the option ${args.options.key}. Allowed values: ${Object.values(AuthType).join(', ')}`;
        }

        if (args.options.key === settingsNames.helpTarget &&
          !cli.helpTargets.includes(args.options.value)) {
          return `${args.options.value} is not a valid value for the option ${args.options.key}. Allowed values: ${cli.helpTargets.join(', ')}`;
        }

        if (args.options.key === settingsNames.clientId &&
          !validation.isValidGuid(args.options.value)) {
          return `${args.options.value} is not a valid value for the option ${args.options.key}. The value has to be a valid GUID.`;
        }

        if (args.options.key === settingsNames.tenantId &&
          !(args.options.value === 'common' || validation.isValidGuid(args.options.value))) {
          return `${args.options.value} is not a valid value for the option ${args.options.key}. The value has to be a valid GUID or 'common'.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let value: any = undefined;

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
      default:
        value = args.options.value;
        break;
    }

    cli.getConfig().set(args.options.key, value);
  }
}

export default new CliConfigSetCommand();