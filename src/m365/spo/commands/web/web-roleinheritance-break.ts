import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  clearExistingPermissions?: boolean;
  force?: boolean;
}

class SpoWebRoleInheritanceBreakCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_ROLEINHERITANCE_BREAK;
  }

  public get description(): string {
    return 'Break role inheritance of subsite';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        clearExistingPermissions: !!args.options.clearExistingPermissions,
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-c, --clearExistingPermissions'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Break role inheritance of subsite with URL ${args.options.webUrl}...`);
    }

    if (args.options.force) {
      await this.breakRoleInheritance(args.options);
    }
    else {
      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to break the role inheritance of subsite ${args.options.webUrl}?` });

      if (result) {
        await this.breakRoleInheritance(args.options);
      }
    }
  }

  private async breakRoleInheritance(options: Options): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/web/breakroleinheritance(${!options.clearExistingPermissions})`,
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoWebRoleInheritanceBreakCommand();
