import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  force?: boolean;
}

class SpoWebRoleInheritanceResetCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_ROLEINHERITANCE_RESET;
  }

  public get description(): string {
    return 'Restores role inheritance of subsite';
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
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Restore role inheritance of subsite at ${args.options.webUrl}...`);
    }

    if (args.options.force) {
      await this.resetWebRoleInheritance(args.options);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to reset the role inheritance of ${args.options.webUrl}`
      });

      if (result.continue) {
        await this.resetWebRoleInheritance(args.options);
      }
    }
  }

  private async resetWebRoleInheritance(options: Options): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${options.webUrl}/_api/web/resetroleinheritance`,
        method: 'POST',
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoWebRoleInheritanceResetCommand();