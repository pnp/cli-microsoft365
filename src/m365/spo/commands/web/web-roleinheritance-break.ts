import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  confirm?: boolean;
  clearExistingPermissions?: boolean;
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
        clearExistingPermissions: args.options.clearExistingPermissions === true,
        confirm: (!(!args.options.confirm)).toString()
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
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }
        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Break role inheritance of subsite at ${args.options.webUrl}...`);
    }

    let keepExistingPermissions: boolean = true;
    if (args.options.clearExistingPermissions) {
      keepExistingPermissions = !args.options.clearExistingPermissions;
    }

    const breakroleInheritance = (): void => {
      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/breakroleinheritance(${keepExistingPermissions})`,
        method: 'POST',
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      request
        .post(requestOptions)
        .then(_ => (err: any): void => this.handleRejectedODataJsonPromise(err));
    };
    if (args.options.confirm) {
      breakroleInheritance();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to break the role inheritance of subsite ${args.options.webUrl}`
      });

      if (result.continue) {
        breakroleInheritance();
      }
    }
  }
}

module.exports = new SpoWebRoleInheritanceBreakCommand();
