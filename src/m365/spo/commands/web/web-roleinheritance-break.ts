import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import request from '../../../../request';
import { validation } from '../../../../utils';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
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
        .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };
    if (args.options.confirm) {
      breakroleInheritance();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to break the role inheritance of subsite ${args.options.webUrl}`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          breakroleInheritance();
        }
      });
    }
  }
}

module.exports = new SpoWebRoleInheritanceBreakCommand();
