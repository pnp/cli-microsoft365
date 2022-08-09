import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id: string;
  confirm?: boolean;
}

class SpoRoleDefinitionRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_REMOVE;
  }

  public get description(): string {
    return 'Removes the role definition from the specified site';
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
        option: '-i, --id <id>'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const id: number = parseInt(args.options.id);
        if (isNaN(id)) {
          return `${args.options.id} is not a valid role definition ID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeRoleDefinition: () => void = (): void => {
      if (this.verbose) {
        logger.logToStderr(`Removing role definition from site ${args.options.webUrl}...`);
      }

      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/roledefinitions(${args.options.id})`,
        method: 'delete',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .delete(requestOptions)
        .then((): void => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeRoleDefinition();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the role definition with id ${args.options.id} from site ${args.options.webUrl}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeRoleDefinition();
        }
      });
    }
  }
}

module.exports = new SpoRoleDefinitionRemoveCommand();