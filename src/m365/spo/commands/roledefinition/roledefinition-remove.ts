import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const id: number = parseInt(args.options.id);
    if (isNaN(id)) {
      return `${args.options.id} is not a valid role definition ID`;
    }

    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoRoleDefinitionRemoveCommand();