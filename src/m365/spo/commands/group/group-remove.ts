import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
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
  id?: number;
  name?: string;
  confirm?: boolean;
}

class SpoGroupRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_REMOVE;
  }

  public get description(): string {
    return 'Removes group from specific web';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.name = (!(!args.options.name)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeGroup: () => void = (): void => {
      if (this.verbose) {
        logger.logToStderr(`Removing group in web at ${args.options.webUrl}...`);
      }

      let groupId: number | undefined;

      ((): Promise<any> => {
        if (args.options.name) {
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/sitegroups/GetByName('${args.options.name}')?$select=Id`,
            headers: {
              accept: 'application/json'
            },
            responseType: 'json'
          };
          return request.get(requestOptions);
        }

        groupId = args.options.id;
        return Promise.resolve(undefined as any);
      })().then((res?: { Id: number }) => {
        if (res && res.Id) {
          groupId = res.Id;
        }

        const requestUrl = `${args.options.webUrl}/_api/web/sitegroups/RemoveById(${groupId})`;
        const requestOptions: any = {
          url: requestUrl,
          method: 'POST',
          headers: {
            'content-length': 0,
            'accept': 'application/json'
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      }).then((): void => {
        // REST post call doesn't return anything
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeGroup();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the group ${args.options.id || args.options.name} from web ${args.options.webUrl}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeGroup();
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
        option: '--id [id]'
      },
      {
        option: '--name [name]'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.id && args.options.name) {
      return 'Specify id or name, but not both';
    }

    if (!args.options.id && !args.options.name) {
      return 'Specify id or name';
    }

    if (args.options.id && typeof args.options.id !== 'number') {
      return `${args.options.id} is not a number`;
    }

    return true;
  }
}

module.exports = new SpoGroupRemoveCommand();