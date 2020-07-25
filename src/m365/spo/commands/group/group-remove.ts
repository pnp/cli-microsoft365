import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeGroup: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Removing group in web at ${args.options.webUrl}...`);
      }

      let groupId: number | undefined;

      ((): Promise<any> => {
        if (args.options.name) {
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/sitegroups/GetByName('${args.options.name}')?$select=Id`,
            headers: {
              accept: 'application/json'
            },
            json: true
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
          json: true
        };

        return request.post(requestOptions)
      }).then((): void => {
        // REST post call doesn't return anything
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeGroup();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the group ${args.options.id || args.options.name} from web ${args.options.webUrl}?`,
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
        option: '-u, --webUrl <webUrl>',
        description: 'Url of the web to remove the group from'
      },
      {
        option: '--id [id]',
        description: 'ID of the group to remove. Use ID or name but not both'
      },
      {
        option: '--name [name]',
        description: 'Name of the group to remove. Use ID or name but not both'
      },
      {
        option: '--confirm',
        description: 'Confirm removal of the group'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
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
    };
  }
}

module.exports = new SpoGroupRemoveCommand();