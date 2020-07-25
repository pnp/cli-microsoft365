import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  userName?: string;
  properties?: string;
}

class AadUserGetCommand extends GraphCommand {
  public get name(): string {
    return `${commands.USER_GET}`;
  }

  public get description(): string {
    return 'Gets information about the specified user';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.userName = typeof args.options.userName !== 'undefined';
    telemetryProps.properties = args.options.properties;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const properties: string = args.options.properties ?
      `?$select=${args.options.properties.split(',').map(p => encodeURIComponent(p.trim())).join(',')}` :
      '';

    const requestOptions: any = {
      url: `${this.resource}/v1.0/users/${encodeURIComponent(args.options.id ? args.options.id : args.options.userName as string)}${properties}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        cmd.log(res);

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]',
        description: 'The ID of the user to retrieve information for. Specify id or userName but not both'
      },
      {
        option: '-n, --userName [userName]',
        description: 'The name of the user to retrieve information for. Specify id or userName but not both'
      },
      {
        option: '-p, --properties [properties]',
        description: 'Comma-separated list of properties to retrieve'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id && !args.options.userName) {
        return 'Specify either id or userName';
      }

      if (args.options.id && args.options.userName) {
        return 'Specify either id or userName but not both';
      }

      if (args.options.id &&
        !Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new AadUserGetCommand();
