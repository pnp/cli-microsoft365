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
}

class SpoUserGetCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_GET;
  }

  public get description(): string {
    return 'Gets site group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.name = (!(!args.options.name)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving information for group in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/sitegroups/GetById('${encodeURIComponent(args.options.id)}')`;
    }
    else if (args.options.name) {
      requestUrl = `${args.options.webUrl}/_api/web/sitegroups/GetByName('${encodeURIComponent(args.options.name as string)}')`;
    }

    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((groupInstance): void => {
        cmd.log(groupInstance);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the group is located'
      },
      {
        option: '-i, --id [id]',
        description: 'Id of the site group to get. Use either "id" or "name", but not both.'
      },
      {
        option: '--name [name]',
        description: 'Name of the site group to get. Use either "id" or "name", but not both.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.id && args.options.name) {
        return 'Use either "id" or "name", but not all.';
      }

      if (!args.options.id && !args.options.name) {
        return 'Specify id or name, one is required';
      }

      if (args.options.id && isNaN(args.options.id)) {
        return `Specified id ${args.options.id} is not a number`;
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoUserGetCommand();
