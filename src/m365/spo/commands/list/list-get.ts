import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListInstance } from "./ListInstance";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  title?: string;
  properties?: string;
  withPermissions?: boolean;
}

class SpoListGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_GET;
  }

  public get description(): string {
    return 'Gets information about the specific list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.title = (!(!args.options.title)).toString();
    telemetryProps.properties = (!(!args.options.properties)).toString();
    telemetryProps.withPermissions = typeof args.options.withPermissions !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information for list in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.title as string)}')`;
    }

    let propertiesSelect: string = args.options.properties ? `?$select=${encodeURIComponent(args.options.properties)}` : ``;
    propertiesSelect += args.options.withPermissions ? `${args.options.properties ? '&' : '?'}$expand=HasUniqueRoleAssignments,RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings` : ``;

    const requestOptions: any = {
      url: requestUrl + propertiesSelect,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<ListInstance>(requestOptions)
      .then((listInstance: ListInstance): void => {
        logger.log(listInstance);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-p, --properties [properties]'
      },
      {
        option: '--withPermissions'
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

    if (args.options.id) {
      if (!validation.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }
    }

    if (args.options.id && args.options.title) {
      return 'Specify id or title, but not both';
    }

    if (!args.options.id && !args.options.title) {
      return 'Specify id or title, one is required';
    }

    return true;
  }
}

module.exports = new SpoListGetCommand();