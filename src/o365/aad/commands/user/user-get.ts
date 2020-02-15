import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
          cmd.log(vorpal.chalk.green('DONE'));
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

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    You can retrieve information about a user, either by specifying that user's
    id or user name (${chalk.grey(`userPrincipalName`)}), but not both.

    If the user with the specified id or user name doesn't exist, you will get
    a ${chalk.grey(`Resource 'xyz' does not exist or one of its queried reference-property`)}
    ${chalk.grey(`objects are not present.`)} error.

  Examples:

    Get information about the user with id ${chalk.grey(`1caf7dcd-7e83-4c3a-94f7-932a1299c844`)}
      ${this.name} --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844

    Get information about the user with user name ${chalk.grey(`AarifS@contoso.onmicrosoft.com`)}
      ${this.name} --userName AarifS@contoso.onmicrosoft.com

    For the user with id ${chalk.grey(`1caf7dcd-7e83-4c3a-94f7-932a1299c844`)}
    retrieve the user name, e-mail address and full name
      ${this.name} --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --properties "userPrincipalName,mail,displayName"

  More information:

    Microsoft Graph User properties
      https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/user#properties
`);
  }
}

module.exports = new AadUserGetCommand();
