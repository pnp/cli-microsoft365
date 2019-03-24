import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../GraphCommand';
import { Group } from './Group';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  includeSiteUrl: boolean;
}

class GraphO365GroupGetCommand extends GraphCommand {
  public get name(): string {
    return `${commands.O365GROUP_GET}`;
  }

  public get description(): string {
    return 'Gets information about the specified Office 365 Group';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let group: Group;

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): Promise<Group> => {
        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/groups/${args.options.id}`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res: Group): Promise<{ webUrl: string }> => {
        group = res;

        if (args.options.includeSiteUrl) {
          const requestOptions: any = {
            url: `${auth.service.resource}/v1.0/groups/${group.id}/drive?$select=webUrl`,
            headers: {
              authorization: `Bearer ${auth.service.accessToken}`,
              accept: 'application/json;odata.metadata=none'
            },
            json: true
          };

          return request.get(requestOptions);
        }
        else {
          return Promise.resolve(undefined as any);
        }
      })
      .then((res?: { webUrl: string }): void => {
        if (res) {
          group.siteUrl = res.webUrl ? res.webUrl.substr(0, res.webUrl.lastIndexOf('/')) : '';
        }

        cmd.log(group);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The ID of the Office 365 Group to retrieve information for'
      },
      {
        option: '--includeSiteUrl',
        description: 'Set to retrieve the site URL for the group'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To get information about a Office 365 Group, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:
  
    Get information about the Office 365 Group with id ${chalk.grey(`1caf7dcd-7e83-4c3a-94f7-932a1299c844`)}
      ${chalk.grey(config.delimiter)} ${this.name} --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844

    Get information about the Office 365 Group with id ${chalk.grey(`1caf7dcd-7e83-4c3a-94f7-932a1299c844`)}
    and also retrieve the URL of the corresponding SharePoint site
      ${chalk.grey(config.delimiter)} ${this.name} --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --includeSiteUrl
`);
  }
}

module.exports = new GraphO365GroupGetCommand();