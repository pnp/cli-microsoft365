import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  groupId?: number;
  groupName?: string;
}

class SpoGroupUserListCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_USER_LIST;
  }

  public get description(): string {
    return `List members of a SharePoint Group`;
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'UserPrincipalName', 'Id', 'Email'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving the list of members from the SharePoint group :  ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
    }

    const requestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/${args.options.groupId
      ? `GetById('${encodeURIComponent(args.options.groupId)}')`
      : `GetByName('${encodeURIComponent(args.options.groupName as string)}')`}/users`;

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((response: any): void => {
        logger.log(response);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--groupId [groupId]'
      },
      {
        option: '--groupName [groupName]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.groupId && args.options.groupName) {
      return 'Use either "groupName" or "groupId", but not both';
    }

    if (!args.options.groupId && !args.options.groupName) {
      return 'Either "groupName" or "groupId" is required';
    }

    if (args.options.groupId && isNaN(args.options.groupId)) {
      return `Specified "groupId" ${args.options.groupId} is not valid`;
    }

    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }

}

module.exports = new SpoGroupUserListCommand();