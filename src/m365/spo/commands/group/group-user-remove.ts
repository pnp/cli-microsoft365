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
  groupId?: number;
  groupName?: string;
  userName: string;
  confirm?: boolean;
}

class SpoGroupUserRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_USER_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified user from a SharePoint group';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        groupId: (!(!args.options.groupId)).toString(),
        groupName: (!(!args.options.groupName)).toString(),
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
        option: '--groupId [groupId]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '--userName <userName>'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && isNaN(args.options.groupId)) {
          return `Specified "groupId" ${args.options.groupId} is not valid`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['groupName', 'groupId']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const removeUserfromSPGroup = (): void => {
      if (this.verbose) {
        logger.logToStderr(`Removing User with Username ${args.options.userName} from Group: ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
      }

      const loginName: string = `i:0#.f|membership|${args.options.userName}`;
      const requestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/${args.options.groupId
        ? `GetById('${encodeURIComponent(args.options.groupId)}')`
        : `GetByName('${encodeURIComponent(args.options.groupName as string)}')`}/users/removeByLoginName(@LoginName)?@LoginName='${encodeURIComponent(loginName)}'`;

      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .post(requestOptions)
        .then((): void => {
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      if (this.debug) {
        logger.logToStderr('Confirmation bypassed by entering confirm option. Removing the user from SharePoint Group...');
      }
      removeUserfromSPGroup();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove user User ${args.options.userName} from SharePoint group?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeUserfromSPGroup();
        }
      });
    }
  }
}

module.exports = new SpoGroupUserRemoveCommand();