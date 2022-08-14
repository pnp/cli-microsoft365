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
  id?: string;
  loginName?: string;
  confirm: boolean;
}

class SpoUserRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_REMOVE;
  }

  public get description(): string {
    return 'Removes user from specific web';
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
        id: (!(!args.options.id)).toString(),
        loginName: (!(!args.options.loginName)).toString(),
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
        option: '-i, --id [id]'
      },
      {
        option: '--loginName [loginName]'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'loginName']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeUser = (): void => {
      if (this.verbose) {
        logger.logToStderr(`Removing user from  subsite ${args.options.webUrl} ...`);
      }

      let requestUrl: string = '';

      if (args.options.id) {
        requestUrl = `${encodeURI(args.options.webUrl)}/_api/web/siteusers/removebyid(${args.options.id})`;
      }

      if (args.options.loginName) {
        requestUrl = `${encodeURI(args.options.webUrl)}/_api/web/siteusers/removeByLoginName('${encodeURIComponent(args.options.loginName as string)}')`;
      }

      const requestOptions: any = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .post(requestOptions)
        .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeUser();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove specified user from the site ${args.options.webUrl}`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeUser();
        }
      });
    }
  }
}

module.exports = new SpoUserRemoveCommand();