import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  loginName?: string;
  force: boolean;
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
        force: (!(!args.options.force)).toString()
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
        option: '-f, --force'
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
    this.optionSets.push({ options: ['id', 'loginName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeUser(logger, args.options);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove specified user from the site ${args.options.webUrl}`
      });

      if (result.continue) {
        await this.removeUser(logger, args.options);
      }
    }
  }

  private async removeUser(logger: Logger, options: GlobalOptions): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing user from  subsite ${options.webUrl} ...`);
    }

    let requestUrl: string = `${encodeURI(options.webUrl)}/_api/web/siteusers/`;
    if (options.id) {
      requestUrl += `removebyid(${options.id})`;
    }
    else if (options.loginName) {
      requestUrl += `removeByLoginName('${formatting.encodeQueryParameter(options.loginName as string)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoUserRemoveCommand();