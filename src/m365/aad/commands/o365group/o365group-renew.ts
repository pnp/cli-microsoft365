import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class AadO365GroupRenewCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_RENEW;
  }

  public get description(): string {
    return `Renews Microsoft 365 group's expiration`;
  }

  constructor() {
    super();
  
    this.#initOptions();
    this.#initValidators();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Renewing Microsoft 365 group's expiration: ${args.options.id}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups/${args.options.id}/renew/`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      }
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new AadO365GroupRenewCommand();