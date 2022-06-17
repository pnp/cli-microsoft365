import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  id: string;
}

class TenantServiceAnnouncementMessageGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SERVICEANNOUNCEMENT_MESSAGE_GET;
  }

  public get description(): string {
    return 'Retrieves a specified service update message for the tenant';
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
        if (!this.isValidId(args.options.id)) {
          return `${args.options.id} is not a valid message ID`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving service update message ${args.options.id}`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/admin/serviceAnnouncement/messages/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private isValidId(id: string): boolean {
    return (/MC\d{6}/).test(id);
  }
}

module.exports = new TenantServiceAnnouncementMessageGetCommand();