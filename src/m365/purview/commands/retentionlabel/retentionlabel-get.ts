import { Logger } from '../../../../cli/Logger';
import GraphCommand from '../../../base/GraphCommand';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import request from '../../../../request';
import { AxiosRequestConfig } from 'axios';
import { validation } from '../../../../utils/validation';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class PurviewRetentionLabelGetCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONLABEL_GET;
  }

  public get description(): string {
    return 'Retrieve the specified retention label';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `'${args.options.id}' is not a valid GUID.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Retrieving retention label with id ${args.options.id}`);
      }

      const requestOptions: AxiosRequestConfig = {
        url: `${this.resource}/beta/security/labels/retentionLabels/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };
      const res: any = await request.get<any>(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PurviewRetentionLabelGetCommand();