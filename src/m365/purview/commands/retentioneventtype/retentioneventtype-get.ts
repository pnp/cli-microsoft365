import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GraphCommand from '../../../base/GraphCommand';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import { accessToken } from '../../../../utils/accessToken';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class PurviewRetentionEventtypeGetCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONEVENTTYPE_GET;
  }

  public get description(): string {
    return 'Retrieve the specified retention event type';
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
          return `'${args.options.id}' is not a valid GUID.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.');
    }

    try {
      if (this.verbose) {
        logger.logToStderr(`Retrieving retention event type with id ${args.options.id}`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/security/triggerTypes/retentionEventTypes/${args.options.id}`,
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

module.exports = new PurviewRetentionEventtypeGetCommand();