import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { WebPropertiesCollection } from "./WebPropertiesCollection";

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  withGroups?: boolean;
}

class SpoWebGetCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_GET;
  }

  public get description(): string {
    return 'Retrieve information about the specified site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        withGroups: typeof args.options.withGroups !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--withGroups'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let url: string = `${args.options.webUrl}/_api/web`;
    if (args.options.withGroups) {
      url += '?$expand=AssociatedMemberGroup,AssociatedOwnerGroup,AssociatedVisitorGroup';
    }
    const requestOptions: any = {
      url,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const webProperties: WebPropertiesCollection = await request.get<WebPropertiesCollection>(requestOptions);
      logger.log(webProperties);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoWebGetCommand();