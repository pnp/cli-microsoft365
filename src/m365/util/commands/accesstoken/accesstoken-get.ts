import auth, { Auth } from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  new?: boolean;
  resource: string;
}

class UtilAccessTokenGetCommand extends Command {
  public get name(): string {
    return commands.ACCESSTOKEN_GET;
  }

  public get description(): string {
    return 'Gets access token for the specified resource';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        new: args.options.new
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-r, --resource <resource>'
      },
      {
        option: '--new'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let resource: string = args.options.resource;

    if (resource.toLowerCase() === 'sharepoint') {
      if (auth.service.spoUrl) {
        resource = auth.service.spoUrl;
      }
      else {
        throw `SharePoint URL undefined. Use the 'm365 spo set --url https://contoso.sharepoint.com' command to set the URL`;
      }
    }
    else if (resource.toLowerCase() === 'graph') {
      resource = Auth.getEndpointForResource('https://graph.microsoft.com', auth.service.cloudType);
    }

    try {
      const accessToken: string = await auth.ensureAccessToken(resource, logger, this.debug, args.options.new);
      await logger.log(accessToken);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new UtilAccessTokenGetCommand();