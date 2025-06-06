import auth, { Auth } from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import commands from '../../commands.js';
import { accessToken } from '../../../../utils/accessToken.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  new?: boolean;
  resource: string;
  decoded?: boolean;
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
        new: args.options.new,
        decoded: args.options.decoded
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
      },
      {
        option: '--decoded'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let resource: string = args.options.resource;

    if (resource.toLowerCase() === 'sharepoint') {
      if (auth.connection.spoUrl) {
        resource = auth.connection.spoUrl;
      }
      else {
        throw `SharePoint URL undefined. Use the 'm365 spo set --url https://contoso.sharepoint.com' command to set the URL`;
      }
    }
    else if (resource.toLowerCase() === 'graph') {
      resource = Auth.getEndpointForResource('https://graph.microsoft.com', auth.connection.cloudType);
    }

    try {
      const token: string = await auth.ensureAccessToken(resource, logger, this.debug, args.options.new);

      if (args.options.decoded) {
        const { header, payload } = accessToken.getDecodedAccessToken(token);

        await logger.logRaw(`${JSON.stringify(header, null, 2)}.${JSON.stringify(payload, null, 2)}.[signature]`);
      }
      else {
        await logger.log(token);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new UtilAccessTokenGetCommand();