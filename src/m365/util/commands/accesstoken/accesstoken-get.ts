import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';

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

    try {
      if (resource.toLowerCase() === 'sharepoint') {
        if (auth.service.spoUrl) {
          resource = auth.service.spoUrl;
        }
        else {
          throw `SharePoint URL undefined. Use the 'm365 spo set --url https://contoso.sharepoint.com' command to set the URL`;
        }
      }
  
      const accessToken: string = await auth.ensureAccessToken(resource, logger, this.debug, args.options.new);
      logger.log(accessToken);      
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new UtilAccessTokenGetCommand();