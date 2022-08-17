import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ContextInfo, spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  description?: string;
  content: string;
}

class SpoSiteScriptAddCommand extends SpoCommand {
  public get name(): string {
    return commands.SITESCRIPT_ADD;
  }

  public get description(): string {
    return 'Adds site script for use with site designs';
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
        description: typeof args.options.description !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title <title>'
      },
      {
        option: '-c, --content <content>'
      },
      {
        option: '-d, --description [description]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        try {
          JSON.parse(args.options.content);
        }
        catch (e) {
          return `Specified content value is not a valid JSON string. Error: ${e}`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let spoUrl: string = '';

    spo
      .getSpoUrl(logger, this.debug)
      .then((_spoUrl: string): Promise<ContextInfo> => {
        spoUrl = _spoUrl;
        return spo.getRequestDigest(spoUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title, Description=@description)?@title='${encodeURIComponent(args.options.title)}'&@description='${encodeURIComponent(args.options.description || '')}'`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          data: JSON.parse(args.options.content),
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoSiteScriptAddCommand();