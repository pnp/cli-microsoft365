import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ContextInfo, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  title?: string;
  description?: string;
  version?: string;
  content?: string;
}

class SpoSiteScriptSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITESCRIPT_SET;
  }

  public get description(): string {
    return 'Updates existing site script';
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
        title: (!(!args.options.title)).toString(),
        description: (!(!args.options.description)).toString(),
        version: (!(!args.options.version)).toString(),
        content: (!(!args.options.content)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '-v, --version [version]'
      },
      {
        option: '-c, --content [content]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.version) {
          const version: number = parseInt(args.options.version);
          if (isNaN(version)) {
            return `${args.options.version} is not a number`;
          }
        }

        if (args.options.content) {
          try {
            JSON.parse(args.options.content);
          }
          catch (e) {
            return `Specified content value is not a valid JSON string. Error: ${e}`;
          }
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
        const updateInfo: any = {
          Id: args.options.id
        };
        if (args.options.title) {
          updateInfo.Title = args.options.title;
        }
        if (args.options.description) {
          updateInfo.Description = args.options.description;
        }
        if (args.options.version) {
          updateInfo.Version = parseInt(args.options.version);
        }
        if (args.options.content) {
          updateInfo.Content = args.options.content;
        }

        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          data: { updateInfo: updateInfo },
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

module.exports = new SpoSiteScriptSetCommand();