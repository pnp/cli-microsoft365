import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  url?: string;
  id?: string;
  type?: string;
  comment?: string;
}

enum CheckinType {
  Minor = 0,
  Major = 1,
  Overwrite = 2,
}

class SpoFileCheckinCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_CHECKIN;
  }

  public get description(): string {
    return 'Checks in specified file';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: (!(!args.options.id)).toString(),
        url: (!(!args.options.url)).toString(),
        type: args.options.type || 'Major',
        comment: typeof args.options.comment !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-f, --url [url]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --type [type]',
        autocomplete: ['Minor', 'Major', 'Overwrite']
      },
      {
        option: '--comment [comment]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.id) {
          if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
          }
        }

        if (args.options.comment && args.options.comment.length > 1023) {
          return 'The length of the comment must be less than 1024 letters';
        }

        if (args.options.type) {
          const allowedValues: string[] = ['minor', 'major', 'overwrite'];
          const type: string = args.options.type.toLowerCase();
          if (allowedValues.indexOf(type) === -1) {
            return 'Wrong type specified. Available values are Minor|Major|Overwrite';
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['url', 'id']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let type: CheckinType = CheckinType.Major;
    if (args.options.type) {
      switch (args.options.type.toLowerCase()) {
        case 'minor':
          type = CheckinType.Minor;
          break;
        case 'overwrite':
          type = CheckinType.Overwrite;
      }
    }

    let comment: string = '';
    if (args.options.comment) {
      comment = formatting.encodeQueryParameter(args.options.comment);
    }

    let requestUrl: string = '';
    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${formatting.encodeQueryParameter(args.options.id)}')/checkin(comment='${comment}',checkintype=${type})`;
    }

    if (args.options.url) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(args.options.url)}')/checkin(comment='${comment}',checkintype=${type})`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoFileCheckinCommand();