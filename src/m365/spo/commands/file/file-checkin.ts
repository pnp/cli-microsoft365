import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.url = (!(!args.options.fileUrl)).toString();
    telemetryProps.type = args.options.type || 'Major';
    telemetryProps.comment = typeof args.options.comment !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
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
      comment = encodeURIComponent(args.options.comment);
    }

    let requestUrl: string = '';
    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${encodeURIComponent(args.options.id)}')/checkin(comment='${comment}',checkintype=${type})`;
    }

    if (args.options.fileUrl) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(args.options.fileUrl)}')/checkin(comment='${comment}',checkintype=${type})`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .post(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log('DONE');
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the file is located'
      },
      {
        option: '-f, --fileUrl [fileUrl]',
        description: 'The server-relative URL of the file to retrieve. Specify either fileUrl or id but not both'
      },
      {
        option: '-i, --id [id]',
        description: 'The UniqueId (GUID) of the file to retrieve. Specify either fileUrl or id but not both'
      },
      {
        option: '-t, --type [type]',
        description: 'Type of the check in. Available values Minor|Major|Overwrite. Default is Major',
        autocomplete: ['Minor', 'Major', 'Overwrite']
      },
      {
        option: '--comment [comment]',
        description: 'Comment to set when checking the file in. Its length must be less than 1024 letters. Default is empty string'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.id) {
        if (!Utils.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }
      }

      if (args.options.id && args.options.fileUrl) {
        return 'Specify either fileUrl or id but not both';
      }

      if (!args.options.id && !args.options.fileUrl) {
        return 'Specify fileUrl or id, one is required';
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
    };
  }
}

module.exports = new SpoFileCheckinCommand();