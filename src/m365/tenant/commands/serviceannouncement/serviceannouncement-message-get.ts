import { Logger } from '../../../../cli';
import GraphCommand from '../../../base/GraphCommand';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  id: string;
}

class TenantServiceAnnouncementMessageGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SERVICEANNOUNCEMENT_MESSAGE_GET;
  }

  public get description(): string {
    return 'Retrieves a specified service update message for the tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['startDateTime', 'endDateTime', 'lastModifiedDateTime', 'title', 'id', 'category', 'severity', 'tags', 'isMajorChange', 'actionRequiredByDateTime', 'services', 'expiryDateTime', 'hasAttachments', 'viewPoint' ];
  }

  public commandAction(logger: Logger, args: any, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving service update message ${args.id}`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/admin/serviceAnnouncement/messages/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get<{ value: [{ id: string }] }>(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  private isValidId(id: string): boolean {
    return (/MC\d{6}/).test(id);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!this.isValidId(args.options.id)) {
      return `${args.options.id} is not a valid message ID`;
    }

    return true;
  }
}

module.exports = new TenantServiceAnnouncementMessageGetCommand();