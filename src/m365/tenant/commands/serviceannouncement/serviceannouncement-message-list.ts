import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import { ServiceUpdateMessage } from '@microsoft/microsoft-graph-types';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  service: string;
}

class TenantServiceAnnouncementMessageListCommand extends GraphItemsListCommand<ServiceUpdateMessage> {
  public get name(): string {
    return commands.SERVICEANNOUNCEMENT_MESSAGE_LIST;
  }

  public get description(): string {
    return 'Gets all service update messages that exist for the tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let endpoint: string = `${this.resource}/v1.0/admin/serviceAnnouncement/messages`;

    if (args.options.service) {
      endpoint += `?$filter=services/any(c:c+eq+'${encodeURIComponent(args.options.service)}')`;
    }

    const requestOptions: any = {
      url: endpoint,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        logger.log(res.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-s, --service [service]	'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new TenantServiceAnnouncementMessageListCommand();