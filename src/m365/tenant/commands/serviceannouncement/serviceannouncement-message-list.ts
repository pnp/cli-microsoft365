import { ServiceUpdateMessage } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  service?: string;
}

class TenantServiceAnnouncementMessageListCommand extends GraphCommand {
  public get name(): string {
    return commands.SERVICEANNOUNCEMENT_MESSAGE_LIST;
  }

  public get description(): string {
    return 'Gets all service update messages for the tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        service: typeof args.options.service !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-s, --service [service]'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let endpoint: string = `${this.resource}/v1.0/admin/serviceAnnouncement/messages`;

    if (args.options.service) {
      endpoint += `?$filter=services/any(c:c+eq+'${formatting.encodeQueryParameter(args.options.service)}')`;
    }

    try {
      const items: ServiceUpdateMessage[] = await odata.getAllItems<ServiceUpdateMessage>(endpoint);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TenantServiceAnnouncementMessageListCommand();