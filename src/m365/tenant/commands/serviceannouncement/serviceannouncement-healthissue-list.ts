import { ServiceHealthIssue } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  service?: string;
}

class TenantServiceAnnouncementHealthIssueListCommand extends GraphCommand {
  public get name(): string {
    return commands.SERVICEANNOUNCEMENT_HEALTHISSUE_LIST;
  }

  public get description(): string {
    return 'Gets all service health issues for the tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title'];
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-s, --service [service]'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let endpoint: string = `${this.resource}/v1.0/admin/serviceAnnouncement/issues`;

    if (args.options.service) {
      endpoint += `?$filter=service eq '${encodeURIComponent(args.options.service)}'`;
    }

    try {
      const items: any = await odata.getAllItems<ServiceHealthIssue>(endpoint);
      logger.log(items);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TenantServiceAnnouncementHealthIssueListCommand();