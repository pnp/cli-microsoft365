import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  issueId: string;
}

class TenantServiceAnnouncementHealthissueGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SERVICEANNOUNCEMENT_HEALTHISSUE_GET;
  }

  public get description(): string {
    return 'Gets a specified service health issue for tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title'];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.issueId = typeof args.options.issueId !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/admin/serviceAnnouncement/issues/${encodeURIComponent(args.options.issueId)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --issueId <issueId>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new TenantServiceAnnouncementHealthissueGetCommand(); 