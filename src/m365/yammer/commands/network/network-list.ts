import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  includeSuspended: boolean;
}

class YammerNetworkListCommand extends YammerCommand {
  public get name(): string {
    return `${commands.YAMMER_NETWORK_LIST}`;
  }

  public get description(): string {
    return 'Returns a list of networks to which the current user has access';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.includeSuspended = args.options.includeSuspended;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1/networks/current.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      json: true,
      body: {
        includeSuspended: args.options.includeSuspended !== undefined && args.options.includeSuspended !== false
      }
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (args.options.output === 'json') {
          logger.log(res);
        }
        else {
          logger.log((res as any[]).map(n => ({ id: n.id, name: n.name, email: n.email, community: n.community, permalink: n.permalink, web_url: n.web_url })));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--includeSuspended',
        description: 'Include the networks in which the user is suspended'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new YammerNetworkListCommand();