import { PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url?: string;
  alias?: string;
  taskId: string;
  confirm?: boolean;
}

class PlannerTaskReferenceRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_REFERENCE_REMOVE;
  }

  public get description(): string {
    return 'Removes the reference from the Planner task';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.url = typeof args.options.url !== 'undefined';
    telemetryProps.alias = typeof args.options.alias !== 'undefined';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (args.options.confirm) {
      this.removeReference(logger, args, cb);
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the reference from the Planner task?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          this.removeReference(logger, args, cb);
        }
      });
    }
  }

  private removeReference(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getTaskDetailsEtagAndUrl(args.options)
      .then(({ etag, url }) => {
        const requestOptionsTaskDetails: any = {
          url: `${this.resource}/v1.0/planner/tasks/${args.options.taskId}/details`,
          headers: {
            'accept': 'application/json;odata.metadata=none',
            'If-Match': etag,
            'Prefer': 'return=representation'
          },
          responseType: 'json',
          data: {
            references: {
              [formatting.openTypesEncoder(url)]: null
            }
          }
        };

        return request.patch(requestOptionsTaskDetails);
      })
      .then((): void => {
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTaskDetailsEtagAndUrl(options: Options): Promise<{ etag: string, url: string }> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(options.taskId)}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };
    
    let url: string = options.url!;

    return request
      .get<PlannerTaskDetails>(requestOptions)
      .then((taskDetails: PlannerTaskDetails) => {        
        if (options.alias) {
          const urls: string[] = [];

          if (taskDetails.references) {
            Object.entries(taskDetails.references!).forEach((ref: any) => {
              if (ref[1].alias?.toLocaleLowerCase() === options.alias!.toLocaleLowerCase()) {
                urls.push(decodeURIComponent(ref[0]));
              }
            });
          }

          if (urls.length === 0) {
            return Promise.reject(`The specified reference with alias ${options.alias} does not exist`);
          }

          if (urls.length > 1) {
            return Promise.reject(`Multiple references with alias ${options.alias} found. Pass one of the following urls within the "--url" option : ${urls}`);
          }

          url = urls[0];
        }

        return Promise.resolve({ etag: (taskDetails as any)['@odata.etag'], url });
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-u, --url [url]' },
      { option: '--alias [alias]' },
      { option: '-i, --taskId <taskId>' },
      { option: '--confirm' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public optionSets(): string[][] | undefined {
    return [
      ['url', 'alias']
    ];
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.url && args.options.url.indexOf('https://') !== 0 && args.options.url.indexOf('http://') !== 0) {
      return 'The url option should contain a valid URL. A valid URL starts with http(s)://';
    }

    return true;
  }
}

module.exports = new PlannerTaskReferenceRemoveCommand();