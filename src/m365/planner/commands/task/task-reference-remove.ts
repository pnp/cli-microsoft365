import * as _ from 'lodash';
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
      .then(({etag, url}) => {
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
      .then((res: any): void => {
        logger.log(res.references);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTaskDetailsEtagAndUrl(options: Options): Promise<{etag: string, url: string}> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(options.taskId)}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };
    let url: string = options.url!;

    return request
      .get(requestOptions)
      .then((response: any) => {
        const etag: string | undefined = response ? response['@odata.etag'] : undefined;

        if (!etag) {
          return Promise.reject(`Error fetching task details`);
        }

        if (options.alias) {
          const alias = options.alias as string;
          const urls: string[] = [];

          _.each(response.references, (ref, key) => {
            if (ref.alias?.toLocaleLowerCase() === alias.toLocaleLowerCase()) {
              urls.push(decodeURIComponent(key));
            }
          });

          if (!urls.length) {
            return Promise.reject(`The specified reference with alias ${options.alias} does not exist`);
          }
  
          if (urls.length > 1) {
            return Promise.reject(`Multiple references with alias ${options.alias} found: ${urls}`);
          }

          url = urls[0];
        }

        return Promise.resolve({etag, url});
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
}

module.exports = new PlannerTaskReferenceRemoveCommand();
