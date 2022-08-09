import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  taskId: string;
}

class PlannerTaskReferenceListCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_REFERENCE_LIST;
  }

  public get description(): string {
    return 'Retrieve the references of the specified planner task';
  }

  constructor() {
    super();
  
    this.#initOptions();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --taskId <taskId>'
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.', logger, cb);
      return;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(args.options.taskId)}/details?$select=references`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        logger.log(res.references);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new PlannerTaskReferenceListCommand();
