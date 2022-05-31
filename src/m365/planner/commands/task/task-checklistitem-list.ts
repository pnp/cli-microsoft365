import { Logger } from "../../../../cli";
import { CommandOption } from "../../../../Command";
import { accessToken } from "../../../../utils";
import GlobalOptions from "../../../../GlobalOptions";
import Auth from "../../../../Auth";
import request from "../../../../request";
import GraphCommand from "../../../base/GraphCommand";
import commands from "../../commands";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  taskId: string;
}

class PlannerTaskChecklistitemListCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_CHECKLISTITEM_LIST;
  }

  public get description(): string {
    return "Lists the checklist items of a Planner task.";
  }
  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'isChecked'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (accessToken.isAppOnlyAccessToken(Auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.', logger, cb);
      return;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(args.options.taskId)}/details?$select=checklist`,
      headers: {
        accept: "application/json;odata.metadata=none"
      },
      responseType: "json"
    };

    request.get(requestOptions).then(
      (res: any): void => {
        let checklistitems = res.checklist;
        if (args.options.output === 'text') {
          checklistitems = this.ModifyToJsonArray(res.checklist);
        }
        logger.log(checklistitems);
        cb();
      },
      (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb)
    );
  }
  private ModifyToJsonArray(checklist: any): any[] {
    const checklistitems = [];
    for (const item in checklist) {
      const checklistitem: any = {};
      checklistitem["id"] = item;
      for (const obj in checklist[item]) {
        checklistitem[obj] = checklist[item][obj];
      }
      checklistitems.push(checklistitem);
    }
    return checklistitems;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: "-i, --taskId <taskId>"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new PlannerTaskChecklistitemListCommand();
