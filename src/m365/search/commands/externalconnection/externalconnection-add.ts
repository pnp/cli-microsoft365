import { Logger } from "../../../../cli";
import { CommandOption } from "../../../../Command";
import GlobalOptions from "../../../../GlobalOptions";
import request from "../../../../request";
import GraphCommand from "../../../base/GraphCommand";
import commands from "../../commands";
import { ExternalConnectors } from "@microsoft/microsoft-graph-types/microsoft-graph";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  name: string;
  description?: string;
  authorizedAppIds?: string;
}

class SearchExternalConnectionAddCommand extends GraphCommand {
  public get name(): string {
    return commands.EXTERNALCONNECTION_ADD;
  }

  public get description(): string {
    return "Adds a new External Connection for Microsoft Search";
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ["id", "name", "description"];
  }

  public commandAction(
    logger: Logger,
    args: CommandArgs,
    cb: () => void
  ): void {

    let appIds: string[] = [];

    if (
      args.options.authorizedAppIds !== undefined &&
      args.options.authorizedAppIds !== ""
    ) {
      appIds = args.options.authorizedAppIds?.split(",");
    }

    const commandData: ExternalConnectors.ExternalConnection  = {
      id: args.options.id,
      name: args.options.name,
      description: args.options.description,
      configuration: {
        authorizedAppIds: appIds
      }
    };

    const requestOptions: any = {
      url: `${this.resource}/v1.0/external/connections`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: commandData
    };

    request.post(requestOptions).then(
      (_) => cb(),
      (err: any) => {
        this.handleRejectedODataJsonPromise(err, logger, cb);
      }
    );
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: "-i --id <id>"
      },
      {
        option: "-n --name <name>"
      },
      {
        option: "-d --description <description>"
      },
      {
        option: "--authorizedAppIds [authorizedAppIds]"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.id !== undefined) {
      const idToValidate = args.options.id.toString();
      if (idToValidate.length < 3 || idToValidate.length > 32) {
        return "ID field must be between 3 and 32 characters in length.";
      }

      //var alphanumeric = "someStringHere";
      const alphaNumericRegEx = /[^\w]|_/g;

      if (alphaNumericRegEx.test(idToValidate)) {
        return "ID field must only contain alphanumeric characters.";
      }

      if (
        idToValidate.length > 9 &&
        idToValidate.startsWith("Microsoft")
      ) {
        return "ID field cannot begin with Microsoft";
      }

      if (
        idToValidate === "None" ||
        idToValidate === "Directory" ||
        idToValidate === "Exchange" ||
        idToValidate === "ExchangeArchive" ||
        idToValidate === "LinkedIn" ||
        idToValidate === "Mailbox" ||
        idToValidate === "OneDriveBusiness" ||
        idToValidate === "SharePoint" ||
        idToValidate === "Teams" ||
        idToValidate === "Yammer" ||
        idToValidate === "Connectors" ||
        idToValidate === "TaskFabric" ||
        idToValidate === "PowerBI" ||
        idToValidate === "Assistant" ||
        idToValidate === "TopicEngine" ||
        idToValidate === "MSFT_All_Connectors"
      ) {
        return "ID field cannot be one of the following values: None, Directory, Exchange, ExchangeArchive, LinkedIn, Mailbox, OneDriveBusiness, SharePoint, Teams, Yammer, Connectors, TaskFabric, PowerBI, Assistant, TopicEngine, MSFT_All_Connectors.";
      }
    }
    return true;
  }
}

module.exports = new SearchExternalConnectionAddCommand();
