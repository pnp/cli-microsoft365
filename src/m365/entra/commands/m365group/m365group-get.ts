import { z } from "zod";
import { Logger } from "../../../../cli/Logger.js";
import { globalOptionsZod } from "../../../../Command.js";
import request, { CliRequestOptions } from "../../../../request.js";
import { entraGroup } from "../../../../utils/entraGroup.js";
import { zod } from "../../../../utils/zod.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import { GroupExtended } from "./GroupExtended.js";

const options = globalOptionsZod
  .extend({
    id: zod.alias("i", z.string().uuid().optional()),
    displayName: zod.alias("n", z.string().optional()),
    withSiteUrl: z.boolean().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraM365GroupGetCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_GET;
  }

  public get description(): string {
    return "Gets information about the specified Microsoft 365 Group or Microsoft Teams team";
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema.refine((options) => [options.id, options.displayName].filter(Boolean).length === 1, {
      message: "Specify either id or displayName"
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let group: GroupExtended;

    try {
      if (args.options.id) {
        group = await entraGroup.getGroupById(args.options.id);
      }
      else {
        group = await entraGroup.getGroupByDisplayName(args.options.displayName!);
      }

      const isUnifiedGroup = await entraGroup.isUnifiedGroup(group.id!);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${group.id}' is not a Microsoft 365 group.`);
      }

      const requestExtendedOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groups/${group.id}?$select=allowExternalSenders,autoSubscribeNewMembers,hideFromAddressLists,hideFromOutlookClients,isSubscribedByMail`,
        headers: {
          accept: "application/json;odata.metadata=none"
        },
        responseType: "json"
      };
      const groupExtended = await request.get<{ allowExternalSenders: boolean; autoSubscribeNewMembers: boolean; hideFromAddressLists: boolean; hideFromOutlookClients: boolean; isSubscribedByMail: boolean }>(requestExtendedOptions);
      group = { ...group, ...groupExtended };

      if (args.options.withSiteUrl) {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${group.id}/drive?$select=webUrl`,
          headers: {
            accept: "application/json;odata.metadata=none"
          },
          responseType: "json"
        };

        const res = await request.get<{ webUrl: string }>(requestOptions);
        group.siteUrl = res.webUrl ? res.webUrl.substring(0, res.webUrl.lastIndexOf("/")) : "";
      }

      await logger.log(group);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraM365GroupGetCommand();
