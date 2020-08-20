import config from "../../../../config";
import commands from "../../commands";
import GlobalOptions from "../../../../GlobalOptions";
import request from "../../../../request";
import { CommandOption, CommandValidate } from "../../../../Command";
import SpoCommand from "../../../base/SpoCommand";
import {
  ContextInfo,
  ClientSvcResponse,
  ClientSvcResponseContents,
} from "../../spo";
import * as fs from "fs";
import * as path from "path";
import Utils from "../../../../Utils";

const vorpal: Vorpal = require("../../../../vorpal-init");

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  filePath: string;
  isInverted: boolean;
}

class SpoThemeSetCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_SET;
  }

  public get description(): string {
    return "Add or update a theme";
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.inverted = (!!args.options.isInverted).toString();
    return telemetryProps;
  }

  public commandAction(
    cmd: CommandInstance,
    args: CommandArgs,
    cb: () => void
  ): void {
    let spoAdminUrl: string = "";

    this.getSpoAdminUrl(cmd, this.debug)
      .then(
        (_spoAdminUrl: string): Promise<ContextInfo> => {
          spoAdminUrl = _spoAdminUrl;
          return this.getRequestDigest(spoAdminUrl);
        }
      )
      .then(
        (res: ContextInfo): Promise<string> => {
          const fullPath: string = path.resolve(args.options.filePath);

          if (this.verbose) {
            cmd.log(`Adding theme from ${fullPath} to tenant...`);
          }

          const palette: any = JSON.parse(fs.readFileSync(fullPath, "utf8"));

          if (this.debug) {
            cmd.log("");
            cmd.log("Palette");
            cmd.log(JSON.stringify(palette));
          }

          const isInverted: boolean = args.options.isInverted ? true : false;

          const requestOptions: any = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              "X-RequestDigest": res.FormDigestValue,
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${
              config.applicationName
            }" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="UpdateTenantTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">${Utils.escapeXml(
              args.options.name
            )}</Parameter><Parameter Type="String">{"isInverted":${isInverted},"name":"${Utils.escapeXml(
              args.options.name
            )}","palette":${JSON.stringify(
              palette
            )}}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`,
          };

          return request.post(requestOptions);
        }
      )
      .then(
        (res: string): Promise<void> => {
          const json: ClientSvcResponse = JSON.parse(res);
          const contents: ClientSvcResponseContents = json.find((x) => {
            return x["ErrorInfo"];
          });

          if (contents && contents.ErrorInfo) {
            return Promise.reject(
              contents.ErrorInfo.ErrorMessage || "ClientSvc unknown error"
            );
          }
          return Promise.resolve();
        }
      )
      .then(
        (): void => {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green("DONE"));
          }

          cb();
        },
        (err: any): void => this.handleRejectedPromise(err, cmd, cb)
      );
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: "-n, --name <name>",
        description: "Name of the theme to add or update",
      },
      {
        option: "-p, --filePath <filePath>",
        description: "Absolute or relative path to the theme json file",
      },
      {
        option: "--isInverted",
        description: "Set to specify that the theme is inverted",
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return "Required parameter name missing";
      }

      if (!args.options.filePath) {
        return "Required parameter filePath missing";
      }

      const fullPath: string = path.resolve(args.options.filePath);

      if (!fs.existsSync(fullPath)) {
        return `File '${fullPath}' not found`;
      }

      if (fs.lstatSync(fullPath).isDirectory()) {
        return `Path '${fullPath}' points to a directory`;
      }

      if (!Utils.isValidTheme(fs.readFileSync(fullPath, "utf-8"))) {
        return "File contents is not a valid theme";
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow(
        "Important:"
      )} to use this command you have to have permissions to access
    the tenant admin site.
    
  Examples:
  
    Add or update a theme from a theme JSON file
      ${
        commands.THEME_SET
      } --name Contoso-Blue --filePath /Users/rjesh/themes/contoso-blue.json

    Add or update an inverted theme from a theme JSON file
      ${
        commands.THEME_SET
      } --name Contoso-Blue --filePath /Users/rjesh/themes/contoso-blue.json --isInverted

    A valid theme object is as follows:

    {
      "themePrimary": "#d81e05",
      "themeLighterAlt": "#fdf5f4",
      "themeLighter": "#f9d6d2",
      "themeLight": "#f4b4ac",
      "themeTertiary": "#e87060",
      "themeSecondary": "#dd351e",
      "themeDarkAlt": "#c31a04",
      "themeDark": "#a51603",
      "themeDarker": "#791002",
      "neutralLighterAlt": "#eeeeee",
      "neutralLighter": "#f5f5f5",
      "neutralLight": "#e1e1e1",
      "neutralQuaternaryAlt": "#d1d1d1",
      "neutralQuaternary": "#c8c8c8",
      "neutralTertiaryAlt": "#c0c0c0",
      "neutralTertiary": "#c2c2c2",
      "neutralSecondary": "#858585",
      "neutralPrimaryAlt": "#4b4b4b",
      "neutralPrimary": "#333333",
      "neutralDark": "#272727",
      "black": "#1d1d1d",
      "white": "#f5f5f5"
    }

    Validation checks the following:

    The specified string is a valid JSON string
    The deserialized object contains all properties defined in the above example
    The deserialized object doesn't contain any other properties
    Each property of the deserialized object contains a valid hex color value prefixed with a #

  More information:

    SharePoint site theming
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview

    Theme Generator
      https://aka.ms/themedesigner
      `
    );
  }
}

module.exports = new SpoThemeSetCommand();
