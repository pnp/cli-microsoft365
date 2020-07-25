import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import Utils from '../../../../Utils';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
  sharePointTheme?: boolean;
}

const SharePointThemes = {
  Blue: "Blue",
  Orange: "Orange",
  Red: "Red",
  Purple: "Purple",
  Green: "Green",
  Gray: "Gray",
  "Dark Yellow": "Dark Yellow",
  "Dark Blue": "Dark Blue"
}

class SpoThemeApplyCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_APPLY;
  }

  public get description(): string {
    return 'Applies theme to the specified site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const isSharePointTheme: boolean = args.options.sharePointTheme ? true : false;
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        if (isSharePointTheme) {
          return Promise.resolve(undefined as any);
        }

        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Applying theme ${args.options.name} to the ${args.options.webUrl} site...`);
        }

        let requestOptions: any = {};

        if (isSharePointTheme) {
          const requestBody: any = this.getSharePointTheme(args.options.name);

          requestOptions = {
            url: `${args.options.webUrl}/_api/ThemeManager/ApplyTheme`,
            headers: {
              'accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata'
            },
            body: requestBody
          }
        }
        else {
          requestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="SetWebTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.name)}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.webUrl)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
          };
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        if (isSharePointTheme) {
          const json: any = JSON.parse(res);

          if (json.error) {
            cb(new CommandError(json.error));
            return;
          }
          else {
            cmd.log(json.value);
          }
        }
        else {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];

          if (response.ErrorInfo) {
            cb(new CommandError(response.ErrorInfo.ErrorMessage));
            return;
          }
          else {
            const result: boolean = json[json.length - 1];
            cmd.log(result);
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-n, --name <name>',
      description: 'Name of the theme to apply'
    },
    {
      option: '-u, --webUrl <webUrl>',
      description: 'URL of the site to which the theme should be applied'
    },
    {
      option: '--sharePointTheme',
      description: 'Set to specify if the supplied theme name is a standard SharePoint theme'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.sharePointTheme && !(args.options.name in SharePointThemes)) {
        return 'Please check if the theme name is entered correctly.'
      }

      return true;
    };
  }

  private getSharePointTheme(themeName: string): any {
    let palette: any = ""

    switch (themeName) {
      case SharePointThemes.Blue:
        palette = "\"themePrimary\":{\"R\":0,\"G\":120,\"B\":212,\"A\":255},\"themeLighterAlt\":{\"R\":239,\"G\":246,\"B\":252,\"A\":255},\"themeLighter\":{\"R\":222,\"G\":236,\"B\":249,\"A\":255},\"themeLight\":{\"R\":199,\"G\":224,\"B\":244,\"A\":255},\"themeTertiary\":{\"R\":113,\"G\":175,\"B\":229,\"A\":255},\"themeSecondary\":{\"R\":43,\"G\":136,\"B\":216,\"A\":255},\"themeDarkAlt\":{\"R\":16,\"G\":110,\"B\":190,\"A\":255},\"themeDark\":{\"R\":0,\"G\":90,\"B\":158,\"A\":255},\"themeDarker\":{\"R\":0,\"G\":69,\"B\":120,\"A\":255},\"accent\":{\"R\":135,\"G\":100,\"B\":184,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}";
        break;
      case SharePointThemes.Orange:
        palette = "\"themePrimary\":{\"R\":202,\"G\":80,\"B\":16,\"A\":255},\"themeLighterAlt\":{\"R\":253,\"G\":247,\"B\":244,\"A\":255},\"themeLighter\":{\"R\":246,\"G\":223,\"B\":210,\"A\":255},\"themeLight\":{\"R\":239,\"G\":196,\"B\":173,\"A\":255},\"themeTertiary\":{\"R\":223,\"G\":143,\"B\":100,\"A\":255},\"themeSecondary\":{\"R\":208,\"G\":98,\"B\":40,\"A\":255},\"themeDarkAlt\":{\"R\":181,\"G\":73,\"B\":15,\"A\":255},\"themeDark\":{\"R\":153,\"G\":62,\"B\":12,\"A\":255},\"themeDarker\":{\"R\":113,\"G\":45,\"B\":9,\"A\":255},\"accent\":{\"R\":152,\"G\":111,\"B\":11,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}"
        break;
      case SharePointThemes.Red:
        palette = "\"themePrimary\":{\"R\":164,\"G\":38,\"B\":44,\"A\":255},\"themeLighterAlt\":{\"R\":251,\"G\":244,\"B\":244,\"A\":255},\"themeLighter\":{\"R\":240,\"G\":211,\"B\":212,\"A\":255},\"themeLight\":{\"R\":227,\"G\":175,\"B\":178,\"A\":255},\"themeTertiary\":{\"R\":200,\"G\":108,\"B\":112,\"A\":255},\"themeSecondary\":{\"R\":174,\"G\":56,\"B\":62,\"A\":255},\"themeDarkAlt\":{\"R\":147,\"G\":34,\"B\":39,\"A\":255},\"themeDark\":{\"R\":124,\"G\":29,\"B\":33,\"A\":255},\"themeDarker\":{\"R\":91,\"G\":21,\"B\":25,\"A\":255},\"accent\":{\"R\":202,\"G\":80,\"B\":16,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}";
        break;
      case SharePointThemes.Purple:
        palette = "\"themePrimary\":{\"R\":135,\"G\":100,\"B\":184,\"A\":255},\"themeLighterAlt\":{\"R\":249,\"G\":248,\"B\":252,\"A\":255},\"themeLighter\":{\"R\":233,\"G\":226,\"B\":244,\"A\":255},\"themeLight\":{\"R\":215,\"G\":201,\"B\":234,\"A\":255},\"themeTertiary\":{\"R\":178,\"G\":154,\"B\":212,\"A\":255},\"themeSecondary\":{\"R\":147,\"G\":114,\"B\":192,\"A\":255},\"themeDarkAlt\":{\"R\":121,\"G\":89,\"B\":165,\"A\":255},\"themeDark\":{\"R\":102,\"G\":75,\"B\":140,\"A\":255},\"themeDarker\":{\"R\":75,\"G\":56,\"B\":103,\"A\":255},\"accent\":{\"R\":3,\"G\":131,\"B\":135,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}";
        break;
      case SharePointThemes.Green:
        palette = "\"themePrimary\":{\"R\":73,\"G\":130,\"B\":5,\"A\":255},\"themeLighterAlt\":{\"R\":246,\"G\":250,\"B\":240,\"A\":255},\"themeLighter\":{\"R\":219,\"G\":235,\"B\":199,\"A\":255},\"themeLight\":{\"R\":189,\"G\":218,\"B\":155,\"A\":255},\"themeTertiary\":{\"R\":133,\"G\":180,\"B\":76,\"A\":255},\"themeSecondary\":{\"R\":90,\"G\":145,\"B\":23,\"A\":255},\"themeDarkAlt\":{\"R\":66,\"G\":117,\"B\":5,\"A\":255},\"themeDark\":{\"R\":56,\"G\":99,\"B\":4,\"A\":255},\"themeDarker\":{\"R\":41,\"G\":73,\"B\":3,\"A\":255},\"accent\":{\"R\":3,\"G\":131,\"B\":135,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}";
        break;
      case SharePointThemes.Gray:
        palette = "\"themePrimary\":{\"R\":105,\"G\":121,\"B\":126,\"A\":255},\"themeLighterAlt\":{\"R\":248,\"G\":249,\"B\":250,\"A\":255},\"themeLighter\":{\"R\":228,\"G\":233,\"B\":234,\"A\":255},\"themeLight\":{\"R\":205,\"G\":213,\"B\":216,\"A\":255},\"themeTertiary\":{\"R\":159,\"G\":173,\"B\":177,\"A\":255},\"themeSecondary\":{\"R\":120,\"G\":136,\"B\":141,\"A\":255},\"themeDarkAlt\":{\"R\":93,\"G\":108,\"B\":112,\"A\":255},\"themeDark\":{\"R\":79,\"G\":91,\"B\":95,\"A\":255},\"themeDarker\":{\"R\":58,\"G\":67,\"B\":70,\"A\":255},\"accent\":{\"R\":0,\"G\":120,\"B\":212,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}";
        break;
      case SharePointThemes["Dark Yellow"]:
        palette = "\"themePrimary\":{\"R\":255,\"G\":200,\"B\":61,\"A\":255},\"themeLighterAlt\":{\"R\":10,\"G\":8,\"B\":2,\"A\":255},\"themeLighter\":{\"R\":41,\"G\":32,\"B\":10,\"A\":255},\"themeLight\":{\"R\":77,\"G\":60,\"B\":18,\"A\":255},\"themeTertiary\":{\"R\":153,\"G\":120,\"B\":37,\"A\":255},\"themeSecondary\":{\"R\":224,\"G\":176,\"B\":54,\"A\":255},\"themeDarkAlt\":{\"R\":255,\"G\":206,\"B\":81,\"A\":255},\"themeDark\":{\"R\":255,\"G\":213,\"B\":108,\"A\":255},\"themeDarker\":{\"R\":255,\"G\":224,\"B\":146,\"A\":255},\"accent\":{\"R\":255,\"G\":200,\"B\":61,\"A\":255},\"neutralLighterAlt\":{\"R\":40,\"G\":40,\"B\":40,\"A\":255},\"neutralLighter\":{\"R\":49,\"G\":49,\"B\":49,\"A\":255},\"neutralLight\":{\"R\":63,\"G\":63,\"B\":63,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":72,\"G\":72,\"B\":72,\"A\":255},\"neutralQuaternary\":{\"R\":79,\"G\":79,\"B\":79,\"A\":255},\"neutralTertiaryAlt\":{\"R\":109,\"G\":109,\"B\":109,\"A\":255},\"neutralTertiary\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralSecondary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralPrimaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralPrimary\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"neutralDark\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"black\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"white\":{\"R\":31,\"G\":31,\"B\":31,\"A\":255},\"primaryBackground\":{\"R\":31,\"G\":31,\"B\":31,\"A\":255},\"primaryText\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255}";
        break;
      case SharePointThemes["Dark Blue"]:
        palette = "\"themePrimary\":{\"R\":58,\"G\":150,\"B\":221,\"A\":255},\"themeLighterAlt\":{\"R\":2,\"G\":6,\"B\":9,\"A\":255},\"themeLighter\":{\"R\":9,\"G\":24,\"B\":35,\"A\":255},\"themeLight\":{\"R\":17,\"G\":45,\"B\":67,\"A\":255},\"themeTertiary\":{\"R\":35,\"G\":90,\"B\":133,\"A\":255},\"themeSecondary\":{\"R\":51,\"G\":133,\"B\":195,\"A\":255},\"themeDarkAlt\":{\"R\":75,\"G\":160,\"B\":225,\"A\":255},\"themeDark\":{\"R\":101,\"G\":174,\"B\":230,\"A\":255},\"themeDarker\":{\"R\":138,\"G\":194,\"B\":236,\"A\":255},\"accent\":{\"R\":58,\"G\":150,\"B\":221,\"A\":255},\"neutralLighterAlt\":{\"R\":29,\"G\":43,\"B\":60,\"A\":255},\"neutralLighter\":{\"R\":34,\"G\":50,\"B\":68,\"A\":255},\"neutralLight\":{\"R\":43,\"G\":61,\"B\":81,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":50,\"G\":68,\"B\":89,\"A\":255},\"neutralQuaternary\":{\"R\":55,\"G\":74,\"B\":95,\"A\":255},\"neutralTertiaryAlt\":{\"R\":79,\"G\":99,\"B\":122,\"A\":255},\"neutralTertiary\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralSecondary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralPrimaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralPrimary\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"neutralDark\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"black\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"white\":{\"R\":24,\"G\":37,\"B\":52,\"A\":255},\"primaryBackground\":{\"R\":24,\"G\":37,\"B\":52,\"A\":255},\"primaryText\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255}";
        break;
      default:
        palette = "";
        break;
    }

    return `{
      'name': '${themeName}' ,
      'themeJson': '{\"palette\": {${palette}}}'
    }`
  }
}

module.exports = new SpoThemeApplyCommand();