import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import * as chalk from 'chalk';
import { telemetry } from '../../../../telemetry';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import path = require('path');
const command: Command = require('./package-teams-enable');

const admZipMock = {
  // we need these unused params so that they can be properly mocked with sinon
  /* eslint-disable @typescript-eslint/no-unused-vars */
  extractAllTo: (path: string) => { },
  addLocalFolder: (path: string) => { },
  toBuffer: () => { }
  /* eslint-enable @typescript-eslint/no-unused-vars */
};

describe(commands.PACKAGE_TEAMS_ENABLE, () => {
  const filePath = 'solution.sppkg';
  const fullPath = `C:\\temp\\${filePath}`;
  const supportedHost = 'TeamsPersonalApp';
  const tmpDir = '/tmp/cli-solution';
  const sppkgTopLevelContent = [
    "7eaef6a8-b579-4eba-bdf0-f0eb1591647d",
    "AppManifest.xml",
    "ClientSideAssets",
    "ClientSideAssets.xml",
    "ClientSideAssets.xml.config.xml",
    "feature_7eaef6a8-b579-4eba-bdf0-f0eb1591647d.xml",
    "feature_7eaef6a8-b579-4eba-bdf0-f0eb1591647d.xml.config.xml",
    "[Content_Types].xml",
    "_rels"
  ];
  const webpartXmls = [
    "WebPart_2a47a728-f0b5-4abc-9c6a-acae44cb0759.xml",
    "WebPart_2d6310ab-19be-453e-bad8-d8e648978b75.xml"
  ];
  const validWebPartId = '2a47a728-f0b5-4abc-9c6a-acae44cb0759';
  const validWebPartName = 'Valid WebPart';
  const validWebpartBody = `<?xml version="1.0" encoding="utf-8"?><Elements xmlns="http://schemas.microsoft.com/sharepoint/"><ClientSideComponent Name="${validWebPartName}" Id="${validWebPartId}" ComponentManifest="{&quot;id&quot;:&quot;${validWebPartId}&quot;,&quot;alias&quot;:&quot;${validWebPartName}&quot;,&quot;componentType&quot;:&quot;WebPart&quot;,&quot;version&quot;:&quot;0.0.1&quot;,&quot;manifestVersion&quot;:2,&quot;supportedHosts&quot;:[&quot;SharePointWebPart&quot;,&quot;TeamsPersonalApp&quot;],&quot;supportsThemeVariants&quot;:true,&quot;preconfiguredEntries&quot;:[{&quot;groupId&quot;:&quot;5c03119e-3074-46fd-976b-c60198311f70&quot;,&quot;group&quot;:{&quot;default&quot;:&quot;Advanced&quot;},&quot;title&quot;:{&quot;default&quot;:&quot;${validWebPartName}&quot;},&quot;description&quot;:{&quot;default&quot;:&quot;${validWebPartName}&quot;},&quot;officeFabricIconFontName&quot;:&quot;TaskSolid&quot;,&quot;properties&quot;:{&quot;description&quot;:&quot;${validWebPartName}&quot;}}],&quot;loaderConfig&quot;:{&quot;internalModuleBaseUrls&quot;:[&quot;HTTPS://SPCLIENTSIDEASSETLIBRARY/&quot;],&quot;entryModuleId&quot;:&quot;${validWebPartName}&quot;,&quot;scriptResources&quot;:{&quot;${validWebPartName}&quot;:{&quot;type&quot;:&quot;path&quot;,&quot;path&quot;:&quot;web-part_1ddcf13f7216f9a2a92a.js&quot;},&quot;@microsoft/sp-property-pane&quot;:{&quot;type&quot;:&quot;component&quot;,&quot;id&quot;:&quot;f9e737b7-f0df-4597-ba8c-3060f82380db&quot;,&quot;version&quot;:&quot;1.16.1&quot;},&quot;@microsoft/sp-lodash-subset&quot;:{&quot;type&quot;:&quot;component&quot;,&quot;id&quot;:&quot;73e1dc6c-8441-42cc-ad47-4bd3659f8a3a&quot;,&quot;version&quot;:&quot;1.16.1&quot;},&quot;@microsoft/sp-core-library&quot;:{&quot;type&quot;:&quot;component&quot;,&quot;id&quot;:&quot;7263c7d0-1d6a-45ec-8d85-d4d1d234171b&quot;,&quot;version&quot;:&quot;1.16.1&quot;},&quot;@microsoft/sp-webpart-base&quot;:{&quot;type&quot;:&quot;component&quot;,&quot;id&quot;:&quot;974a7777-0990-4136-8fa6-95d80114c2e0&quot;,&quot;version&quot;:&quot;1.16.1&quot;},&quot;react&quot;:{&quot;type&quot;:&quot;component&quot;,&quot;id&quot;:&quot;0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d&quot;,&quot;version&quot;:&quot;17.0.1&quot;},&quot;react-dom&quot;:{&quot;type&quot;:&quot;component&quot;,&quot;id&quot;:&quot;aa0a46ec-1505-43cd-a44a-93f3a5aa460a&quot;,&quot;version&quot;:&quot;17.0.1&quot;},&quot;ToDoWebPartStrings&quot;:{&quot;type&quot;:&quot;path&quot;,&quot;path&quot;:&quot;ToDoWebPartStrings_en-us_d32717b12d14d01f33fa09b9f868b0bd.js&quot;},&quot;PropertyControlStrings&quot;:{&quot;type&quot;:&quot;localizedPath&quot;,&quot;paths&quot;:{&quot;ar-SA&quot;:&quot;PropertyControlStrings_ar-sa_c2bfcb8203030a4e22b6ce832477cc8d.js&quot;,&quot;bg-BG&quot;:&quot;PropertyControlStrings_bg-bg_d769b67c49a2193b7589cf77ae8f9891.js&quot;,&quot;ca-ES&quot;:&quot;PropertyControlStrings_ca-es_ae71b607bdec2062787ba9403df796a5.js&quot;,&quot;cs-CZ&quot;:&quot;PropertyControlStrings_cs-cz_387455f8d307e56df7fbea367de5d179.js&quot;,&quot;da-DK&quot;:&quot;PropertyControlStrings_da-dk_d45860d6cfeaa3a704e0633d63cc9b38.js&quot;,&quot;de-DE&quot;:&quot;PropertyControlStrings_de-de_159e868a4e92f6f28f7fdb1bb1a7f4ad.js&quot;,&quot;el-GR&quot;:&quot;PropertyControlStrings_el-gr_94aeac5730842371272231dbd949b55a.js&quot;,&quot;en-GB&quot;:&quot;PropertyControlStrings_en-gb_b0c4348eae812c2a7229a4db129b9be6.js&quot;,&quot;en-US&quot;:&quot;PropertyControlStrings_en-us_c3915679a97436fd69e420b5b4064a25.js&quot;,&quot;es-ES&quot;:&quot;PropertyControlStrings_es-es_4745311c9839342364e154c4d6cdcd54.js&quot;,&quot;et-EE&quot;:&quot;PropertyControlStrings_et-ee_f2b12fcc0b1640b6c85b2f56e38be17b.js&quot;,&quot;fi-FI&quot;:&quot;PropertyControlStrings_fi-fi_6cf34edf8caf385bea84528cc386cd5f.js&quot;,&quot;fr-FR&quot;:&quot;PropertyControlStrings_fr-fr_40412f0ae3f35c004548887f7771f932.js&quot;,&quot;it-IT&quot;:&quot;PropertyControlStrings_it-it_41f18fa581d461db7c640679a277b3f3.js&quot;,&quot;ja-JP&quot;:&quot;PropertyControlStrings_ja-jp_744730f9cd57ff5fa842e19feee43689.js&quot;,&quot;lt-LT&quot;:&quot;PropertyControlStrings_lt-lt_dd9c4dc605f7cc52c865e75fbe69e815.js&quot;,&quot;lv-LV&quot;:&quot;PropertyControlStrings_lv-lv_094f83cf2fa3af52b75bffbb5901884f.js&quot;,&quot;nb-NO&quot;:&quot;PropertyControlStrings_nb-no_32f68fc0283e323486d606e50d93b415.js&quot;,&quot;nl-NL&quot;:&quot;PropertyControlStrings_nl-nl_a7ff2ae2a1b24528db15f91c4a101585.js&quot;,&quot;no&quot;:&quot;PropertyControlStrings_no_9010b4d755bda00769d9ef249eafb363.js&quot;,&quot;pl-PL&quot;:&quot;PropertyControlStrings_pl-pl_d93eb3e8701bd0c9a95d61257bb9c3c6.js&quot;,&quot;pt-PT&quot;:&quot;PropertyControlStrings_pt-pt_48a0c94520a06ac244dd8bcdd8dd1af3.js&quot;,&quot;ro-RO&quot;:&quot;PropertyControlStrings_ro-ro_d3c39ea618dd810b7e965e6be29b81ce.js&quot;,&quot;ru-RU&quot;:&quot;PropertyControlStrings_ru-ru_858b38db3deb4265f0204ee5e6a73cca.js&quot;,&quot;sk-SK&quot;:&quot;PropertyControlStrings_sk-sk_ecfbffa01d2e85ea82262ea25a0bf1d2.js&quot;,&quot;sr-Latn-RS&quot;:&quot;PropertyControlStrings_sr-latn-rs_8bfa79b5cdb880bab855577fef642f5b.js&quot;,&quot;sv-SE&quot;:&quot;PropertyControlStrings_sv-se_e437e277b30c73090976881d3ab3f71e.js&quot;,&quot;tr-TR&quot;:&quot;PropertyControlStrings_tr-tr_dd9b4d5310d9f61f6dbf497c5ac28472.js&quot;,&quot;vi-VN&quot;:&quot;PropertyControlStrings_vi-vn_07c1001da775bf2ce49d534eb2c13f7e.js&quot;,&quot;zh-CN&quot;:&quot;PropertyControlStrings_zh-cn_b6ff4b7c62d86fb9c2631b9ae98ff1b7.js&quot;,&quot;zh-TW&quot;:&quot;PropertyControlStrings_zh-tw_54b7510d208fa574ce536c9a8a82179e.js&quot;},&quot;defaultPath&quot;:&quot;PropertyControlStrings_en-us_c3915679a97436fd69e420b5b4064a25.js&quot;}}}}" Type="WebPart"></ClientSideComponent><Module Name="${validWebPartName}" Url="_catalogs/wp" List="113"></Module></Elements>`;
  const invalidWebPartName = 'Invalid WebPart';
  const invalidWebPartId = '2d6310ab-19be-453e-bad8-d8e648978b75';
  const invalidWebpartBody = `<?xml version="1.0" encoding="utf-8"?><Elements xmlns="http://schemas.microsoft.com/sharepoint/"><ClientSideComponent Name="${invalidWebPartName}" Id="${invalidWebPartId}" ComponentManifest="{&quot;id&quot;:&quot;${invalidWebPartId}&quot;,&quot;alias&quot;:&quot;${invalidWebPartName}&quot;,&quot;componentType&quot;:&quot;WebPart&quot;,&quot;version&quot;:&quot;0.0.1&quot;,&quot;manifestVersion&quot;:2,&quot;supportedHosts&quot;:[&quot;SharePointWebPart&quot;],&quot;supportsThemeVariants&quot;:true,&quot;preconfiguredEntries&quot;:[{&quot;groupId&quot;:&quot;5c03119e-3074-46fd-976b-c60198311f70&quot;,&quot;group&quot;:{&quot;default&quot;:&quot;Advanced&quot;},&quot;title&quot;:{&quot;default&quot;:&quot;${invalidWebPartName}&quot;},&quot;description&quot;:{&quot;default&quot;:&quot;${invalidWebPartName}&quot;},&quot;officeFabricIconFontName&quot;:&quot;TaskSolid&quot;,&quot;properties&quot;:{&quot;description&quot;:&quot;${invalidWebPartName}&quot;}}],&quot;loaderConfig&quot;:{&quot;internalModuleBaseUrls&quot;:[&quot;HTTPS://SPCLIENTSIDEASSETLIBRARY/&quot;],&quot;entryModuleId&quot;:&quot;to-do-minified-web-part&quot;,&quot;scriptResources&quot;:{&quot;to-do-minified-web-part&quot;:{&quot;type&quot;:&quot;path&quot;,&quot;path&quot;:&quot;to-do-minified-web-part_0f5d1255f33a416eb975.js&quot;},&quot;@microsoft/sp-core-library&quot;:{&quot;type&quot;:&quot;component&quot;,&quot;id&quot;:&quot;7263c7d0-1d6a-45ec-8d85-d4d1d234171b&quot;,&quot;version&quot;:&quot;1.16.1&quot;},&quot;@microsoft/sp-webpart-base&quot;:{&quot;type&quot;:&quot;component&quot;,&quot;id&quot;:&quot;974a7777-0990-4136-8fa6-95d80114c2e0&quot;,&quot;version&quot;:&quot;1.16.1&quot;},&quot;react&quot;:{&quot;type&quot;:&quot;component&quot;,&quot;id&quot;:&quot;0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d&quot;,&quot;version&quot;:&quot;17.0.1&quot;},&quot;react-dom&quot;:{&quot;type&quot;:&quot;component&quot;,&quot;id&quot;:&quot;aa0a46ec-1505-43cd-a44a-93f3a5aa460a&quot;,&quot;version&quot;:&quot;17.0.1&quot;},&quot;ToDoWebPartStrings&quot;:{&quot;type&quot;:&quot;path&quot;,&quot;path&quot;:&quot;ToDoWebPartStrings_en-us_d32717b12d14d01f33fa09b9f868b0bd.js&quot;}}}}" Type="WebPart"></ClientSideComponent><Module Name="${invalidWebPartName}" Url="_catalogs/wp" List="113"></Module></Elements>`;

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    commandInfo = Cli.getCommandInfo(command);
    (command as any).solutionZip = admZipMock;
    (command as any).fixZip = admZipMock;
    Cli.getInstance().config;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(fs, 'mkdtempSync').callsFake(_ => tmpDir);
    sinon.stub(path, 'resolve').returns(fullPath);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.mkdtempSync,
      fs.readdirSync,
      fs.readFileSync,
      fs.rmdirSync,
      fs.writeFileSync,
      path.resolve,
      admZipMock.extractAllTo
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PACKAGE_TEAMS_ENABLE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('successfully checks SPFx solution on Teams webparts without fixing', async () => {
    sinon.stub(fs, 'readdirSync').callsFake(args => {
      if (args === tmpDir) {
        return sppkgTopLevelContent as any;
      }

      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}`) {
        return webpartXmls as any;
      }
      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(args => {
      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}\\${webpartXmls[0]}`) {
        return validWebpartBody;
      }

      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}\\${webpartXmls[1]}`) {
        return invalidWebpartBody;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'rmdirSync').returns();

    await command.action(logger, { options: { filePath: filePath, verbose: true } });
    assert.strictEqual(loggerLogSpy.lastCall.args[0], chalk.red(`Webpart with id ${invalidWebPartId} and alias ${invalidWebPartName} is not set-up as a Teams app.`));
  });

  it('successfully checks SPFx solution on Teams webparts with fixing broken webparts while passing new hosts to set', async () => {
    sinon.stub(fs, 'readdirSync').callsFake(args => {
      if (args === tmpDir) {
        return sppkgTopLevelContent as any;
      }

      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}`) {
        return webpartXmls as any;
      }
      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(args => {
      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}\\${webpartXmls[0]}`) {
        return validWebpartBody;
      }

      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}\\${webpartXmls[1]}`) {
        return invalidWebpartBody;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'writeFileSync').returns();
    sinon.stub(fs, 'rmdirSync').returns();

    await command.action(logger, { options: { filePath: filePath, fix: true, supportedHost: 'TeamsPersonalApp,TeamsTab', verbose: true } });
    assert.strictEqual(loggerLogSpy.lastCall.args[0], `Time to fix the webpart to make it possible to set up as a Teams app.`);
  });

  it('successfully checks SPFx solution on Teams webparts with fixing broken webparts without passing new hosts to set', async () => {
    sinon.stub(fs, 'readdirSync').callsFake(args => {
      if (args === tmpDir) {
        return sppkgTopLevelContent as any;
      }

      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}`) {
        return webpartXmls as any;
      }
      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(args => {
      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}\\${webpartXmls[0]}`) {
        return validWebpartBody;
      }

      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}\\${webpartXmls[1]}`) {
        return invalidWebpartBody;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'writeFileSync').returns();
    sinon.stub(fs, 'rmdirSync').returns();

    await command.action(logger, { options: { filePath: filePath, fix: true, verbose: true } });
    assert.strictEqual(loggerLogSpy.lastCall.args[0], `Time to fix the webpart to make it possible to set up as a Teams app.`);
  });

  it('throws an error when saving the new sppkg version when opened in archive manager', async () => {
    sinon.stub(fs, 'readdirSync').callsFake(args => {
      if (args === tmpDir) {
        return sppkgTopLevelContent as any;
      }

      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}`) {
        return webpartXmls as any;
      }
      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(args => {
      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}\\${webpartXmls[0]}`) {
        return validWebpartBody;
      }

      if (args === `${tmpDir}\\${sppkgTopLevelContent[0]}\\${webpartXmls[1]}`) {
        return invalidWebpartBody;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'writeFileSync').callsFake(path => {
      if (path === `${tmpDir}\\${sppkgTopLevelContent[0]}\\${webpartXmls[1]}`) {
        return;
      }

      if (path === fullPath) {
        throw `EBUSY: resource busy or locked, open '${fullPath}'`;
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'rmdirSync').returns();

    await assert.rejects(command.action(logger, { options: { filePath: filePath, fix: true, verbose: true } }), new CommandError(`EBUSY: resource busy or locked, open '${fullPath}'`));
  });

  it('fails validation if fullPath to sppkg does not exist', async () => {
    sinonUtil.restore(fs.existsSync);

    sinon.stub(fs, 'existsSync').callsFake(() => false);

    const actual = await command.validate({ options: { filePath: filePath } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if fullPath does not end with .sppkg', async () => {
    sinonUtil.restore(path.resolve);

    const failingPath = fullPath.replace('.sppkg', '.zip');
    sinon.stub(path, 'resolve').returns(failingPath);

    const actual = await command.validate({ options: { filePath: filePath } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if supportedHost is not valid', async () => {
    const actual = await command.validate({ options: { filePath: filePath, fix: true, supportedHost: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if filePath to sppkg exists', async () => {
    const actual = await command.validate({ options: { filePath: filePath } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if filePath to sppkg exists with a supported host', async () => {
    const actual = await command.validate({ options: { filePath: filePath, fix: true, supportedHost: supportedHost } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
