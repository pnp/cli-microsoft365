#!/usr/bin/env node

const omelette: (template: string) => Omelette = require('omelette');
import * as os from 'os';
import * as fs from 'fs';
import * as path from 'path';
import { Cli } from './cli/Cli';
import { CommandInfo } from './cli/CommandInfo';
import { CommandOptionInfo } from './cli/CommandOptionInfo';

class Autocomplete {
  private static autocompleteFilePath: string = path.join(__dirname, `..${path.sep}commands.json`);
  private omelette!: Omelette;
  private commands: any = {};

  constructor() {
    this.init();
  }

  private init(): void {
    if (fs.existsSync(Autocomplete.autocompleteFilePath)) {
      try {
        const data: string = fs.readFileSync(Autocomplete.autocompleteFilePath, 'utf-8');
        this.commands = JSON.parse(data);
      }
      catch { }
    }

    this.omelette = omelette('m365_comp|m365|microsoft365');
    this.omelette.on('complete', this.handleAutocomplete.bind(this));
    this.omelette.init();
  }

  public handleAutocomplete(fragment: string, data: EventData): void {
    let replies: Object | string[] = {};
    let allWords: string[] = [];

    if (data.fragment === 1) {
      replies = Object.keys(this.commands);
    }
    else {
      allWords = data.line.split(/\s+/).slice(1, -1);
      // build array of words to use as a path to retrieve completion
      // options from the commands tree
      const words: string[] = allWords
        .filter((e: string, i: number): boolean => {
          if (e.indexOf('-') !== 0) {
            // if the word is not an option check if it's not
            // option's value, eg. --output json, in which case
            // the suggestion should be command options
            return i === 0 || allWords[i - 1].indexOf('-') !== 0;
          }
          else {
            // remove all options but last one
            return i === allWords.length - 1;
          }
        });
      let accessor: Function = new Function('_', "return _['" + (words.join("']['")) + "']");

      try {
        replies = accessor(this.commands);
        // if the last word is an option without autocomplete
        // suggest other options from the same command
        if (words[words.length - 1].indexOf('-') === 0 &&
          !Array.isArray(replies)) {
          accessor = new Function('_', "return _['" + (words.filter(w => w.indexOf('-') !== 0).join("']['")) + "']");
          replies = accessor(this.commands);
          replies = Object.keys(replies);
        }
      }
      catch { }
    }

    if (!replies) {
      replies = [];
    }

    if (!Array.isArray(replies)) {
      replies = Object.keys(replies);
    }

    // remove options that already have been used
    replies = (replies as string[]).filter(r => r.indexOf('-') !== 0 || allWords.indexOf(r) === -1);

    data.reply(replies);
  }

  public generateShCompletion(): void {
    const cli: Cli = Cli.getInstance();
    const commandsInfo: any = this.getCommandsInfo(cli);
    fs.writeFileSync(Autocomplete.autocompleteFilePath, JSON.stringify(commandsInfo));
  }

  public setupShCompletion(): void {
    this.omelette.setupShellInitFile();
  }

  public getClinkCompletion(): string {
    const cli: Cli = Cli.getInstance();
    const cmd: any = this.getCommandsInfo(cli);
    const lua: string[] = ['local parser = clink.arg.new_parser'];
    const functions: any = {};

    this.buildClinkForBranch(cmd, functions, 'm365');

    Object.keys(functions).forEach(k => {
      functions[k] = functions[k].replace(/#([^#]+)#/g, (m: string, p1: string): string => functions[p1]);
    });

    lua.push(
      'local m365_parser = ' + functions['m365'],
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    );

    return lua.join(os.EOL);
  }

  private buildClinkForBranch(branch: any, functions: any, luaFunctionName: string): void {
    if (!Array.isArray(branch)) {
      const keys: string[] = Object.keys(branch);

      keys.forEach(k => {
        if (Object.keys(branch[k]).length > 0) {
          this.buildClinkForBranch(branch[k], functions, this.getLuaFunctionName(`${luaFunctionName}_${k}`));
        }
      });
    }

    const parser: string[] = [];

    parser.push(
      `parser({`
    );

    let printingArgs: boolean = false;

    if (Array.isArray(branch)) {
      branch.sort().forEach((c, i) => {
        const separator = i < branch.length - 1 ? ',' : '';
        parser.push(`"${c}"${separator}`);
      });
    }
    else {
      const keys = Object.keys(branch);
      if (keys.find(c => c.indexOf('-') === 0)) {
        printingArgs = true;
        const tmp: string[] = [];
        keys.sort().forEach((k, i) => {
          if (Object.keys(branch[k]).length > 0) {
            tmp.push(`"${k}"..#${this.getLuaFunctionName(`${luaFunctionName}_${k}`)}#`);
          }
          else {
            tmp.push(`"${k}"`);
          }
        });

        parser.push(`},${tmp.join(', ')}`);
      }
      else {
        keys.sort().forEach((k, i) => {
          const separator = i < keys.length - 1 ? ',' : '';
          parser.push(`"${k}"..#${this.getLuaFunctionName(`${luaFunctionName}_${k}`)}#${separator}`);
        });
      }
    }

    parser.push(`${printingArgs ? '' : '}'})`);
    functions[luaFunctionName] = parser.join('');
  }

  private getLuaFunctionName(functionName: string): string {
    return functionName.replace(/-/g, '_');
  }

  private getCommandsInfo(cli: Cli): any {
    const commandsInfo: any = {};
    const commands: CommandInfo[] = cli.commands;
    commands.forEach(c => {
      Autocomplete.processCommand(c.name, c, commandsInfo);
      if (c.aliases) {
        c.aliases.forEach(a => Autocomplete.processCommand(a, c, commandsInfo));
      }
    });

    return commandsInfo;
  }

  private static processCommand(commandName: string, commandInfo: CommandInfo, autocomplete: any) {
    const chunks: string[] = commandName.split(' ');
    let parent: any = autocomplete;
    for (let i: number = 0; i < chunks.length; i++) {
      const current: any = chunks[i];

      if (!parent[current]) {
        if (i < chunks.length - 1) {
          parent[current] = {};
        }
        else {
          // last chunk, add options
          const optionsArr: string[] = commandInfo.options
            .map(o => o.short)
            .concat(commandInfo.options.map(o => o.long))
            .filter(o => o != null)
            .map(o => (o as string).length === 1 ? `-${o}` : `--${o}`);
          optionsArr.push('--help');
          optionsArr.push('-h');
          const optionsObj: any = {};
          optionsArr.forEach(o => {
            const optionName: string = o.replace(/^-+/, '');
            const option: CommandOptionInfo = commandInfo.options.filter(opt => opt.long === optionName || opt.short === optionName)[0];
            if (option && option.autocomplete) {
              optionsObj[o] = option.autocomplete;
            }
            else {
              optionsObj[o] = {};
            }
          });
          parent[current] = optionsObj;
        }
      }

      parent = parent[current];
    }
  }
}

export const autocomplete = new Autocomplete();