import { ZodTypeAny, z } from 'zod';
import { CommandOptionInfo } from '../cli/CommandOptionInfo';
import { CommandOption } from '../Command';

function parseEffect(def: z.ZodEffectsDef, _options: CommandOptionInfo[], _currentOption?: CommandOptionInfo): z.ZodTypeDef | undefined {
  return def.schema._def;
}

function parseIntersection(def: z.ZodIntersectionDef, _options: CommandOptionInfo[], _currentOption?: CommandOptionInfo): z.ZodTypeDef | undefined {
  if (def.left._def.typeName !== z.ZodFirstPartyTypeKind.ZodAny) {
    return def.left._def;
  }

  if (def.right._def.typeName !== z.ZodFirstPartyTypeKind.ZodAny) {
    return def.right._def;
  }

  return;
}

function parseObject(def: z.ZodObjectDef, options: CommandOptionInfo[], _currentOption?: CommandOptionInfo): z.ZodTypeDef | undefined {
  const properties = def.shape();
  for (const key in properties) {
    const property = properties[key];

    const option: CommandOptionInfo = {
      name: key,
      long: key,
      short: property._def.alias,
      required: true,
      type: 'string'
    };

    parseDef(property._def, options, option);
    options.push(option);
  }

  return;
}

function parseString(_def: z.ZodStringDef, _options: CommandOptionInfo[], currentOption?: CommandOptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.type = 'string';
  }

  return;
}

function parseNumber(_def: z.ZodNumberDef, _options: CommandOptionInfo[], currentOption?: CommandOptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.type = 'number';
  }

  return;
}

function parseBoolean(_def: z.ZodBooleanDef, _options: CommandOptionInfo[], currentOption?: CommandOptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.type = 'boolean';
  }

  return;
}

function parseOptional(def: z.ZodOptionalDef, _options: CommandOptionInfo[], currentOption?: CommandOptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.required = false;
  }

  return def.innerType._def;
}

function parseDefault(def: z.ZodDefaultDef, _options: CommandOptionInfo[], currentOption?: CommandOptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.required = false;
  }

  return def.innerType._def;
}

function parseEnum(def: z.ZodEnumDef, _options: CommandOptionInfo[], currentOption?: CommandOptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.type = 'string';
    currentOption.autocomplete = [...def.values];
  }

  return;
}

function parseNativeEnum(def: z.ZodNativeEnumDef, _options: CommandOptionInfo[], currentOption?: CommandOptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.type = 'string';
    currentOption.autocomplete = Object.values(def.values).map(v => String(v));
  }

  return;
}

function getParseFn(typeName: z.ZodFirstPartyTypeKind): undefined | ((def: any, options: CommandOptionInfo[], currentOption?: CommandOptionInfo) => z.ZodTypeDef | undefined) {
  switch (typeName) {
    case z.ZodFirstPartyTypeKind.ZodEffects:
      return parseEffect;
    case z.ZodFirstPartyTypeKind.ZodObject:
      return parseObject;
    case z.ZodFirstPartyTypeKind.ZodOptional:
      return parseOptional;
    case z.ZodFirstPartyTypeKind.ZodString:
      return parseString;
    case z.ZodFirstPartyTypeKind.ZodNumber:
      return parseNumber;
    case z.ZodFirstPartyTypeKind.ZodBoolean:
      return parseBoolean;
    case z.ZodFirstPartyTypeKind.ZodEnum:
      return parseEnum;
    case z.ZodFirstPartyTypeKind.ZodNativeEnum:
      return parseNativeEnum;
    case z.ZodFirstPartyTypeKind.ZodDefault:
      return parseDefault;
    case z.ZodFirstPartyTypeKind.ZodIntersection:
      return parseIntersection;
    default:
      return;
  }
}

function parseDef(def: z.ZodTypeDef, options: CommandOptionInfo[], currentOption?: CommandOptionInfo): void {
  let parsedDef: z.ZodTypeDef | undefined = def;

  do {
    const parse = getParseFn((parsedDef as any).typeName);
    if (!parse) {
      break;
    }

    parsedDef = parse(parsedDef as any, options, currentOption);
    if (!parsedDef) {
      break;
    }

  } while (parsedDef);
}

function optionToString(optionInfo: CommandOptionInfo): string {
  let s = '';

  if (optionInfo.short) {
    s += `-${optionInfo.short}, `;
  }

  s += `--${optionInfo.long}`;

  if (optionInfo.type !== 'boolean') {
    s += ' ';
    s += optionInfo.required ? '<' : '[';
    s += optionInfo.long;
    s += optionInfo.required ? '>' : ']';
  }

  return s;
};

type EnumLike = {
  [k: string]: string | number
  [nu: number]: string
};

export const zod = {
  alias<T extends ZodTypeAny>(alias: string, type: T): T {
    type._def.alias = alias;
    return type;
  },

  schemaToOptionInfo(schema: z.ZodSchema<any>): CommandOptionInfo[] {
    const options: CommandOptionInfo[] = [];
    parseDef(schema._def, options);
    return options;
  },

  schemaToOptions(schema: z.ZodSchema<any>): CommandOption[] {
    const optionsInfo: CommandOptionInfo[] = this.schemaToOptionInfo(schema);
    const options: CommandOption[] = optionsInfo.map(option => {
      return {
        option: optionToString(option),
        autocomplete: option.autocomplete
      };
    });
    return options;
  },

  coercedEnum: <T extends EnumLike>(e: T): z.ZodEffects<z.ZodNativeEnum<T>, T[keyof T], unknown> =>
    z.preprocess(val => {
      const target = String(val)?.toLowerCase();
      for (const k of Object.values(e)) {
        if (String(k)?.toLowerCase() === target) {
          return k;
        }
      }

      return null;
    }, z.nativeEnum(e))
};