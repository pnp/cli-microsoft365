import { z } from 'zod';
import { JSONSchema } from 'zod/v4/core';
import { EnumLike } from 'zod/v4/core/util.cjs';
import { CommandOptionInfo } from '../cli/CommandOptionInfo';
import { CommandOption } from '../Command';

declare module 'zod' {
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  interface ZodType<out Output = unknown, out Input = unknown, out Internals extends z.core.$ZodTypeInternals<Output, Input> = z.core.$ZodTypeInternals<Output, Input>> {
    alias(name: string): this & { alias?: string };
  }
}

z.ZodType.prototype.alias = function (name: string) {
  (this.def as any).alias = name;
  return this;
};

function parseObject(schema: JSONSchema.JSONSchema, options: CommandOptionInfo[], _currentOption?: CommandOptionInfo): JSONSchema.JSONSchema | undefined {
  for (const key in schema.properties) {
    const property = schema.properties[key];

    const option: CommandOptionInfo = {
      name: key,
      long: key,
      short: (property as any)['x-alias'],
      required: schema.required?.includes(key) && (property as any).default === undefined || false,
      type: 'string'
    };

    parseJSONSchema(property as JSONSchema.JSONSchema, options, option);
    options.push(option);
  }

  return;
}

function parseString(schema: JSONSchema.JSONSchema, _options: CommandOptionInfo[], currentOption?: CommandOptionInfo): JSONSchema.JSONSchema | undefined {
  if (currentOption) {
    currentOption.type = 'string';

    if (schema.enum) {
      currentOption.autocomplete = schema.enum.map(e => String(e));
    }
  }

  return;
}

function parseNumber(schema: JSONSchema.JSONSchema, _options: CommandOptionInfo[], currentOption?: CommandOptionInfo): JSONSchema.JSONSchema | undefined {
  if (currentOption) {
    currentOption.type = 'number';
  }

  return;
}

function parseBoolean(schema: JSONSchema.JSONSchema, _options: CommandOptionInfo[], currentOption?: CommandOptionInfo): JSONSchema.JSONSchema | undefined {
  if (currentOption) {
    currentOption.type = 'boolean';
  }

  return;
}

function getParseFn(typeName?: "object" | "array" | "string" | "number" | "boolean" | "null" | "integer"): undefined | ((schema: JSONSchema.JSONSchema, options: CommandOptionInfo[], currentOption?: CommandOptionInfo) => JSONSchema.JSONSchema | undefined) {
  switch (typeName) {
    case 'object':
      return parseObject;
    case 'string':
      return parseString;
    case 'number':
    case 'integer':
      return parseNumber;
    case 'boolean':
      return parseBoolean;
    default:
      return;
  }
}

function parseJSONSchema(jsonSchema: JSONSchema.JSONSchema, options: CommandOptionInfo[], currentOption?: CommandOptionInfo): void {
  let parsedSchema: JSONSchema.JSONSchema | undefined = jsonSchema;

  do {
    if (parsedSchema.allOf) {
      parsedSchema.allOf.forEach(s => parseJSONSchema(s, options, currentOption));
    }

    const parse = getParseFn(parsedSchema.type);
    if (!parse) {
      break;
    }

    parsedSchema = parse(parsedSchema, options, currentOption);
    if (!parsedSchema) {
      break;
    }

  } while (parsedSchema);
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

export const zod = {
  schemaToOptionInfo(schema: z.ZodSchema<any>): CommandOptionInfo[] {
    const jsonSchema = z.toJSONSchema(schema, {
      override: s => {
        const alias = ((s.zodSchema as unknown as z.core.$ZodTypeInternals).def as any).alias;
        if (alias) {
          s.jsonSchema['x-alias'] = alias;
        }
      },
      unrepresentable: 'any'
    });

    const options: CommandOptionInfo[] = [];
    parseJSONSchema(jsonSchema, options);
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

  coercedEnum: <T extends EnumLike>(e: T): z.ZodPipe<z.ZodTransform<string | number | null, unknown>, z.ZodEnum<T>> =>
    z.preprocess(val => {
      const target = String(val)?.toLowerCase();
      for (const k of Object.values(e)) {
        if (String(k)?.toLowerCase() === target) {
          return k;
        }
      }

      return null;
    }, z.enum(e))
};