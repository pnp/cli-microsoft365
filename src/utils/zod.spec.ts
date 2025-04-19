import assert from 'assert';
import { z } from 'zod';
import { zod } from '../utils/zod.js';

describe('utils/zod', () => {
  it('parses string option', () => {
    const schema = z.object({
      stringOption: z.string()
    }).strict();
    const options = zod.schemaToOptionInfo(schema);
    assert.strictEqual(options[0].type, 'string');
  });

  it('parses enum option', () => {
    const schema = z.object({
      enumOption: z.enum(['a', 'b', 'c'])
    }).strict();
    const options = zod.schemaToOptionInfo(schema);
    assert.deepStrictEqual(options[0].autocomplete, ['a', 'b', 'c']);
  });

  it('parses native enum option', () => {
    enum TestEnum {
      A = 'A',
      B = 'B',
      C = 'C'
    }
    const schema = z.object({
      enumOption: z.nativeEnum(TestEnum)
    }).strict();
    const options = zod.schemaToOptionInfo(schema);
    assert.deepStrictEqual(options[0].autocomplete, ['A', 'B', 'C']);
  });

  it('parses boolean option', () => {
    const schema = z.object({
      booleanOption: z.boolean()
    }).strict();
    const options = zod.schemaToOptionInfo(schema);
    assert.strictEqual(options[0].type, 'boolean');
  });

  it('parses number option', () => {
    const schema = z.object({
      numberOption: z.number()
    }).strict();
    const options = zod.schemaToOptionInfo(schema);
    assert.strictEqual(options[0].type, 'number');
  });

  it('parses required string option', () => {
    const schema = z.object({
      stringOption: z.string()
    }).strict();
    const options = zod.schemaToOptionInfo(schema);
    assert.strictEqual(options[0].required, true);
  });

  it('parses optional string option', () => {
    const schema = z.object({
      stringOption: z.string().optional()
    }).strict();
    const options = zod.schemaToOptionInfo(schema);
    assert.strictEqual(options[0].required, false);
  });

  it('parses optional boolean option with a default value', () => {
    const schema = z.object({
      boolOption: z.boolean().default(false)
    }).strict();
    const options = zod.schemaToOptionInfo(schema);
    assert.strictEqual(options[0].required, false);
  });

  it('parses refined schema', () => {
    const schema = z.object({
      boolOption: z.boolean().default(false)
    })
      .strict()
      .refine(data => data.boolOption === true, {
        message: 'boolOption must be true'
      });
    const options = zod.schemaToOptionInfo(schema);
    assert.strictEqual(options[0].name, 'boolOption');
  });

  it('parses schema that supports undefined options', () => {
    const schema = z.object({
      boolOption: z.boolean().default(false)
    })
      .and(z.any());
    const options = zod.schemaToOptionInfo(schema);
    assert.strictEqual(options[0].name, 'boolOption');
  });

  it('ignores unknown properties', () => {
    const schema = z.object({
      boolOption: z.boolean().nullable()
    });
    const options = zod.schemaToOptionInfo(schema);
    assert.strictEqual(options[0].name, 'boolOption');
  });

  it('parses extended schema', () => {
    const schema = z.object({
      boolOption: z.boolean().default(false)
    })
      .extend({
        stringOption: z.string()
      });
    const options = zod.schemaToOptionInfo(schema);
    assert.strictEqual(options[1].name, 'stringOption');
  });

  it('parses reversed schema', () => {
    const schema = z.any().and(z.object({
      boolOption: z.boolean().default(false)
    }));
    const options = zod.schemaToOptionInfo(schema);
    assert.strictEqual(options[0].name, 'boolOption');
  });

  it('parses noop', () => {
    const schema = z.any().and(z.any());
    zod.schemaToOptionInfo(schema);
    assert(true);
  });
});