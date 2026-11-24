/**
 * Properties schema for this Copilot Component.
 *
 * This schema is defined with Zod and exported as JSON Schema via
 * Zod's built-in converter. The manifest references the compiled `.js` default
 * export, which the Copilot host uses to validate and describe the tool
 * arguments that Copilot passes when invoking this component.
 *
 * To add more properties, extend the `z.object({...})` below — they will
 * automatically appear as tool parameters in the Copilot UI.
 */
import { z } from 'zod';

const propertiesSchema = z.object({
  message: z.string().describe('A message to display.')
});

export type IHelloWorldCopilotComponentProperties = z.infer<typeof propertiesSchema>;

export default z.toJSONSchema(propertiesSchema, { target: 'draft-7' });
