import type { WalnutContext } from './walnut';

/** @walnut_method
 * name: Concatenate Values
 * description: Concatenate ${value1} and ${value2} and ${value3} and ${value4} and ${value5} with separator ${separator} and store in $[result]
 * actionType: custom_concat_values
 * context: shared
 * needsLocator: false
 * category: Data Processing
 */
export async function concatValues(ctx: WalnutContext) {
  // ctx.args layout (all 7 placeholders in description order):
  //   args[0] = value1   — direct local param OR runtime variable name (from ${value1})
  //   args[1] = value2   — direct local param OR runtime variable name (from ${value2})
  //   args[2] = value3   — direct local param OR runtime variable name (from ${value3})
  //   args[3] = value4   — direct local param OR runtime variable name (from ${value4})
  //   args[4] = value5   — direct local param OR runtime variable name (from ${value5})
  //   args[5] = separator — string to place between parts, e.g. " ", ",", "" (from ${separator})
  //   args[6] = "result"  — runtime variable name to store the concatenated output (from $[result])
  //
  // Smart resolution: for each value slot, if the arg matches a key in the variable context
  // (i.e. a runtime variable exists with that name), its stored value is used.
  // Otherwise, the raw arg string is used as-is. This lets you pass either:
  //   - a static/local param value directly  → ${myParam}   = "Hello"
  //   - a runtime variable by name           → ${myVar}     = "tokenValue"  (resolved transparently)
  //   - a $[varName] reference               → $[myVar]     = "myVar" string, then getVariable("myVar")

  const separator = ctx.args[5] ?? '';
  const outputVar = ctx.args[6];

  if (!outputVar) {
    throw new Error('Output variable name (last placeholder $[result]) is required.');
  }

  /**
   * Resolve a single arg: if a runtime variable with that name exists,
   * return its value (coerced to string). Otherwise return the raw arg string.
   * Skips empty/undefined args so unused slots are ignored.
   */
  function resolve(arg: string | undefined): string | null {
    if (arg === undefined || arg === null || arg === '') return null;
    const stored = ctx.getVariable(arg);
    if (stored !== undefined && stored !== null) {
      // Runtime variable found — use its value
      return String(stored);
    }
    // No matching runtime variable — use the raw arg as a literal value
    return String(arg);
  }

  const parts: string[] = [];
  for (let i = 0; i <= 4; i++) {
    const resolved = resolve(ctx.args[i]);
    if (resolved !== null) {
      parts.push(resolved);
    }
  }

  if (parts.length === 0) {
    throw new Error('No values provided to concatenate. Supply at least one non-empty value.');
  }

  const concatenated = parts.join(separator);
  ctx.setVariable(outputVar, concatenated);
  ctx.log(`Concatenated ${parts.length} value(s) with separator "${separator}" → "${concatenated}" stored in $[${outputVar}]`);
}
