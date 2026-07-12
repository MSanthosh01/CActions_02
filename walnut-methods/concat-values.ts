import type { WalnutContext } from './walnut';

/** @walnut_method
 * name: Concatenate Values
 * description: Concatenate ${value1} and ${value2} and ${value3} with separator ${separator} and store in $[result]
 * actionType: custom_concat_values
 * context: shared
 * needsLocator: false
 * category: Data Processing
 */
export async function concatValues(ctx: WalnutContext) {
  // Args are always read from the TAIL so the value count is flexible:
  //   args[last]       = output variable name  (from $[result])
  //   args[last - 1]   = separator             (from ${separator})
  //   args[0 .. n-2]   = values to concatenate (local params OR runtime variable names)
  //
  // This means the user can write any number of ${valueN} placeholders in the description
  // before ${separator} and $[result] without breaking the index mapping.

  const totalArgs = ctx.args.length;

  if (totalArgs < 3) {
    throw new Error(
      'At least 3 args are required: one value, a separator, and an output variable name.'
    );
  }

  const outputVar  = ctx.args[totalArgs - 1];       // last   → $[result]
  const separator  = ctx.args[totalArgs - 2] ?? '';  // last-1 → ${separator}
  const valueArgs  = ctx.args.slice(0, totalArgs - 2); // everything before → values

  if (!outputVar) {
    throw new Error('Output variable name ($[result]) is missing or empty.');
  }

  /**
   * Resolve one arg:
   *  - If a runtime variable with that name exists → return its stored value (any type → string).
   *  - Otherwise → use the raw arg string as a literal value.
   *  - Empty / undefined args are skipped (returns null).
   */
  function resolve(arg: string | undefined): string | null {
    if (arg === undefined || arg === null || arg === '') return null;
    const stored = ctx.getVariable(arg);
    if (stored !== undefined && stored !== null) {
      return String(stored);
    }
    return String(arg);
  }

  const parts: string[] = [];
  for (const arg of valueArgs) {
    const resolved = resolve(arg);
    if (resolved !== null) {
      parts.push(resolved);
    }
  }

  if (parts.length === 0) {
    throw new Error('No values to concatenate — all value slots were empty.');
  }

  const concatenated = parts.join(separator);
  ctx.setVariable(outputVar, concatenated);
  ctx.log(
    `Concatenated ${parts.length} value(s) with separator "${separator}" → "${concatenated}" stored in $[${outputVar}]`
  );
}
