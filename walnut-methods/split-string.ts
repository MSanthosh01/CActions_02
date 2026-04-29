import type { WalnutContext } from './walnut';

/** @walnut_method
 * name: Split String by Delimiter
 * description: Split ${inputString} by delimiter ${delimiter} and store value at index ${index} in $[result]
 * actionType: custom_split_string
 * context: shared
 * needsLocator: false
 * category: Data Processing
 */
export async function splitStringandStore(ctx: WalnutContext) {
  // ctx.args[0] = inputString value  (from ${inputString})
  // ctx.args[1] = delimiter value    (from ${delimiter})
  // ctx.args[2] = index value        (from ${index})
  // ctx.args[3] = "result"           (from $[result])

  const inputString = ctx.args[0];
  const delimiter   = ctx.args[1];
  const index       = parseInt(ctx.args[2], 10);
  const outputVar   = ctx.args[3];

  if (!inputString) {
    throw new Error('inputString is empty or undefined');
  }

  const parts = inputString.split(delimiter);

  if (isNaN(index)) {
    throw new Error(`Invalid index: "${ctx.args[2]}" is not a number`);
  }

  if (index < 0 || index >= parts.length) {
    throw new Error(
      `Index ${index} is out of range. The split produced ${parts.length} part(s) (valid indices: 0–${parts.length - 1})`
    );
  }

  const value = parts[index];
  ctx.setVariable(outputVar, value);
  ctx.log(`Split "${inputString}" by "${delimiter}" → part[${index}] = "${value}" stored in $[${outputVar}]`);
}
