import type { WalnutContext } from './walnut';

/** @walnut_method
 * name: custom Generate Random Number
 * description: Generate a random numeric string of length ${length} and store in $[result]
 * actionType: custom_random_number
 * context: shared
 * needsLocator: false
 * category: Data Generation
 */
export async function generateRandomNumber(ctx: WalnutContext) {
  // ctx.args[0] = length value   (from ${length})
  // ctx.args[1] = "result"        (from $[result]) — runtime variable name to store output

  const length   = parseInt(ctx.args[0], 10);
  const outputVar = ctx.args[1];

  if (isNaN(length) || length <= 0) {
    throw new Error(`Invalid length: "${ctx.args[0]}" must be a positive number`);
  }

  let result = '';
  for (let i = 0; i < length; i++) {
    result += Math.floor(Math.random() * 10).toString();
  }

  ctx.setVariable(outputVar, result);
  ctx.log(`Generated random number (length ${length}) stored in $[${outputVar}]`);
}
