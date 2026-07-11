import type { WalnutContext } from './walnut';

/** @walnut_method
 * name: custom Generate Random Alphabetic
 * description: Generate a random alphabetic string of length ${length} and store in $[result]
 * actionType: custom_random_alpha
 * context: shared
 * needsLocator: false
 * category: Data Generation
 */
export async function generateRandomAlpha(ctx: WalnutContext) {
  // ctx.args[0] = length value   (from ${length})
  // ctx.args[1] = "result"        (from $[result]) — runtime variable name to store output

  const length   = parseInt(ctx.args[0], 10);
  const outputVar = ctx.args[1];

  if (isNaN(length) || length <= 0) {
    throw new Error(`Invalid length: "${ctx.args[0]}" must be a positive number`);
  }

  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }

  ctx.setVariable(outputVar, result);
  ctx.log(`Generated random alphabetic string (length ${length}) stored in $[${outputVar}]`);
}
