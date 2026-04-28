import type { WalnutContext, WalnutWebContext } from './walnut';

/** @walnut_method
 * name: Clear Text
 * description: Clear the text content of the target element
 * actionType: custom_clear_text
 * context: web
 * needsLocator: true
 * category: Forms
 */
export async function clearText(ctx: WalnutContext) {
  // Cast to WalnutWebContext — this method is context: web only
  // ctx.args[0] carries the step's target element selector when needsLocator: true
  await (ctx as WalnutWebContext).clear(ctx.args[0]);
}
