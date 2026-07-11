import type { WalnutContext, WalnutWebContext } from './walnut';

/** @walnut_method
 * name: custom Click Element
 * description: Click on element using selector ${selector}
 * actionType: custom_click_element
 * context: web
 * needsLocator: false
 * category: Element Interaction
 */
export async function clickElement(ctx: WalnutContext) {
  // ctx.args[0] = selector value (from ${selector})
  const webCtx = ctx as WalnutWebContext;
  const selector = ctx.args[0];
  await webCtx.click(selector);
  ctx.log(`Clicked element: ${selector}`);
}
