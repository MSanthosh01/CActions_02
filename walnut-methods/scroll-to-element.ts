import type { WalnutContext, WalnutWebContext } from './walnut';

/** @walnut_method
 * name: custom Scroll to Element
 * description: Scroll to element using selector ${selector}
 * actionType: custom_scroll_to_element
 * context: web
 * needsLocator: false
 * category: Element Interaction
 */
export async function scrollToElement(ctx: WalnutContext) {
  // ctx.args[0] = selector value (from ${selector})
  const webCtx = ctx as WalnutWebContext;
  const selector = ctx.args[0];
  await webCtx.scroll({ selector });
  ctx.log(`Scrolled to element: ${selector}`);
}
