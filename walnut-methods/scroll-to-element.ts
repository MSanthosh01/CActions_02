import type { WalnutContext, WalnutWebContext } from './walnut';

/** @walnut_method
 * name: Scroll to Element
 * description: Scroll to element using selector ${selector}
 * actionType: custom_scroll_to_element
 * context: web
 * needsLocator: true
 * category: Element Interaction
 */
export async function scrollToElement(ctx: WalnutContext) {
  // needsLocator: true — runtime injects ctx.locator with the resolved ${selector} value
  const webCtx = ctx as WalnutWebContext & { locator: string };
  await webCtx.scroll({ selector: webCtx.locator });
  ctx.log(`Scrolled to element: ${webCtx.locator}`);
}
