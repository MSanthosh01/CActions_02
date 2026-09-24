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

  // Wait for DOM content to be loaded
  await webCtx.page.waitForLoadState('domcontentloaded');
  ctx.log('DOM content loaded');

  // Wait for full page load (network idle)
  await webCtx.page.waitForLoadState('load');
  ctx.log('Page fully loaded');

  // Wait for the target element to be attached and visible before scrolling
  await webCtx.waitForAttached(selector);
  await webCtx.waitForVisible(selector);

  await webCtx.scroll({ selector });
  ctx.log(`Scrolled to element: ${selector}`);
}
