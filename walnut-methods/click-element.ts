import type { WalnutContext, WalnutWebContext } from './walnut';

/** @walnut_method
 * name: Click Element
 * description: Click on element using selector ${selector}
 * actionType: custom_click_element
 * context: web
 * needsLocator: true
 * category: Element Interaction
 */
export async function clickElement(ctx: WalnutContext) {
  // needsLocator: true — runtime injects ctx.locator with the resolved ${selector} value
  const webCtx = ctx as WalnutWebContext & { locator: string };
  await webCtx.click(webCtx.locator);
  ctx.log(`Clicked element: ${webCtx.locator}`);
}
