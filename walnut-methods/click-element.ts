import type { WalnutWebContext } from './walnut';

/** @walnut_method
 * name: custom Click
 * description: Click the element linked to this step
 * actionType: custom_click
 * context: web
 * needsLocator: true
 * category: Element Interaction
 */
export async function customClick(ctx: WalnutWebContext) {
  // The linked step object is injected as a Playwright Locator.
  const webCtx = ctx as WalnutWebContext & { locator: { click(): Promise<void> } };
  await webCtx.locator.click();
}
