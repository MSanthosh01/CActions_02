import type { WalnutContext, WalnutWebContext } from './walnut';

/** @walnut_method
 * name: custom Clear Text
 * description: Clear the text content of ${element}
 * actionType: custom_clear_text
 * context: web
 * needsLocator: true
 * category: Forms
 */
export async function clearText(ctx: WalnutContext) {
  // Cast to access both ctx.locator (runtime-injected selector) and ctx.clear()
  // ctx.locator is injected at runtime when needsLocator: true — not yet typed in walnut.d.ts
  const webCtx = ctx as WalnutWebContext & { locator: string };
  await webCtx.clear(webCtx.locator);
}
