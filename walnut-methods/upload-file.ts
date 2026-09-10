import type { WalnutContext, WalnutWebContext } from './walnut';
import * as fs from 'fs';

/** @walnut_method
 * name: Upload File to Element
 * description: Upload a file from local path or artifact id ${fileInput} to the linked element
 * actionType: custom_upload_file
 * context: web
 * needsLocator: true
 * category: Forms
 */
export async function uploadFile(ctx: WalnutContext) {
  // Object reference from the step — injected at runtime when needsLocator: true.
  // It is a Playwright Locator object, not a string. The linked object is often a
  // styled button (<div>) rather than an <input type="file">, so we use the file
  // chooser mechanism (same as the built-in upload step) instead of setInputFiles.
  const webCtx = ctx as WalnutWebContext & { locator: any };
  const locator = webCtx.locator;

  // User input: a local file path OR an artifact id
  const fileInput = ctx.args[0];

  let filePath = fileInput;

  if (!fs.existsSync(filePath)) {
    // Not a local file — treat it as an artifact id and resolve it to a local
    // file path using the native Walnut resolver (same one the built-in
    // file_upload action uses).
    ctx.log('Resolving artifact: ' + fileInput);
    filePath = await ctx.resolveArtifact(fileInput);
    ctx.log('Resolved artifact to: ' + filePath);
  }

  // Click the linked element and intercept the native file chooser, then set the file.
  // This works whether the element is a <div> button, a <label>, or an <input type="file">.
  const [chooser] = await Promise.all([
    webCtx.page.waitForEvent('filechooser'),
    locator.click(),
  ]);
  await chooser.setFiles(filePath);
  ctx.log('Uploaded file to element: ' + filePath);
}
