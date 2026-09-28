let showing = false;

/** Office.js disables browser confirmation dialogs. No write runs without approval. */
export function confirmAction(message: string, title = "Confirm change", confirmLabel = "Continue"): Promise<boolean> {
  if (showing) return Promise.resolve(false);
  showing = true;
  return new Promise(resolve => {
    const previousFocus = document.activeElement;
    const dialog = document.createElement("dialog");
    dialog.className = "cupola-confirmation";
    dialog.setAttribute("aria-labelledby", "cupola-confirm-title");
    dialog.setAttribute("aria-describedby", "cupola-confirm-message");
    const heading = document.createElement("h2");
    heading.id = "cupola-confirm-title";
    heading.textContent = title;
    const description = document.createElement("p");
    description.id = "cupola-confirm-message";
    description.textContent = message;
    const actions = document.createElement("div");
    actions.className = "actions";
    const cancel = document.createElement("button");
    cancel.textContent = "Cancel";
    cancel.autofocus = true;
    const approve = document.createElement("button");
    approve.textContent = confirmLabel;
    approve.className = "primary";
    let finished = false;
    const finish = (accepted: boolean) => {
      if (finished) return;
      finished = true;
      if (dialog.open) dialog.close();
      dialog.remove();
      showing = false;
      if (previousFocus instanceof HTMLElement && previousFocus.isConnected) previousFocus.focus();
      resolve(accepted);
    };
    cancel.onclick = () => finish(false);
    approve.onclick = () => finish(true);
    dialog.addEventListener("cancel", event => { event.preventDefault(); finish(false); });
    dialog.addEventListener("close", () => finish(false));
    dialog.addEventListener("keydown", event => {
      if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); finish(false); }
      if (event.key === "Tab") {
        event.preventDefault();
        (document.activeElement === cancel ? approve : cancel).focus();
      }
    });
    actions.append(cancel, approve);
    dialog.append(heading, description, actions);
    document.body.append(dialog);
    try { dialog.showModal(); cancel.focus(); } catch { finish(false); }
  });
}
