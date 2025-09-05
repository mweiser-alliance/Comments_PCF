/* eslint-disable @typescript-eslint/no-explicit-any */
import { IInputs, IOutputs } from "./generated/ManifestTypes";

/**
 * CommentsControl - tailored to your schema:
 *  - MessageProperty  -> ta_Text
 *  - DateProperty     -> ta_MessageDate
 *  - UserProperty     -> ta_User
 *  - ParentLookUpProp -> ta_SolutionCenterRequest   // change to ta_EmployerRequest if needed
 */
const MESSAGE_PROP = "ta_Text";
const DATE_PROP    = "ta_MessageDate";
const USER_PROP    = "ta_User";
const PARENT_PROP  = "ta_SolutionCenterRequest"; // <-- switch to "ta_EmployerRequest" if that's your parent column

export class CommentsControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private _context!: ComponentFramework.Context<IInputs>;
  private _container!: HTMLDivElement;
  private _input!: HTMLTextAreaElement;
  private _isInitialized = false;

  public init(
    context: ComponentFramework.Context<IInputs>,
    _notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._context = context;
    this._container = container;

    // Root
    this._container.classList.add("comments-root");

    // List
    const list = document.createElement("div");
    list.className = "comments-list";
    list.setAttribute("role", "log");

    // Input
    const inputWrap = document.createElement("div");
    inputWrap.className = "comments-input-wrap";

    this._input = document.createElement("textarea");
    this._input.className = "comments-input";
    this._input.rows = 1;
    this._input.placeholder = "Write a comment and press Enter…";
    this._input.addEventListener("keydown", (e) => this.onKeyDown(e));

    inputWrap.appendChild(this._input);
    this._container.appendChild(list);
    this._container.appendChild(inputWrap);

    this._isInitialized = true;
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;
    if (!this._isInitialized) return;

    const list = this._container.querySelector(".comments-list") as HTMLDivElement;
    if (!list) return;

    list.innerHTML = "";

    const ds = context.parameters.dataset;
    if (!ds || ds.loading) return;

    // Render oldest -> newest; use your writable date if present, else Created On
    const sorted = Object.keys(ds.records)
      .map((k) => ds.records[k])
      .sort((a, b) => {
        const ad = (a.getValue(DATE_PROP) as any) ?? a.getValue("createdon");
        const bd = (b.getValue(DATE_PROP) as any) ?? b.getValue("createdon");
        const at = ad ? new Date(ad).getTime() : 0;
        const bt = bd ? new Date(bd).getTime() : 0;
        return at - bt;
      });

    for (const rec of sorted) {
      const row = document.createElement("div");
      row.className = "comment-row";

      const body = document.createElement("div");
      body.className = "comment-body";
      body.textContent = (rec.getFormattedValue(MESSAGE_PROP) as string) ?? "";

      const meta = document.createElement("div");
      meta.className = "comment-meta";
      const userTxt = (rec.getFormattedValue(USER_PROP) as string) ?? "";
      const dateTxt =
        (rec.getFormattedValue(DATE_PROP) as string) ??
        (rec.getFormattedValue("createdon") as string) ??
        "";
      meta.textContent = [userTxt, dateTxt].filter(Boolean).join(" · ");

      row.appendChild(body);
      row.appendChild(meta);
      list.appendChild(row);
    }
  }

  public getOutputs(): IOutputs { return {}; }

  public destroy(): void {
    // no-op
  }

  // ===== Internals =====

  private async onKeyDown(e: KeyboardEvent) {
    if (e.key === "Enter" && !e.shiftKey) {
      // critical: prevent the form/subgrid from swallowing Enter
      e.preventDefault();
      e.stopPropagation();

      const text = (this._input.value || "").trim();
      if (!text) return;

      try {
        await this.createComment(text);
        this._input.value = "";
        await this._context.parameters.dataset.refresh();
      } catch (err) {
        // eslint-disable-next-line no-console
        console.error("Failed to create comment", err);
      }
    }
  }

  private async createComment(messageText: string): Promise<void> {
    const ctx = this._context;

    // Determine child entity (Chat) from dataset binding
    const chatEntityLogicalName =
      ctx.parameters.dataset.getTargetEntityType &&
      ctx.parameters.dataset.getTargetEntityType();

    if (!chatEntityLogicalName) {
      throw new Error("Cannot determine chat entity from dataset binding.");
    }

    // Resolve current user id and parent record (entity name + id)
    const userId = (ctx.userSettings.userId || "").replace(/[{}]/g, "");
    const { parentEntityName, parentId } = this.getParentFromPageContext(ctx);

    if (!parentEntityName || !parentId) {
      throw new Error("Could not determine the parent record from page context.");
    }

    // Resolve PLURAL entity set names for @odata.bind
    const sysUserMeta = await (ctx.utils as any).getEntityMetadata("systemuser");
    const parentMeta  = await (ctx.utils as any).getEntityMetadata(parentEntityName);
    const systemUsersSet = sysUserMeta.EntitySetName; // e.g., "systemusers"
    const parentSet      = parentMeta.EntitySetName;  // e.g., "ta_solutioncenterrequests"

    // Build payload
    const payload: Record<string, any> = {};
    payload[MESSAGE_PROP] = messageText;
    payload[DATE_PROP]    = new Date().toISOString();
    payload[`${USER_PROP}@odata.bind`]   = `/${systemUsersSet}(${userId})`;
    payload[`${PARENT_PROP}@odata.bind`] = `/${parentSet}(${parentId})`;

    // Create row
    await ctx.webAPI.createRecord(chatEntityLogicalName, payload);
  }

  /**
   * Tries multiple strategies to get the parent record info when the control is hosted on a subgrid within a form.
   */
  private getParentFromPageContext(ctx: ComponentFramework.Context<IInputs>): { parentEntityName: string; parentId: string } {
    // 1) Preferred: context.mode.contextInfo (newer PCF runtimes)
    const ci = (ctx.mode as any)?.contextInfo;
    if (ci?.entityTypeName && ci?.entityId) {
      return {
        parentEntityName: ci.entityTypeName as string,
        parentId: (ci.entityId as string).replace(/[{}]/g, "")
      };
    }

    // 2) Fallback: legacy Xrm.Page (if available)
    const XrmAny = (window as any).Xrm;
    if (XrmAny?.Page?.data?.entity) {
      const id = XrmAny.Page.data.entity.getId?.() || "";
      const name = XrmAny.Page.data.entity.getEntityName?.() || "";
      if (id && name) {
        return {
          parentEntityName: name as string,
          parentId: (id as string).replace(/[{}]/g, "")
        };
      }
    }

    // 3) Fallback: parse from URL (etn + id)
    try {
      const url = new URL(window.location.href);
      const etn = url.searchParams.get("etn");
      const id  = url.searchParams.get("id");
      if (etn && id) {
        return { parentEntityName: etn, parentId: id.replace(/[{}]/g, "") };
      }
    } catch { /* ignore */ }

    return { parentEntityName: "", parentId: "" };
  }
}
