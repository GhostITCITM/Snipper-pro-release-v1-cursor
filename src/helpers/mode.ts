import { writeTick, writeCross } from "../excel/excel";

export type SnipMode = "text" | "sum" | "table" | "validation" | "exception";

class ModeManager {
  private _mode: SnipMode | null = null;

  async setMode(m: SnipMode): Promise<void> {
    this._mode = m;
    await OfficeRuntime.storage.setItem("snip-mode", m);
    await Office.addin.showAsTaskpane();
  }

  async current(): Promise<SnipMode | null> {
    if (this._mode) return this._mode;
    return OfficeRuntime.storage.getItem<SnipMode>("snip-mode");
  }

  async clearMode(): Promise<void> {
    this._mode = null;
    await OfficeRuntime.storage.removeItem("snip-mode");
  }

  async isActive(): Promise<boolean> {
    const mode = await this.current();
    return mode !== null;
  }
}

export const modeManager = new ModeManager();

// Selection event to re-enable buttons and handle immediate validation/exception modes
Office.onReady(() => {
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, async () => {
    try {
      const m = await modeManager.current();
      if (m === "validation") {
        await writeTick();
        await modeManager.clearMode();
      }
      if (m === "exception") {
        await writeCross();
        await modeManager.clearMode();
      }
    } catch (error) {
      console.warn("Error handling selection change:", error);
    }
  });
});
