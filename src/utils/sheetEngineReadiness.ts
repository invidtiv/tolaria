import ironCalcWasmUrl from "@ironcalc/wasm/wasm_bg.wasm?url";
import { init as initIronCalc } from "@ironcalc/workbook";

let initialization: Promise<void> | null = null;

export function ensureIronCalcReady(): Promise<void> {
	if (!initialization) {
		initialization = initIronCalc(ironCalcWasmUrl)
			.then(() => undefined)
			.catch((error: unknown) => {
				initialization = null;
				throw error;
			});
	}
	return initialization;
}
