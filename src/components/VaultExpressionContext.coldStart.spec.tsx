import { act, renderHook, waitFor } from "@testing-library/react";
import type { ReactNode } from "react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { VaultEntry } from "../types";
import {
	useResolvedVaultExpressionTemplate,
	VaultExpressionProvider,
} from "./VaultExpressionContext";

const engine = vi.hoisted(() => ({
	ready: false,
	resolve: () => {},
	init: vi.fn(),
}));
vi.mock("@ironcalc/workbook", () => ({ init: engine.init }));
vi.mock("../utils/sheetWorkbook", () => ({
	SHEET_INDEX: 0,
	sheetExternalFormulaContext: vi.fn(),
	resolveExternalSheetDependencyEntries: () => [],
	buildWorkbook: () => {
		if (!engine.ready)
			throw new TypeError(
				"Cannot read properties of undefined (reading '__wbindgen_add_to_stack_pointer')",
			);
		return {
			model: {
				getCellContent: (_sheet: number, row: number) =>
					row === 1 ? "42" : "",
				free: vi.fn(),
			},
		};
	},
}));
const sheet = {
	path: "/vault/sheet.md",
	title: "sheet",
	filename: "sheet.md",
	aliases: [],
	properties: {},
} as VaultEntry;
function wrapper({ children }: { children: ReactNode }) {
	return (
		<VaultExpressionProvider
			currentContent="42"
			entries={[sheet]}
			locale="en-US"
			sourceEntry={sheet}
			vaultPath="/vault"
		>
			{children}
		</VaultExpressionProvider>
	);
}

describe("cold-start HTML sheet expressions", () => {
	beforeEach(() => {
		engine.ready = false;
		engine.init.mockImplementation(
			() =>
				new Promise<void>((resolve) => {
					engine.resolve = () => {
						engine.ready = true;
						resolve();
					};
				}),
		);
	});
	it("waits for the engine, then resolves simultaneous sheet-cell references", async () => {
		const first = renderHook(
			() => useResolvedVaultExpressionTemplate("Answer: {{[[sheet]].A1}}"),
			{ wrapper },
		);
		const second = renderHook(
			() =>
				useResolvedVaultExpressionTemplate(
					"{{[[sheet]].A1}} / {{[[sheet]].A1}}",
				),
			{ wrapper },
		);
		const closed = renderHook(
			() => useResolvedVaultExpressionTemplate("{{[[sheet]].A1}}"),
			{ wrapper },
		);
		closed.unmount();
		expect(engine.init).toHaveBeenCalledTimes(1);
		expect(first.result.current).toEqual({
			html: "<div></div>",
			unresolved: [],
		});
		await act(async () => engine.resolve());
		await waitFor(() =>
			expect(first.result.current).toEqual({
				html: "Answer: 42",
				unresolved: [],
			}),
		);
		expect(second.result.current).toEqual({ html: "42 / 42", unresolved: [] });
		expect(first.result.current.unresolved).toEqual([]);
		const missing = renderHook(
			() =>
				useResolvedVaultExpressionTemplate(
					"{{[[missing]].A1}} / {{[[sheet]].A2}}",
				),
			{ wrapper },
		);
		await waitFor(() =>
			expect(missing.result.current).toEqual({
				html: "{{[[missing]].A1}} / ",
				unresolved: ["[[missing]].A1"],
			}),
		);
		expect(missing.result.current.unresolved).toEqual(["[[missing]].A1"]);
	});
});
