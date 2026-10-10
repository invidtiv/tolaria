import {
	createContext,
	createElement,
	type ReactNode,
	useContext,
	useEffect,
	useMemo,
	useState,
} from "react";
import {
	getCachedNoteContentEntry,
	hasResolvedCachedContent,
	prefetchNoteContent,
	subscribeNoteContentResolved,
} from "../hooks/noteContentCache";
import type { VaultEntry } from "../types";
import { ensureIronCalcReady } from "../utils/sheetEngineReadiness";
import { resolveExternalSheetDependencyEntries } from "../utils/sheetWorkbook";
import {
	type CompiledVaultExpressionTemplate,
	compileVaultExpressionTemplate,
	renderVaultExpressionTemplate,
	vaultExpressionDependencySource,
} from "../utils/vaultExpressions";

interface VaultExpressionContextValue {
	currentContent: string;
	entries: VaultEntry[];
	locale: string;
	sourceEntry: VaultEntry | null;
	vaultPath: string;
}

interface VaultExpressionProviderProps extends VaultExpressionContextValue {
	children: ReactNode;
}

interface ResolvedVaultExpressionTemplate {
	html: string;
	unresolved: string[];
}

const VaultExpressionContext =
	createContext<VaultExpressionContextValue | null>(null);
const EMPTY_VAULT_EXPRESSION_CONTEXT: VaultExpressionContextValue = {
	currentContent: "",
	entries: [],
	locale: "en-US",
	sourceEntry: null,
	vaultPath: "",
};

export function useVaultExpressionContext(): VaultExpressionContextValue {
	return useContext(VaultExpressionContext) ?? EMPTY_VAULT_EXPRESSION_CONTEXT;
}

export function VaultExpressionProvider({
	children,
	currentContent,
	entries,
	locale,
	sourceEntry,
	vaultPath,
}: VaultExpressionProviderProps) {
	const value = useMemo(
		() => ({
			currentContent,
			entries,
			locale,
			sourceEntry,
			vaultPath,
		}),
		[currentContent, entries, locale, sourceEntry, vaultPath],
	);

	return createElement(VaultExpressionContext.Provider, { value }, children);
}

function cachedContentForEntry(entry: VaultEntry): string | null {
	const cached = getCachedNoteContentEntry(entry.path);
	return hasResolvedCachedContent(cached) ? cached.value : null;
}

function dependencyEntries({
	compiled,
	contentsByPath,
	context,
}: {
	compiled: CompiledVaultExpressionTemplate;
	contentsByPath: Map<string, string>;
	context: VaultExpressionContextValue;
}): VaultEntry[] {
	const dependencySource = vaultExpressionDependencySource(compiled);
	if (dependencySource === "") return [];

	return resolveExternalSheetDependencyEntries({
		content: dependencySource,
		contentsByPath,
		currentPath: context.sourceEntry?.path ?? "",
		entries: context.entries,
		sourceEntry: context.sourceEntry,
	});
}

function mergeCachedDependencyContents(
	entries: VaultEntry[],
): Record<string, string> {
	const cachedContents: Record<string, string> = {};
	for (const entry of entries) {
		const content = cachedContentForEntry(entry);
		if (content === null) {
			prefetchNoteContent(entry, { parsedBlockPreload: false });
		} else {
			Reflect.set(cachedContents, entry.path, content);
		}
	}
	return cachedContents;
}

function retainDependencyContents(
	paths: Set<string>,
	cached: Record<string, string>,
	current: Record<string, string>,
): Record<string, string> {
	const next: Record<string, string> = {};
	for (const path of paths) {
		const cachedContent = Reflect.get(cached, path) as string | undefined;
		const currentContent = Reflect.get(current, path) as string | undefined;
		if (cachedContent !== undefined) {
			Reflect.set(next, path, cachedContent);
		} else if (currentContent !== undefined) {
			Reflect.set(next, path, currentContent);
		}
	}
	return next;
}

function sameContents(
	left: Record<string, string>,
	right: Record<string, string>,
): boolean {
	const leftKeys = Object.keys(left);
	const rightKeys = Object.keys(right);
	return (
		leftKeys.length === rightKeys.length &&
		leftKeys.every((key) => Reflect.get(left, key) === Reflect.get(right, key))
	);
}

function deferStateUpdate(update: () => void): void {
	queueMicrotask(update);
}

function useVaultExpressionDependencyContents(
	compiled: CompiledVaultExpressionTemplate,
	context: VaultExpressionContextValue,
): Map<string, string> {
	const [contents, setContents] = useState<Record<string, string>>({});
	const contentsByPath = useMemo(
		() => new Map(Object.entries(contents)),
		[contents],
	);
	const entries = useMemo(
		() => dependencyEntries({ compiled, contentsByPath, context }),
		[compiled, contentsByPath, context],
	);
	const pathKey = useMemo(
		() =>
			entries
				.map((entry) => entry.path)
				.sort()
				.join("\n"),
		[entries],
	);

	useEffect(() => {
		let subscribed = true;
		const paths = new Set(pathKey === "" ? [] : pathKey.split("\n"));
		const cached = mergeCachedDependencyContents(entries);
		deferStateUpdate(() => {
			if (!subscribed) return;
			setContents((current) => {
				const next = retainDependencyContents(paths, cached, current);
				return sameContents(current, next) ? current : next;
			});
		});

		const unsubscribe = subscribeNoteContentResolved((event) => {
			if (!paths.has(event.path)) return;
			setContents((current) =>
				current[event.path] === event.content
					? current
					: { ...current, [event.path]: event.content },
			);
		});
		return () => {
			subscribed = false;
			unsubscribe();
		};
	}, [entries, pathKey]);

	return contentsByPath;
}

type TemplateExpressionAst = NonNullable<
	Exclude<CompiledVaultExpressionTemplate["parts"][number], string>["ast"]
>;

function expressionReadsCell(ast: TemplateExpressionAst): boolean {
	if (ast.type === "reference") return ast.kind === "cell";
	if (ast.type === "call") return ast.args.some(expressionReadsCell);
	if (ast.type === "binary")
		return expressionReadsCell(ast.left) || expressionReadsCell(ast.right);
	return false;
}

function templateReadsCell(compiled: CompiledVaultExpressionTemplate): boolean {
	return compiled.parts.some(
		(part) =>
			typeof part !== "string" &&
			part.ast !== null &&
			expressionReadsCell(part.ast),
	);
}

function useExpressionEngineReady(required: boolean): boolean {
	const [ready, setReady] = useState(false);
	useEffect(() => {
		if (!required) return;
		let subscribed = true;
		ensureIronCalcReady()
			.then(() => {
				if (subscribed) setReady(true);
			})
			.catch((error: unknown) => {
				console.warn("[html] Failed to initialize spreadsheet engine:", error);
			});
		return () => {
			subscribed = false;
		};
	}, [required]);
	return !required || ready;
}

export function useResolvedVaultExpressionTemplate(
	source: string,
): ResolvedVaultExpressionTemplate {
	const expressionContext = useVaultExpressionContext();
	const compiled = useMemo(
		() => compileVaultExpressionTemplate(source),
		[source],
	);
	const contentsByPath = useVaultExpressionDependencyContents(
		compiled,
		expressionContext,
	);
	const ready = useExpressionEngineReady(templateReadsCell(compiled));

	return useMemo(
		() =>
			ready
				? renderVaultExpressionTemplate({
						compiled,
						context: {
							contentsByPath,
							...expressionContext,
						},
					})
				: { html: "<div></div>", unresolved: [] },
		[compiled, contentsByPath, expressionContext, ready],
	);
}
