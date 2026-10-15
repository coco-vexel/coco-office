import type {
  PptWritebackCapabilities,
  PptWritebackOperation,
  PptWritebackPlan,
} from './writeback-planner.js';

export type PptHostKind =
  | 'microsoft-powerpoint'
  | 'wps-presentation'
  | 'onlyoffice-presentation'
  | 'desktop-package'
  | 'custom';

export type PptHostCapabilityStatus = 'unproven' | 'probed' | 'supported';

export interface PptHostWritebackProfile {
  host: PptHostKind;
  status: PptHostCapabilityStatus;
  capabilities: PptWritebackCapabilities;
  notes: string[];
}

export interface PptHostWritebackProfileOptions {
  capabilities?: PptWritebackCapabilities;
  notes?: string[];
  status?: PptHostCapabilityStatus;
}

export interface PptWritebackPlanValidation {
  ok: boolean;
  unsupportedReasons: string[];
  warnings: string[];
}

export type PptHostProbeCheckName =
  | 'package.replace'
  | 'slide.replace'
  | 'slide.insert'
  | 'slide.delete'
  | 'slide.reorder'
  | 'shape.add'
  | 'shape.delete'
  | 'shape.name.write'
  | 'shape.zOrder.write'
  | 'shape.geometry.write'
  | 'shape.style.write'
  | 'text.write'
  | 'text.style.write'
  | 'image.write'
  | 'table.cell.write'
  | 'table.cell.style.write'
  | 'table.values.write'
  | 'table.structure.write'
  | 'save.freshness';

export interface PptHostProbeCheck {
  name: PptHostProbeCheckName;
  passed: boolean;
  message?: string;
  details?: Record<string, unknown>;
}

export interface PptHostProbeReport {
  host: PptHostKind;
  generatedAt?: string;
  adapterVersion?: string;
  checks: PptHostProbeCheck[];
  maxMutationCount?: number;
  notes?: string[];
}

export interface PptHostProbeSummary {
  host: PptHostKind;
  passed: PptHostProbeCheckName[];
  failed: PptHostProbeCheckName[];
  capabilities: PptWritebackCapabilities;
  notes: string[];
}

export interface PptHostProbeAdapter {
  host: PptHostKind;
  adapterVersion?: string;
  runCheck(check: PptHostProbeCheckName): Promise<Omit<PptHostProbeCheck, 'name'> | boolean>;
}

export interface PptHostProbeRunOptions {
  checks?: PptHostProbeCheckName[];
  generatedAt?: string;
  maxMutationCount?: number;
  notes?: string[];
}

export function createPptHostWritebackProfile(
  host: PptHostKind,
  options: PptHostWritebackProfileOptions = {},
): PptHostWritebackProfile {
  const base = defaultPptHostWritebackProfile(host);
  return {
    host,
    status: options.status ?? base.status,
    capabilities: {
      ...base.capabilities,
      ...options.capabilities,
      supportedMutationTypes:
        options.capabilities?.supportedMutationTypes ?? base.capabilities.supportedMutationTypes,
    },
    notes: [...base.notes, ...(options.notes ?? [])],
  };
}

export async function runPptHostProbe(
  adapter: PptHostProbeAdapter,
  options: PptHostProbeRunOptions = {},
): Promise<PptHostProbeReport> {
  const checks = options.checks ?? getDefaultPptHostProbeChecks(adapter.host);
  const results: PptHostProbeCheck[] = [];
  for (const name of checks) {
    try {
      const result = await adapter.runCheck(name);
      results.push(typeof result === 'boolean' ? { name, passed: result } : { name, ...result });
    } catch (error) {
      results.push({ name, passed: false, message: errorMessage(error) });
    }
  }
  return {
    host: adapter.host,
    ...(options.generatedAt ? { generatedAt: options.generatedAt } : {}),
    ...(adapter.adapterVersion ? { adapterVersion: adapter.adapterVersion } : {}),
    checks: results,
    ...(options.maxMutationCount !== undefined
      ? { maxMutationCount: options.maxMutationCount }
      : {}),
    ...(options.notes ? { notes: options.notes } : {}),
  };
}

export function getDefaultPptHostProbeChecks(host: PptHostKind): PptHostProbeCheckName[] {
  if (host === 'desktop-package') return ['package.replace', 'save.freshness'];
  return [
    'slide.replace',
    'slide.insert',
    'slide.delete',
    'slide.reorder',
    'shape.add',
    'shape.delete',
    'shape.name.write',
    'shape.zOrder.write',
    'shape.geometry.write',
    'shape.style.write',
    'text.write',
    'text.style.write',
    'image.write',
    'table.cell.write',
    'table.cell.style.write',
    'table.values.write',
    'table.structure.write',
    'save.freshness',
  ];
}

export function summarizePptHostProbeReport(report: PptHostProbeReport): PptHostProbeSummary {
  const passed = uniqueChecks(
    report.checks.filter((check) => check.passed).map((check) => check.name),
  );
  const failed = uniqueChecks(
    report.checks.filter((check) => !check.passed).map((check) => check.name),
  );
  const capabilities = capabilitiesFromPassedProbeChecks(passed, report.maxMutationCount);
  const failedNotes = report.checks
    .filter((check) => !check.passed && check.message)
    .map((check) => `${check.name}: ${check.message}`);
  return {
    host: report.host,
    passed,
    failed,
    capabilities,
    notes: [...(report.notes ?? []), ...failedNotes],
  };
}

export function createPptHostWritebackProfileFromProbeReport(
  report: PptHostProbeReport,
): PptHostWritebackProfile {
  const summary = summarizePptHostProbeReport(report);
  return createPptHostWritebackProfile(report.host, {
    status: 'probed',
    capabilities: summary.capabilities,
    notes: summary.notes,
  });
}

export function getPptWritebackOperationRequirements(
  operation: PptWritebackOperation,
): Array<keyof PptWritebackCapabilities> {
  switch (operation.type) {
    case 'replacePresentationPackage':
      return ['canReplacePresentationPackage'];
    case 'replaceSlidePackage':
      return ['canReplaceSlides'];
    case 'insertSlidePackage':
      return ['canInsertSlides'];
    case 'deleteSlide':
      return ['canDeleteSlides'];
    case 'moveSlide':
      return ['canReorderSlides'];
    case 'replayMutations':
      return ['canReplayMutations'];
  }
}

export function validatePptWritebackPlan(
  plan: PptWritebackPlan,
  profileOrCapabilities: PptHostWritebackProfile | PptWritebackCapabilities,
): PptWritebackPlanValidation {
  let profile: PptHostWritebackProfile | undefined;
  let capabilities: PptWritebackCapabilities;
  if (isHostProfile(profileOrCapabilities)) {
    profile = profileOrCapabilities;
    capabilities = profile.capabilities;
  } else {
    capabilities = profileOrCapabilities;
  }
  const unsupportedReasons: string[] = [];
  const warnings: string[] = [...plan.warnings];

  if (plan.strategy === 'unsupported') {
    unsupportedReasons.push(...plan.unsupportedReasons);
  }

  if (profile?.status === 'unproven' && plan.operations.length > 0) {
    warnings.push(
      `host profile ${profile.host} is unproven; adapter must run probes before writeback`,
    );
  }

  for (const operation of plan.operations) {
    const missing = getPptWritebackOperationRequirements(operation).filter(
      (capability) => capabilities[capability] !== true,
    );
    for (const capability of missing) {
      unsupportedReasons.push(`operation ${operation.type} requires ${capability}`);
    }

    if (operation.type === 'replayMutations') {
      const maxMutationCount = capabilities.maxMutationCount ?? Number.POSITIVE_INFINITY;
      if (operation.mutations.length > maxMutationCount) {
        unsupportedReasons.push(
          `operation replayMutations has ${operation.mutations.length} mutations; maxMutationCount is ${maxMutationCount}`,
        );
      }
      const supported = capabilities.supportedMutationTypes
        ? new Set(capabilities.supportedMutationTypes)
        : null;
      if (supported) {
        const unsupportedTypes = operation.mutations
          .map((mutation) => mutation.type)
          .filter((type) => !supported.has(type));
        if (unsupportedTypes.length > 0) {
          unsupportedReasons.push(
            `operation replayMutations includes unsupported mutation types: ${[...new Set(unsupportedTypes)].join(', ')}`,
          );
        }
      }
    }
  }

  return {
    ok: unsupportedReasons.length === 0,
    unsupportedReasons: [...new Set(unsupportedReasons)],
    warnings: [...new Set(warnings)],
  };
}

function defaultPptHostWritebackProfile(host: PptHostKind): PptHostWritebackProfile {
  switch (host) {
    case 'desktop-package':
      return {
        host,
        status: 'supported',
        capabilities: {
          canReplacePresentationPackage: true,
        },
        notes: ['Desktop package profile supports full pptx package replacement only.'],
      };
    case 'microsoft-powerpoint':
      return {
        host,
        status: 'unproven',
        capabilities: {},
        notes: [
          'Office.js PowerPoint capabilities must be enabled only after runnable host probes.',
        ],
      };
    case 'wps-presentation':
      return {
        host,
        status: 'unproven',
        capabilities: {},
        notes: ['WPSJS Presentation capabilities must be enabled only after runnable host probes.'],
      };
    case 'onlyoffice-presentation':
      return {
        host,
        status: 'unproven',
        capabilities: {},
        notes: [
          'ONLYOFFICE Presentation capabilities must be enabled only after runnable host probes.',
        ],
      };
    case 'custom':
      return {
        host,
        status: 'unproven',
        capabilities: {},
        notes: ['Custom host capabilities are adapter-declared and should be probe-backed.'],
      };
  }
}

function capabilitiesFromPassedProbeChecks(
  passed: PptHostProbeCheckName[],
  maxMutationCount: number | undefined,
): PptWritebackCapabilities {
  const checks = new Set(passed);
  const supportedMutationTypes = mutationTypesFromPassedProbeChecks(checks);
  return {
    ...(checks.has('package.replace') && checks.has('save.freshness')
      ? { canReplacePresentationPackage: true }
      : {}),
    ...(checks.has('slide.replace') && checks.has('save.freshness')
      ? { canReplaceSlides: true }
      : {}),
    ...(checks.has('slide.insert') && checks.has('save.freshness')
      ? { canInsertSlides: true }
      : {}),
    ...(checks.has('slide.delete') && checks.has('save.freshness')
      ? { canDeleteSlides: true }
      : {}),
    ...(checks.has('slide.reorder') && checks.has('save.freshness')
      ? { canReorderSlides: true }
      : {}),
    ...(supportedMutationTypes.length > 0 && checks.has('save.freshness')
      ? { canReplayMutations: true, supportedMutationTypes }
      : {}),
    ...(maxMutationCount !== undefined ? { maxMutationCount } : {}),
  };
}

function mutationTypesFromPassedProbeChecks(
  checks: Set<PptHostProbeCheckName>,
): NonNullable<PptWritebackCapabilities['supportedMutationTypes']> {
  const types = [];
  if (checks.has('shape.add')) types.push('shape.add');
  if (checks.has('shape.delete')) types.push('shape.delete');
  if (checks.has('shape.name.write')) types.push('shape.name.set');
  if (checks.has('shape.zOrder.write')) types.push('shape.zOrder.set');
  if (checks.has('shape.geometry.write')) types.push('shape.geometry.set');
  if (checks.has('shape.style.write')) types.push('shape.style.set');
  if (checks.has('text.write')) {
    types.push('shape.text.set', 'shape.text.replace');
  }
  if (checks.has('text.style.write')) types.push('shape.text.style.set');
  if (checks.has('image.write')) types.push('shape.image.replace');
  if (checks.has('table.cell.write')) types.push('table.cell.text.set');
  if (checks.has('table.cell.style.write')) types.push('table.cell.style.set');
  if (checks.has('table.values.write')) types.push('table.values.set');
  if (checks.has('table.structure.write')) {
    types.push('table.row.add', 'table.row.delete', 'table.column.add', 'table.column.delete');
  }
  return types as NonNullable<PptWritebackCapabilities['supportedMutationTypes']>;
}

function uniqueChecks(checks: PptHostProbeCheckName[]): PptHostProbeCheckName[] {
  return [...new Set(checks)];
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function isHostProfile(
  value: PptHostWritebackProfile | PptWritebackCapabilities,
): value is PptHostWritebackProfile {
  return 'host' in value && 'capabilities' in value;
}
