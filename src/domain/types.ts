/**
 * Bit domain types (M2).
 * One workbook = one Project. Main is the default scenario tip pointer.
 */

export interface Project {
  id: string;
  name: string;
  createdAt: string;
  /** Scenario id for Main (always present after create). */
  mainScenarioId: string;
}

export interface Scenario {
  id: string;
  projectId: string;
  name: string;
  /** Tip version id; null only before first version (should not happen after create). */
  tipVersionId: string | null;
  /** Main is the protected default scenario. */
  isMain: boolean;
}

export interface Version {
  id: string;
  projectId: string;
  scenarioId: string;
  /** Parent version ids (one for normal Save version; two after combine later). */
  parentIds: string[];
  author: string;
  timestamp: string;
  message: string;
  /** Content-addressed snapshot blob hash. */
  snapshotHash: string;
}

export interface ProjectDetail extends Project {
  main: Scenario;
  scenarios: Scenario[];
  tipVersion: Version | null;
}

export interface VersionListItem extends Version {
  scenarioName: string;
}

export type ReviewStatus = 'open' | 'changes-requested' | 'combined' | 'closed';

export interface Review {
  id: string;
  projectId: string;
  scenarioId: string;
  /** Main tip frozen when review opened. */
  baseVersionId: string;
  /** Scenario tip frozen when review opened. */
  compareVersionId: string;
  author: string;
  note?: string;
  status: ReviewStatus;
  createdAt: string;
  updatedAt: string;
}
