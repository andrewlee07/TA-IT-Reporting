import { EDITOR_SECTIONS, type SectionId } from "@/lib/drafts/types";
import type {
  NormalizedReportSnapshot,
  PeriodRow,
  ChartSettingRow,
  EntityRow,
  OfficeLocationRow,
  ServiceAvailabilityRow,
  OfficeNetworkAvailabilityRow,
  SupportOperationsRow,
  OldestTicketRow,
  SecurityPatchingRow,
  AssetsLifecycleRow,
  ChangeReleaseRow,
  DevDeliveryRow,
  ProjectPortfolioRow,
  RollingRoadmapRow,
  PortfolioGanttWorkstreamRow,
  PortfolioGanttMilestoneRow,
  BudgetCommercialRow,
  TopRiskRow,
  NarrativeNoteRow,
} from "@/lib/workbook/types";

export interface OverviewSetupSectionData {
  title: string;
  reportSeriesKey: string;
  currentMonth: string;
  availableMonths: string[];
  periods: PeriodRow[];
  entities: EntityRow[];
  officeLocations: OfficeLocationRow[];
  chartSettings: ChartSettingRow[];
}

export interface AvailabilityNetworkSectionData {
  serviceAvailability: ServiceAvailabilityRow[];
  officeNetworkAvailability: OfficeNetworkAvailabilityRow[];
}

export interface SupportOperationsSectionData {
  supportOperations: SupportOperationsRow[];
  oldestTickets: OldestTicketRow[];
}

export interface SecurityAssetsSectionData {
  securityPatching: SecurityPatchingRow[];
  assetsLifecycle: AssetsLifecycleRow[];
}

export interface ChangeDeliverySectionData {
  changeRelease: ChangeReleaseRow[];
  devDelivery: DevDeliveryRow[];
}

export interface ProjectsRoadmapSectionData {
  projectPortfolio: ProjectPortfolioRow[];
  rollingRoadmap: RollingRoadmapRow[];
  portfolioGanttWorkstreams: PortfolioGanttWorkstreamRow[];
  portfolioGanttMilestones: PortfolioGanttMilestoneRow[];
}

export interface FinanceRisksSectionData {
  budgetCommercials: BudgetCommercialRow[];
  topRisks: TopRiskRow[];
}

export interface NotesNarrativeSectionData {
  narrativeNotes: NarrativeNoteRow[];
}

export type SectionPayloadMap = {
  "overview-setup": OverviewSetupSectionData;
  "availability-network": AvailabilityNetworkSectionData;
  "support-operations": SupportOperationsSectionData;
  "security-assets": SecurityAssetsSectionData;
  "change-delivery": ChangeDeliverySectionData;
  "projects-roadmap": ProjectsRoadmapSectionData;
  "finance-risks": FinanceRisksSectionData;
  "notes-narrative": NotesNarrativeSectionData;
};

export function isSectionId(value: string): value is SectionId {
  return EDITOR_SECTIONS.includes(value as SectionId);
}

export function getSectionPayload<S extends SectionId>(
  snapshot: NormalizedReportSnapshot,
  sectionId: S,
  title: string,
  reportSeriesKey: string,
): SectionPayloadMap[S] {
  switch (sectionId) {
    case "overview-setup":
      return {
        title,
        reportSeriesKey,
        currentMonth: snapshot.currentMonth,
        availableMonths: snapshot.availableMonths,
        periods: snapshot.periods,
        entities: snapshot.entities,
        officeLocations: snapshot.officeLocations,
        chartSettings: snapshot.chartSettings,
      } as SectionPayloadMap[S];
    case "availability-network":
      return {
        serviceAvailability: snapshot.serviceAvailability,
        officeNetworkAvailability: snapshot.officeNetworkAvailability,
      } as SectionPayloadMap[S];
    case "support-operations":
      return {
        supportOperations: snapshot.supportOperations,
        oldestTickets: snapshot.oldestTickets,
      } as SectionPayloadMap[S];
    case "security-assets":
      return {
        securityPatching: snapshot.securityPatching,
        assetsLifecycle: snapshot.assetsLifecycle,
      } as SectionPayloadMap[S];
    case "change-delivery":
      return {
        changeRelease: snapshot.changeRelease,
        devDelivery: snapshot.devDelivery,
      } as SectionPayloadMap[S];
    case "projects-roadmap":
      return {
        projectPortfolio: snapshot.projectPortfolio,
        rollingRoadmap: snapshot.rollingRoadmap,
        portfolioGanttWorkstreams: snapshot.portfolioGanttWorkstreams,
        portfolioGanttMilestones: snapshot.portfolioGanttMilestones,
      } as SectionPayloadMap[S];
    case "finance-risks":
      return {
        budgetCommercials: snapshot.budgetCommercials,
        topRisks: snapshot.topRisks,
      } as SectionPayloadMap[S];
    case "notes-narrative":
      return {
        narrativeNotes: snapshot.narrativeNotes,
      } as SectionPayloadMap[S];
  }
}

export function applySectionPayload<S extends SectionId>(
  snapshot: NormalizedReportSnapshot,
  sectionId: S,
  payload: SectionPayloadMap[S],
): { snapshot: NormalizedReportSnapshot; title?: string; reportSeriesKey?: string } {
  switch (sectionId) {
    case "overview-setup": {
      const typed = payload as OverviewSetupSectionData;
      const availableMonths = Array.from(new Set(typed.periods.map((row) => row.reportingMonth))).sort();
      const currentMonth = availableMonths.includes(typed.currentMonth)
        ? typed.currentMonth
        : availableMonths[availableMonths.length - 1] ?? typed.currentMonth;
      return {
        snapshot: {
          ...snapshot,
          currentMonth,
          availableMonths: availableMonths.length > 0 ? availableMonths : typed.availableMonths,
          periods: typed.periods,
          entities: typed.entities,
          officeLocations: typed.officeLocations,
          chartSettings: typed.chartSettings,
        },
        title: typed.title,
        reportSeriesKey: typed.reportSeriesKey,
      };
    }
    case "availability-network":
      return {
        snapshot: {
          ...snapshot,
          serviceAvailability: (payload as AvailabilityNetworkSectionData).serviceAvailability,
          officeNetworkAvailability: (payload as AvailabilityNetworkSectionData).officeNetworkAvailability,
        },
      };
    case "support-operations":
      return {
        snapshot: {
          ...snapshot,
          supportOperations: (payload as SupportOperationsSectionData).supportOperations,
          oldestTickets: (payload as SupportOperationsSectionData).oldestTickets,
        },
      };
    case "security-assets":
      return {
        snapshot: {
          ...snapshot,
          securityPatching: (payload as SecurityAssetsSectionData).securityPatching,
          assetsLifecycle: (payload as SecurityAssetsSectionData).assetsLifecycle,
        },
      };
    case "change-delivery":
      return {
        snapshot: {
          ...snapshot,
          changeRelease: (payload as ChangeDeliverySectionData).changeRelease,
          devDelivery: (payload as ChangeDeliverySectionData).devDelivery,
        },
      };
    case "projects-roadmap":
      return {
        snapshot: {
          ...snapshot,
          projectPortfolio: (payload as ProjectsRoadmapSectionData).projectPortfolio,
          rollingRoadmap: (payload as ProjectsRoadmapSectionData).rollingRoadmap,
          portfolioGanttWorkstreams: (payload as ProjectsRoadmapSectionData).portfolioGanttWorkstreams,
          portfolioGanttMilestones: (payload as ProjectsRoadmapSectionData).portfolioGanttMilestones,
        },
      };
    case "finance-risks":
      return {
        snapshot: {
          ...snapshot,
          budgetCommercials: (payload as FinanceRisksSectionData).budgetCommercials,
          topRisks: (payload as FinanceRisksSectionData).topRisks,
        },
      };
    case "notes-narrative":
      return {
        snapshot: {
          ...snapshot,
          narrativeNotes: (payload as NotesNarrativeSectionData).narrativeNotes,
        },
      };
  }
}
