/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

// Fix: Add an empty export to treat this file as a module, which is required for global augmentation.
export {};

// Tell TypeScript about the Chart.js global variable from the CDN
declare var Chart: any;

// Add MathJax to the window type for TypeScript
declare global {
  interface Window {
    MathJax: any;
  }
}

// --- TYPE DEFINITIONS ---
type ApplianceCategory = 'Furnace' | 'Water Heater' | 'Range' | 'Dryer' | 'Boiler' | 'Space Heater';
type EfficiencyTier = 'Standard' | 'High' | 'Premium';
type FacilityType = 'Residential' | 'Small Commercial' | 'Large Commercial' | 'Industrial' | 'Restaurant' | 'Medical' | 'Nursing Home';
type ClimateZone = 'Zone5' | 'Zone6';

interface Appliance {
  id: string;
  key: string; // Key into APPLIANCE_DEFINITIONS
  btu: number;
  included: boolean;
  efficiency: number;
}

interface State {
  facilityType: FacilityType;
  climateZone: ClimateZone;
  addressNumber: string;
  streetName: string;
  town: string;
  annualHeatingTherms: number;
  annualNonHeatingTherms: number;
  panelAmps: number;
  voltage: number;
  breakerSpaces: number;
  efficiencyTier: EfficiencyTier;
  appliances: Appliance[];
  gasPricePerTherm: number;
  electricityPricePerKwh: number;
  gasPriceEscalation: number;
  electricityPriceEscalation: number;
  highRiskGasEscalation: number;
}

interface ApplianceRecommendation {
    id: string;
    text: string;
    size: string; // e.g., "3.0-ton", "50-gallon"
    kw: number; // Peak kW for load calculation
    amps: number;
    breakerSpaces: number;
}

interface PlanningAnalysis {
    recommendations: ApplianceRecommendation[];
    totalNewAmps: number;
    requiredBreakerSpaces: number;
    breakerStatus: 'Sufficient' | 'Sub-panel Recommended' | 'Panel Upgrade Recommended';
    panelStatus: 'Not Required' | `Upgrade to ${number}A Recommended`;
}

interface CostItem {
    name: string;
    low: number;
    high: number;
}

interface RebateItem {
    name: string;
    amount: number;
}

interface CostAnalysis {
    applianceCosts: Map<string, CostItem[]>; // Map<applianceId, CostItem[]>
    electricalCosts: CostItem[];
    rebates: RebateItem[];
    totalLow: number;
    totalHigh: number;
    totalRebates: number;
    netLow: number;
    netHigh: number;
}

interface FinancialAnalysis {
    currentAnnualGasCost: number;
    projectedAnnualElecCost: number;
    netAnnualSavings: number;
    simplePaybackYears: number | null;
}

interface EnvironmentalImpact {
    currentAnnualGasCo2Kg: number;
    currentAnnualCh4Kg: number;
    currentAnnualGasCo2e20Kg: number;
    currentAnnualGasCo2e100Kg: number;
    projectedAnnualElecCo2Kg: number;
    annualGhgReductionCo2e20Kg: number;
    annualGhgReductionCo2e100Kg: number;
    ghgReductionCarsOffRoad100yr: number;
}

interface LifetimeAnalysis {
    totalGasCost15yr: number;
    totalElecCost15yr: number;
    totalSavings15yr: number;
    cumulativeGasCosts: number[];
    cumulativeElecCosts: number[];
    totalHighRiskSavings15yr: number;
    cumulativeHighRiskGasCosts: number[];
}

interface DistributionImpactAnalysis {
    peakDemandKw: number;
    nonHeatingDemandKw: number;
}


// --- CONSTANTS ---

// --- APPLIANCE DATA LIBRARY ---
const APPLIANCE_DEFINITIONS: Record<string, {name: string, category: ApplianceCategory, defaultBtu: number, defaultEfficiency: number}> = {
    // Residential
    'res-furnace': { name: 'Forced Air Furnace', category: 'Furnace', defaultBtu: 80000, defaultEfficiency: 80 },
    'res-boiler': { name: 'Boiler (Hydronic)', category: 'Boiler', defaultBtu: 100000, defaultEfficiency: 80 },
    'res-tank-wh': { name: 'Tank Water Heater', category: 'Water Heater', defaultBtu: 40000, defaultEfficiency: 62 },
    'res-tankless-wh': { name: 'Tankless Water Heater', category: 'Water Heater', defaultBtu: 150000, defaultEfficiency: 82 },
    'res-range': { name: 'Gas Range', category: 'Range', defaultBtu: 60000, defaultEfficiency: 40 },
    'res-dryer': { name: 'Gas Dryer', category: 'Dryer', defaultBtu: 22000, defaultEfficiency: 75 },
    'res-pool-heater': { name: 'Pool Heater', category: 'Space Heater', defaultBtu: 250000, defaultEfficiency: 82 },
    'res-fireplace': { name: 'Gas Fireplace Insert', category: 'Space Heater', defaultBtu: 30000, defaultEfficiency: 75 },
    // Commercial / Industrial / Medical / Nursing
    'comm-rtu': { name: 'Rooftop Unit (RTU)', category: 'Furnace', defaultBtu: 240000, defaultEfficiency: 80 },
    'comm-boiler': { name: 'Commercial Boiler', category: 'Boiler', defaultBtu: 500000, defaultEfficiency: 80 },
    'comm-wh': { name: 'Commercial Water Heater', category: 'Water Heater', defaultBtu: 199000, defaultEfficiency: 80 },
    'laundry-dryer': { name: 'Large Capacity Dryer', category: 'Dryer', defaultBtu: 120000, defaultEfficiency: 75 },
    // Restaurant
    'rest-range': { name: 'Commercial Range', category: 'Range', defaultBtu: 180000, defaultEfficiency: 40 },
    'rest-convection-oven': { name: 'Convection Oven', category: 'Range', defaultBtu: 50000, defaultEfficiency: 40 },
    'rest-fryer': { name: 'Fryer', category: 'Range', defaultBtu: 90000, defaultEfficiency: 40 },
    'rest-griddle': { name: 'Griddle', category: 'Range', defaultBtu: 60000, defaultEfficiency: 40 },
    'rest-booster-heater': { name: 'Booster Heater (Dishwasher)', category: 'Water Heater', defaultBtu: 58000, defaultEfficiency: 80 },
};

const FACILITY_APPLIANCE_MAP: Record<FacilityType, string[]> = {
    'Residential': ['res-furnace', 'res-boiler', 'res-tank-wh', 'res-tankless-wh', 'res-range', 'res-dryer', 'res-pool-heater', 'res-fireplace'],
    'Small Commercial': ['comm-rtu', 'comm-boiler', 'comm-wh'],
    'Large Commercial': ['comm-rtu', 'comm-boiler', 'comm-wh', 'laundry-dryer'],
    'Industrial': ['comm-rtu', 'comm-boiler', 'comm-wh', 'laundry-dryer'],
    'Restaurant': ['rest-range', 'rest-convection-oven', 'rest-fryer', 'rest-griddle', 'comm-rtu', 'comm-wh', 'rest-booster-heater'],
    'Medical': ['comm-rtu', 'comm-boiler', 'comm-wh', 'laundry-dryer', 'rest-range', 'rest-convection-oven'],
    'Nursing Home': ['comm-rtu', 'comm-boiler', 'comm-wh', 'laundry-dryer', 'rest-range', 'rest-convection-oven'],
};


const BTU_PER_THERM = 100000;
const BTU_TO_KW = 0.000293071;
const PANEL_SIZES = [100, 125, 150, 200, 400, 600];
const LIFETIME_YEARS = 15;
const COP_LOOKUP: Record<ApplianceCategory, Record<EfficiencyTier, number>> = {
    'Furnace':      { 'Standard': 2.5, 'High': 3.0, 'Premium': 3.5 },
    'Boiler':       { 'Standard': 2.5, 'High': 3.0, 'Premium': 3.5 },
    'Water Heater': { 'Standard': 2.8, 'High': 3.5, 'Premium': 4.0 },
    'Space Heater': { 'Standard': 2.5, 'High': 3.0, 'Premium': 3.5 }, // Replaced by ductless mini-split
    'Range':        { 'Standard': 1, 'High': 1, 'Premium': 1 }, // Represents efficiency gain over gas
    'Dryer':        { 'Standard': 1.5, 'High': 2.0, 'Premium': 2.5 }, // Represents HP dryer efficiency
};
const CLIMATE_ZONE_COP_ADJUSTMENT: Record<ClimateZone, number> = {
    'Zone5': 1.0,  // Baseline
    'Zone6': 0.9,  // 10% penalty for colder zone
};
const APPLIANCE_KW_LOOKUP: Partial<Record<ApplianceCategory, number>> = { Range: 9.6, Dryer: 1.8 }; // Peak kW for non-BTU appliances
const BREAKER_SPACE_REQUIREMENTS: Record<ApplianceCategory, number> = {
    'Furnace': 2, 'Boiler': 2, 'Water Heater': 2, 'Range': 2, 'Dryer': 2, 'Space Heater': 2
};

// Source: Mass Save, MA market averages (for estimation only)
const DEFAULT_USAGE_BY_FACILITY: Record<FacilityType, {heating: number, nonHeating: number}> = {
    'Residential': { heating: 700, nonHeating: 250 },
    'Small Commercial': { heating: 2000, nonHeating: 800 },
    'Large Commercial': { heating: 15000, nonHeating: 5000 },
    'Industrial': { heating: 50000, nonHeating: 20000 },
    'Restaurant': { heating: 1500, nonHeating: 4000 },
    'Medical': { heating: 10000, nonHeating: 8000 },
    'Nursing Home': { heating: 8000, nonHeating: 6000 },
};


// --- EMISSIONS DATA ---
const CO2_KG_PER_THERM_GAS = 5.3; // Source: EPA, Combustion only
// Methane leakage and GWP based on MA Climate Plan guidelines
const UPSTREAM_LEAKAGE_RATE = 0.015; // 1.5% upstream leakage
const ONSITE_SLIPPAGE_RATE = 0.01; // 1.0% on-site slippage/unburned methane
const TOTAL_LEAKAGE_RATE = UPSTREAM_LEAKAGE_RATE + ONSITE_SLIPPAGE_RATE;
const KG_CH4_IN_THERM = 2.36; // Approx. kg of CH4 in 1 therm of natural gas
const GWP20_CH4 = 84; // 20-year Global Warming Potential of Methane (IPCC AR6)
const GWP100_CH4 = 28; // 100-year Global Warming Potential of Methane (IPCC AR6)
// Source: ISO New England 2022 data, for planning purposes.
const CO2_KG_PER_KWH_MA = 0.26;
const MILES_DRIVEN_PER_KG_CO2 = 2.48; // Source: EPA, Avg passenger vehicle
const KG_CO2E_PER_CAR_YEAR = 4600; // Source: EPA, avg passenger vehicle emits 4.6 metric tons CO2e/year


// --- COST DATA (for estimation purposes) ---
const EQUIPMENT_COSTS: Record<ApplianceCategory, Record<EfficiencyTier, number[]>> = {
    'Furnace': { 'Standard': [6000, 8000], 'High': [8000, 12000], 'Premium': [12000, 18000] },
    'Boiler': { 'Standard': [6000, 8000], 'High': [8000, 12000], 'Premium': [12000, 18000] }, // Cost is for the unit, not distribution
    'Water Heater': { 'Standard': [1500, 2500], 'High': [2500, 4000], 'Premium': [4000, 6000] },
    'Range': { 'Standard': [1000, 2000], 'High': [2000, 3500], 'Premium': [3500, 5000] },
    'Dryer': { 'Standard': [800, 1200], 'High': [1200, 1800], 'Premium': [1800, 2500] },
    'Space Heater': { 'Standard': [1500, 2500], 'High': [2500, 4000], 'Premium': [4000, 5500] },
};
const INSTALLATION_COSTS: Record<ApplianceCategory, number[]> = {
    'Furnace': [5000, 10000],
    'Boiler': [5000, 10000],
    'Water Heater': [1000, 2000],
    'Range': [300, 600],
    'Dryer': [300, 600],
    'Space Heater': [3000, 6000],
};
const DISTRIBUTION_SYSTEM_COSTS = {
    'Ductwork Modification': [500, 2000], // For furnace -> central HP
    'New Heat Distribution System': [4000, 10000], // For boiler -> ductless mini-split system
};
const DECOMMISSIONING_COSTS: Record<ApplianceCategory, number[]> = {
    'Furnace': [250, 500],
    'Boiler': [250, 500],
    'Water Heater': [150, 300],
    'Range': [100, 200],
    'Dryer': [0, 0], // Usually part of install
    'Space Heater': [50, 100],
};
const ELECTRICAL_UPGRADE_COSTS = {
    'New Circuit': [1000, 1500],
    'Sub-panel': [1500, 2500],
    'Panel Upgrade to 125A': [2000, 3500],
    'Panel Upgrade to 150A': [2500, 4000],
    'Panel Upgrade to 200A': [3000, 5000],
    'Panel Upgrade to 400A': [5000, 8000],
    'Panel Upgrade to 600A': [15000, 25000],
};
// Source: Mass Save, for estimation only.
const REBATES: Record<string, number> = {
    'Heat Pump': 10000, // For whole-home systems
    'HPWH': 750, // Heat Pump Water Heater
    'Panel Upgrade': 1500, // When installed with heat pump
};

// --- DOM ELEMENT SELECTORS ---
const facilityTypeEl = document.getElementById('facility-type') as HTMLSelectElement;
const climateZoneEl = document.getElementById('climate-zone') as HTMLSelectElement;
const addressNumberEl = document.getElementById('address-number') as HTMLInputElement;
const streetNameEl = document.getElementById('street-name') as HTMLInputElement;
const townEl = document.getElementById('town') as HTMLInputElement;

const annualHeatingThermsEl = document.getElementById('annual-heating-therms') as HTMLInputElement;
const annualNonHeatingThermsEl = document.getElementById('annual-non-heating-therms') as HTMLInputElement;
const panelAmperageEl = document.getElementById('panel-amperage') as HTMLSelectElement;
const voltageEl = document.getElementById('voltage') as HTMLInputElement;
const breakerSpacesEl = document.getElementById('breaker-spaces') as HTMLInputElement;
const efficiencyTierEl = document.getElementById('efficiency-tier') as HTMLSelectElement;
const gasPriceEl = document.getElementById('gas-price') as HTMLInputElement;
const electricityPriceEl = document.getElementById('electricity-price') as HTMLInputElement;
const gasPriceEscalationEl = document.getElementById('gas-price-escalation') as HTMLInputElement;
const electricityPriceEscalationEl = document.getElementById('electricity-price-escalation') as HTMLInputElement;
const highRiskGasEscalationEl = document.getElementById('high-risk-gas-escalation') as HTMLInputElement;

const appliancesListEl = document.getElementById('appliances-list') as HTMLDivElement;
const addApplianceBtn = document.getElementById('add-appliance-btn') as HTMLButtonElement;
const applianceRowTemplate = document.getElementById('appliance-row-template') as HTMLTemplateElement;

const reportDetailsEl = document.getElementById('report-details') as HTMLDivElement;
const summaryTotalCostEl = document.getElementById('summary-total-cost') as HTMLSpanElement;
const summaryTotalRebatesEl = document.getElementById('summary-total-rebates') as HTMLSpanElement;
const summaryAnnualSavingsEl = document.getElementById('summary-annual-savings') as HTMLSpanElement;
const summaryPaybackPeriodEl = document.getElementById('summary-payback-period') as HTMLSpanElement;
const summaryGhgReduction20yrEl = document.getElementById('summary-ghg-reduction-20yr') as HTMLSpanElement;
const summaryGhgReduction100yrEl = document.getElementById('summary-ghg-reduction-100yr') as HTMLSpanElement;
const summaryGhgEquivalentEl = document.getElementById('summary-ghg-equivalent') as HTMLSpanElement;
const summaryLifetimeSavingsEl = document.getElementById('summary-lifetime-savings') as HTMLSpanElement;
const summaryLifetimeSavingsHighRiskEl = document.getElementById('summary-lifetime-savings-high-risk') as HTMLSpanElement;

// Modal elements
const modalOverlayEl = document.getElementById('modal-overlay') as HTMLDivElement;
const applianceModalEl = document.getElementById('appliance-modal') as HTMLDivElement;
const modalTitleEl = document.getElementById('modal-title') as HTMLHeadingElement;
const modalCloseBtn = document.getElementById('modal-close-btn') as HTMLButtonElement;
const modalApplianceGridEl = document.getElementById('modal-appliance-grid') as HTMLDivElement;

// Tab elements
const tabPlannerBtn = document.getElementById('tab-planner') as HTMLButtonElement;
const tabReportBtn = document.getElementById('tab-report') as HTMLButtonElement;
const tabSaveLoadBtn = document.getElementById('tab-save-load') as HTMLButtonElement;
const plannerContentEl = document.getElementById('planner-content') as HTMLDivElement;
const reportContentWrapperEl = document.getElementById('report-content-wrapper') as HTMLDivElement;
const fullReportContentEl = document.getElementById('full-report-content') as HTMLDivElement;
const saveLoadContentEl = document.getElementById('save-load-content') as HTMLDivElement;

// New Save/Load/Export elements
const exportPdfBtn = document.getElementById('export-pdf-btn') as HTMLButtonElement;
const saveProjectBtn = document.getElementById('save-project-btn') as HTMLButtonElement;
const loadProjectInput = document.getElementById('load-project-input') as HTMLInputElement;
const loadedFilenameEl = document.getElementById('loaded-filename') as HTMLSpanElement;


let lifetimeChart: any | null = null;
let fullReportLifetimeChart: any | null = null;


// --- APPLICATION STATE ---
let state: State = {
  facilityType: 'Residential',
  climateZone: 'Zone5',
  addressNumber: '45',
  streetName: 'Main Street',
  town: 'Gardner',
  annualHeatingTherms: 700,
  annualNonHeatingTherms: 250,
  panelAmps: 200,
  voltage: 240,
  breakerSpaces: 4,
  efficiencyTier: 'High',
  appliances: [],
  gasPricePerTherm: 1.50,
  electricityPricePerKwh: 0.22,
  gasPriceEscalation: 3,
  electricityPriceEscalation: 2,
  highRiskGasEscalation: 7,
};

// --- CORE LOGIC ---

/**
 * Calculates the PEAK kW of an electric appliance for load calculation purposes.
 * This is based on the appliance's rated power (BTU/hr), not annual energy usage.
 */
function getPeakKw(appliance: Appliance, efficiency: EfficiencyTier): number {
  const definition = APPLIANCE_DEFINITIONS[appliance.key];
  if (!definition) return 0;
  
  const category = definition.category;
  switch (category) {
    case 'Furnace':
    case 'Boiler':
    case 'Water Heater':
    case 'Space Heater':
      // This conversion assumes the electric equivalent is sized to deliver the same peak output.
      return (appliance.btu * BTU_TO_KW) / COP_LOOKUP[category][efficiency];
    case 'Range': return APPLIANCE_KW_LOOKUP.Range ?? 0;
    case 'Dryer': return APPLIANCE_KW_LOOKUP.Dryer ?? 0;
    default: return 0;
  }
}

function getApplianceRecommendation(appliance: Appliance, efficiencyTier: EfficiencyTier): { text: string, size: string, electricType: string } {
    const definition = APPLIANCE_DEFINITIONS[appliance.key];
    const efficiencyLabel = efficiencyTier === 'High' ? 'High-Efficiency' : efficiencyTier;
    let typeName = ''; let size = ''; let electricType = '';

    switch (definition.category) {
        case 'Furnace':
            typeName = 'Central Ducted Heat Pump';
            electricType = 'Heat Pump';
            const tons = appliance.btu / 12000;
            size = `${(Math.round(tons * 2) / 2).toFixed(1)}-ton`;
            break;
        case 'Boiler':
            typeName = 'Ductless Mini-Split System';
            electricType = 'Heat Pump';
            const boilerTons = appliance.btu / 12000;
            size = `${(Math.round(boilerTons * 2) / 2).toFixed(1)}-ton`;
            break;
        case 'Water Heater':
            typeName = 'HPWH';
            electricType = 'HPWH';
            if (appliance.btu <= 40000) size = '50-gallon';
            else if (appliance.btu <= 60000) size = '65-gallon';
            else size = '80-gallon';
            break;
        case 'Space Heater':
            typeName = 'Ductless Mini-Split';
            electricType = 'Heat Pump'; // Qualifies for HP rebate
            const spaceHeaterTons = appliance.btu / 12000;
            size = `${(Math.round(spaceHeaterTons * 2) / 2).toFixed(1)}-ton`;
            break;
        case 'Range': typeName = 'Induction Range'; electricType = 'Range'; size = 'Standard'; break;
        case 'Dryer': typeName = 'Heat Pump Dryer'; electricType = 'Dryer'; size = 'Standard'; break;
    }
    const text = `<strong>${size} ${efficiencyLabel} ${typeName}</strong>`;
    return { text, size, electricType };
}

function calculateAnalysis(includedAppliances: Appliance[], currentState: State): PlanningAnalysis {
    const recommendations: ApplianceRecommendation[] = includedAppliances.map(app => {
        const kw = getPeakKw(app, currentState.efficiencyTier);
        const voltage = currentState.voltage > 0 ? currentState.voltage : 240;
        const amps = (kw * 1000) / voltage;
        const definition = APPLIANCE_DEFINITIONS[app.key];
        const breakerSpaces = BREAKER_SPACE_REQUIREMENTS[definition.category];
        const { text, size } = getApplianceRecommendation(app, currentState.efficiencyTier);

        return { id: app.id, text, size, kw, amps, breakerSpaces };
    });

    const totalNewAmps = recommendations.reduce((sum, r) => sum + r.amps, 0);
    const requiredBreakerSpaces = recommendations.reduce((sum, r) => sum + r.breakerSpaces, 0);

    const assumedExistingLoad = currentState.panelAmps * 0.8 * 0.5;
    const totalCalculatedLoad = assumedExistingLoad + totalNewAmps;
    const panelCapacity = currentState.panelAmps * 0.8;

    let panelStatus: PlanningAnalysis['panelStatus'] = 'Not Required';
    if (totalCalculatedLoad > panelCapacity) {
        const requiredPanelSize = totalCalculatedLoad / 0.8;
        const recommendedPanel = PANEL_SIZES.find(size => size >= requiredPanelSize) || PANEL_SIZES.slice(-1)[0];
        panelStatus = `Upgrade to ${recommendedPanel}A Recommended`;
    }

    let breakerStatus: PlanningAnalysis['breakerStatus'] = 'Sufficient';
    if (panelStatus !== 'Not Required') breakerStatus = 'Panel Upgrade Recommended';
    else if (requiredBreakerSpaces > currentState.breakerSpaces) breakerStatus = 'Sub-panel Recommended';

    return { recommendations, totalNewAmps, requiredBreakerSpaces, breakerStatus, panelStatus };
}

function calculateRebates(analysis: PlanningAnalysis, includedAppliances: Appliance[], state: State): RebateItem[] {
    const rebates: RebateItem[] = [];
    if (includedAppliances.length === 0) return rebates;

    const recommendedApplianceTypes = new Set(
        analysis.recommendations.map(rec => {
            const appliance = state.appliances.find(a => a.id === rec.id)!;
            return getApplianceRecommendation(appliance, state.efficiencyTier).electricType;
        })
    );

    if (recommendedApplianceTypes.has('Heat Pump') && REBATES['Heat Pump']) {
        rebates.push({ name: 'Heat Pump Rebate', amount: REBATES['Heat Pump'] });
    }
    if (recommendedApplianceTypes.has('HPWH') && REBATES['HPWH']) {
        rebates.push({ name: 'HPWH Rebate', amount: REBATES['HPWH'] });
    }
    if (analysis.panelStatus.startsWith('Upgrade') && recommendedApplianceTypes.has('Heat Pump') && REBATES['Panel Upgrade']) {
         rebates.push({ name: 'Panel Upgrade Rebate', amount: REBATES['Panel Upgrade'] });
    }

    return rebates;
}

function calculateCostAnalysis(analysis: PlanningAnalysis, includedAppliances: Appliance[], state: State): CostAnalysis {
    const applianceCosts = new Map<string, CostItem[]>();
    let totalLow = 0, totalHigh = 0;

    analysis.recommendations.forEach(rec => {
        const appliance = state.appliances.find(a => a.id === rec.id)!;
        const definition = APPLIANCE_DEFINITIONS[appliance.key];
        const category = definition.category;
        const costs: CostItem[] = [];

        const [equipLow, equipHigh] = EQUIPMENT_COSTS[category][state.efficiencyTier];
        costs.push({ name: 'Equipment', low: equipLow, high: equipHigh });

        const [installLow, installHigh] = INSTALLATION_COSTS[category];
        costs.push({ name: 'Installation & Materials', low: installLow, high: installHigh });

        const [decomLow, decomHigh] = DECOMMISSIONING_COSTS[category];
        if (decomHigh > 0) {
            costs.push({ name: 'Gas Decommissioning', low: decomLow, high: decomHigh });
        }

        // Add costs for heat distribution system based on original appliance
        if (category === 'Furnace') {
            const [low, high] = DISTRIBUTION_SYSTEM_COSTS['Ductwork Modification'];
            costs.push({ name: 'Ductwork Modification', low, high });
        } else if (category === 'Boiler') {
            const [low, high] = DISTRIBUTION_SYSTEM_COSTS['New Heat Distribution System'];
            costs.push({ name: 'New Heat Distribution System', low, high });
        }

        applianceCosts.set(rec.id, costs);
    });

    const electricalCosts: CostItem[] = [];
    if (analysis.panelStatus.startsWith('Upgrade')) {
        const size = analysis.panelStatus.match(/\d+/)?.[0];
        const key = `Panel Upgrade to ${size}A` as keyof typeof ELECTRICAL_UPGRADE_COSTS;
        if (ELECTRICAL_UPGRADE_COSTS[key]) {
            const [low, high] = ELECTRICAL_UPGRADE_COSTS[key];
            electricalCosts.push({ name: 'Main Panel Upgrade', low, high });
        }
    } else if (analysis.breakerStatus === 'Sub-panel Recommended') {
        const [low, high] = ELECTRICAL_UPGRADE_COSTS['Sub-panel'];
        electricalCosts.push({ name: 'Sub-panel Installation', low, high });
    }

    if (analysis.recommendations.length > 0) {
        const [low, high] = ELECTRICAL_UPGRADE_COSTS['New Circuit'];
        const numCircuits = analysis.recommendations.length;
        electricalCosts.push({ name: `${numCircuits} New Circuit(s)`, low: low * numCircuits, high: high * numCircuits });
    }

    // Sum totals
    for(const costs of applianceCosts.values()) {
        costs.forEach(c => { totalLow += c.low; totalHigh += c.high; });
    }
    electricalCosts.forEach(c => { totalLow += c.low; totalHigh += c.high; });

    // Calculate rebates
    const rebates = calculateRebates(analysis, includedAppliances, state);
    const totalRebates = rebates.reduce((sum, r) => sum + r.amount, 0);

    const netLow = Math.max(0, totalLow - totalRebates);
    const netHigh = Math.max(0, totalHigh - totalRebates);

    return { applianceCosts, electricalCosts, rebates, totalLow, totalHigh, totalRebates, netLow, netHigh };
}

/**
 * Calculates therms attributable to a selection of appliances by scaling total
 * usage based on the selection's share of total listed BTU capacity.
 */
function getThermsForSelection(includedAppliances: Appliance[], allAppliances: Appliance[], state: State): { heatingTherms: number, nonHeatingTherms: number } {
    const heatingCategories: ApplianceCategory[] = ['Furnace', 'Boiler', 'Space Heater'];
    const nonHeatingCategories: ApplianceCategory[] = ['Water Heater', 'Range', 'Dryer'];

    const includedHeatingAppliances = includedAppliances.filter(a => heatingCategories.includes(APPLIANCE_DEFINITIONS[a.key].category));
    const includedNonHeatingAppliances = includedAppliances.filter(a => nonHeatingCategories.includes(APPLIANCE_DEFINITIONS[a.key].category));
    
    const allHeatingAppliances = allAppliances.filter(a => heatingCategories.includes(APPLIANCE_DEFINITIONS[a.key].category));
    const allNonHeatingAppliances = allAppliances.filter(a => nonHeatingCategories.includes(APPLIANCE_DEFINITIONS[a.key].category));

    const totalIncludedHeatingBtu = includedHeatingAppliances.reduce((sum, a) => sum + a.btu, 0);
    const totalAllHeatingBtu = allHeatingAppliances.reduce((sum, a) => sum + a.btu, 0);
    
    const totalIncludedNonHeatingBtu = includedNonHeatingAppliances.reduce((sum, a) => sum + a.btu, 0);
    const totalAllNonHeatingBtu = allNonHeatingAppliances.reduce((sum, a) => sum + a.btu, 0);

    const heatingTherms = totalAllHeatingBtu > 0
        ? state.annualHeatingTherms * (totalIncludedHeatingBtu / totalAllHeatingBtu)
        : 0;
    
    const nonHeatingTherms = totalAllNonHeatingBtu > 0
        ? state.annualNonHeatingTherms * (totalIncludedNonHeatingBtu / totalAllNonHeatingBtu)
        : 0;

    return { heatingTherms, nonHeatingTherms };
}

function calculateProjectedAnnualKwh(includedAppliances: Appliance[], state: State): number {
    const { heatingTherms, nonHeatingTherms } = getThermsForSelection(includedAppliances, state.appliances, state);
    
    const heatingCategories: ApplianceCategory[] = ['Furnace', 'Boiler', 'Space Heater'];
    const heatingAppliances = includedAppliances.filter(a => heatingCategories.includes(APPLIANCE_DEFINITIONS[a.key].category));
    const nonHeatingAppliances = includedAppliances.filter(a => !heatingCategories.includes(APPLIANCE_DEFINITIONS[a.key].category));

    const totalHeatingBtuRating = heatingAppliances.reduce((sum, a) => sum + a.btu, 0);
    const totalNonHeatingBtuRating = nonHeatingAppliances.reduce((sum, a) => sum + a.btu, 0);

    let totalProjectedKwh = 0;

    for (const app of includedAppliances) {
        let thermsForAppliance = 0;
        const definition = APPLIANCE_DEFINITIONS[app.key];
        const isHeating = heatingCategories.includes(definition.category);

        if (isHeating && totalHeatingBtuRating > 0) {
            thermsForAppliance = heatingTherms * (app.btu / totalHeatingBtuRating);
        } else if (!isHeating && totalNonHeatingBtuRating > 0) {
            thermsForAppliance = nonHeatingTherms * (app.btu / totalNonHeatingBtuRating);
        }

        if (thermsForAppliance > 0) {
            // Useful heat delivered by the gas appliance, accounting for its efficiency
            const usefulBtuOutput = thermsForAppliance * BTU_PER_THERM * (app.efficiency / 100);
            
            let cop = COP_LOOKUP[definition.category][state.efficiencyTier];
            if (isHeating) {
                cop *= CLIMATE_ZONE_COP_ADJUSTMENT[state.climateZone];
            }
            
            // Kwh required for electric appliance to deliver the same useful heat
            const kwhInput = (usefulBtuOutput * BTU_TO_KW) / cop;
            totalProjectedKwh += kwhInput;
        }
    }
    return totalProjectedKwh;
}


function calculateFinancialAnalysis(costAnalysis: CostAnalysis, includedAppliances: Appliance[], state: State, projectedAnnualKwh: number): FinancialAnalysis {
    const { heatingTherms, nonHeatingTherms } = getThermsForSelection(includedAppliances, state.appliances, state);
    const totalTherms = heatingTherms + nonHeatingTherms;
    const currentAnnualGasCost = totalTherms * state.gasPricePerTherm;
    const projectedAnnualElecCost = projectedAnnualKwh * state.electricityPricePerKwh;
    const netAnnualSavings = currentAnnualGasCost - projectedAnnualElecCost;

    let simplePaybackYears: number | null = null;
    if (netAnnualSavings > 0 && costAnalysis.netLow > 0) {
        const avgNetCost = (costAnalysis.netLow + costAnalysis.netHigh) / 2;
        simplePaybackYears = avgNetCost / netAnnualSavings;
    }

    return { currentAnnualGasCost, projectedAnnualElecCost, netAnnualSavings, simplePaybackYears };
}

function calculateEnvironmentalImpact(includedAppliances: Appliance[], state: State, projectedAnnualKwh: number): EnvironmentalImpact {
    const { heatingTherms, nonHeatingTherms } = getThermsForSelection(includedAppliances, state.appliances, state);
    const totalTherms = heatingTherms + nonHeatingTherms;

    // Current gas emissions
    const currentAnnualGasCo2Kg = totalTherms * CO2_KG_PER_THERM_GAS;
    const currentAnnualCh4Kg = totalTherms * KG_CH4_IN_THERM * TOTAL_LEAKAGE_RATE;
    const currentAnnualGasCo2e20Kg = currentAnnualGasCo2Kg + (currentAnnualCh4Kg * GWP20_CH4);
    const currentAnnualGasCo2e100Kg = currentAnnualGasCo2Kg + (currentAnnualCh4Kg * GWP100_CH4);

    // Projected electric emissions
    const projectedAnnualElecCo2Kg = projectedAnnualKwh * CO2_KG_PER_KWH_MA;

    // Reductions
    const annualGhgReductionCo2e20Kg = currentAnnualGasCo2e20Kg - projectedAnnualElecCo2Kg;
    const annualGhgReductionCo2e100Kg = currentAnnualGasCo2e100Kg - projectedAnnualElecCo2Kg;

    // Equivalency
    const ghgReductionCarsOffRoad100yr = annualGhgReductionCo2e100Kg / KG_CO2E_PER_CAR_YEAR;

    return { 
        currentAnnualGasCo2Kg, 
        currentAnnualCh4Kg,
        currentAnnualGasCo2e20Kg,
        currentAnnualGasCo2e100Kg,
        projectedAnnualElecCo2Kg,
        annualGhgReductionCo2e20Kg,
        annualGhgReductionCo2e100Kg,
        ghgReductionCarsOffRoad100yr,
    };
}

function calculateLifetimeAnalysis(financialAnalysis: FinancialAnalysis, costAnalysis: CostAnalysis, state: State): LifetimeAnalysis {
    const cumulativeGasCosts: number[] = [0];
    const cumulativeHighRiskGasCosts: number[] = [0];
    const cumulativeElecCosts: number[] = [(costAnalysis.netLow + costAnalysis.netHigh) / 2];

    let currentYearGasCost = financialAnalysis.currentAnnualGasCost;
    let currentYearHighRiskGasCost = financialAnalysis.currentAnnualGasCost;
    let currentYearElecCost = financialAnalysis.projectedAnnualElecCost;

    const gasEscalationRate = 1 + (state.gasPriceEscalation / 100);
    const highRiskGasEscalationRate = 1 + (state.highRiskGasEscalation / 100);
    const elecEscalationRate = 1 + (state.electricityPriceEscalation / 100);

    for (let i = 0; i < LIFETIME_YEARS; i++) {
        cumulativeGasCosts.push(cumulativeGasCosts[i] + currentYearGasCost);
        cumulativeHighRiskGasCosts.push(cumulativeHighRiskGasCosts[i] + currentYearHighRiskGasCost);
        cumulativeElecCosts.push(cumulativeElecCosts[i] + currentYearElecCost);
        
        currentYearGasCost *= gasEscalationRate;
        currentYearHighRiskGasCost *= highRiskGasEscalationRate;
        currentYearElecCost *= elecEscalationRate;
    }
    
    // Total electric cost for savings calc is operating cost, not including upfront.
    const totalElecOperatingCost15yr = cumulativeElecCosts[LIFETIME_YEARS] - cumulativeElecCosts[0];

    const totalGasCost15yr = cumulativeGasCosts[LIFETIME_YEARS];
    const totalSavings15yr = totalGasCost15yr - totalElecOperatingCost15yr;

    const totalHighRiskGasCost15yr = cumulativeHighRiskGasCosts[LIFETIME_YEARS];
    const totalHighRiskSavings15yr = totalHighRiskGasCost15yr - totalElecOperatingCost15yr;


    return { 
        totalGasCost15yr, 
        totalElecCost15yr: totalElecOperatingCost15yr, 
        totalSavings15yr, 
        cumulativeGasCosts, 
        cumulativeElecCosts,
        totalHighRiskSavings15yr,
        cumulativeHighRiskGasCosts,
    };
}

function calculateDistributionImpact(planningAnalysis: PlanningAnalysis, state: State): DistributionImpactAnalysis {
    let peakDemandKw = 0;
    let nonHeatingDemandKw = 0;

    const nonHeatingCategories: ApplianceCategory[] = ['Water Heater', 'Range', 'Dryer'];

    planningAnalysis.recommendations.forEach(rec => {
        peakDemandKw += rec.kw;

        const appliance = state.appliances.find(a => a.id === rec.id);
        if (appliance) {
            const definition = APPLIANCE_DEFINITIONS[appliance.key];
            if (nonHeatingCategories.includes(definition.category)) {
                nonHeatingDemandKw += rec.kw;
            }
        }
    });

    return { peakDemandKw, nonHeatingDemandKw };
}


// --- RENDERING & UI ---
function formatCurrency(value: number, showSign = false): string {
    const sign = value > 0 && showSign ? '+' : '';
    return sign + new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(value);
}

function createLifetimeChart(canvasId: string, analysis: LifetimeAnalysis, chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;

    if (chartInstance) {
        chartInstance.destroy();
    }

    const labels = Array.from({ length: LIFETIME_YEARS + 1 }, (_, i) => `Year ${i}`);

    return new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Cumulative Electric Cost (incl. install)',
                    data: analysis.cumulativeElecCosts,
                    borderColor: 'rgb(0, 95, 204)',
                    backgroundColor: 'rgba(0, 95, 204, 0.1)',
                    fill: true,
                    tension: 0.1,
                },
                {
                    label: 'Cumulative Gas Cost (projected)',
                    data: analysis.cumulativeGasCosts,
                    borderColor: 'rgb(220, 53, 69)',
                    backgroundColor: 'rgba(220, 53, 69, 0.1)',
                    fill: true,
                    tension: 0.1,
                },
                {
                    label: 'Cumulative Gas Cost (High-Risk)',
                    data: analysis.cumulativeHighRiskGasCosts,
                    borderColor: 'rgb(220, 53, 69)',
                    borderDash: [5, 5],
                    backgroundColor: 'rgba(220, 53, 69, 0.05)',
                    fill: true,
                    tension: 0.1,
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                tooltip: {
                    callbacks: {
                        label: function(context: any) {
                            let label = context.dataset.label || '';
                            if (label) {
                                label += ': ';
                            }
                            if (context.parsed.y !== null) {
                                label += formatCurrency(context.parsed.y);
                            }
                            return label;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value: any) {
                            return formatCurrency(value);
                        }
                    }
                }
            }
        }
    });
}


function renderReport(planningAnalysis: PlanningAnalysis, costAnalysis: CostAnalysis, financialAnalysis: FinancialAnalysis, environmentalImpact: EnvironmentalImpact, lifetimeAnalysis: LifetimeAnalysis) {
    // Summary
    summaryTotalCostEl.textContent = `${formatCurrency(costAnalysis.netLow)} - ${formatCurrency(costAnalysis.netHigh)}`;
    summaryTotalRebatesEl.textContent = formatCurrency(costAnalysis.totalRebates);
    summaryAnnualSavingsEl.textContent = formatCurrency(financialAnalysis.netAnnualSavings, true);
    summaryAnnualSavingsEl.className = 'summary-metric-value';
    if(financialAnalysis.netAnnualSavings > 0) summaryAnnualSavingsEl.classList.add('status-sufficient');
    else summaryAnnualSavingsEl.classList.add('status-danger');

    const ghgReduction20yr = environmentalImpact.annualGhgReductionCo2e20Kg;
    summaryGhgReduction20yrEl.textContent = `${ghgReduction20yr.toLocaleString(undefined, {maximumFractionDigits: 0})} kg CO₂e`;
    summaryGhgReduction20yrEl.className = 'summary-metric-value';
    if(ghgReduction20yr > 0) summaryGhgReduction20yrEl.classList.add('status-sufficient');
    else if (ghgReduction20yr < 0) summaryGhgReduction20yrEl.classList.add('status-danger');

    const ghgReduction100yr = environmentalImpact.annualGhgReductionCo2e100Kg;
    summaryGhgReduction100yrEl.textContent = `${ghgReduction100yr.toLocaleString(undefined, {maximumFractionDigits: 0})} kg CO₂e`;
    summaryGhgReduction100yrEl.className = 'summary-metric-value';
    if(ghgReduction100yr > 0) summaryGhgReduction100yrEl.classList.add('status-sufficient');
    else if (ghgReduction100yr < 0) summaryGhgReduction100yrEl.classList.add('status-danger');

    const ghgEquivalent = environmentalImpact.ghgReductionCarsOffRoad100yr;
    summaryGhgEquivalentEl.textContent = `${ghgEquivalent.toLocaleString(undefined, {maximumFractionDigits: 1})} cars`;
    summaryGhgEquivalentEl.className = 'summary-metric-value';
    if(ghgEquivalent > 0) summaryGhgEquivalentEl.classList.add('status-sufficient');
    else if (ghgEquivalent < 0) summaryGhgEquivalentEl.classList.add('status-danger');


    summaryLifetimeSavingsEl.textContent = formatCurrency(lifetimeAnalysis.totalSavings15yr);
    summaryLifetimeSavingsEl.className = 'summary-metric-value';
    if(lifetimeAnalysis.totalSavings15yr > 0) summaryLifetimeSavingsEl.classList.add('status-sufficient');
    else summaryLifetimeSavingsEl.classList.add('status-danger');

    summaryLifetimeSavingsHighRiskEl.textContent = formatCurrency(lifetimeAnalysis.totalHighRiskSavings15yr);
    summaryLifetimeSavingsHighRiskEl.className = 'summary-metric-value';
    if(lifetimeAnalysis.totalHighRiskSavings15yr > 0) summaryLifetimeSavingsHighRiskEl.classList.add('status-sufficient');
    else summaryLifetimeSavingsHighRiskEl.classList.add('status-danger');


    if (financialAnalysis.simplePaybackYears !== null) {
        summaryPaybackPeriodEl.textContent = `${financialAnalysis.simplePaybackYears.toFixed(1)} years`;
    } else {
        summaryPaybackPeriodEl.textContent = 'N/A';
    }

    // Details
    reportDetailsEl.innerHTML = '';
    const includedAppliances = state.appliances.filter(a => a.included);
    if (includedAppliances.length === 0) {
        reportDetailsEl.innerHTML = '<p class="empty-state">Select one or more appliances in the "Plan" column to generate a report.</p>';
        if (fullReportLifetimeChart) fullReportLifetimeChart.destroy();
        if (lifetimeChart) lifetimeChart.destroy();
        return;
    }

    // Regulatory Context
    const contextSection = document.createElement('div');
    contextSection.className = 'report-section';
    contextSection.innerHTML = `<h3 class="report-section-title">Regulatory Context</h3>`;
    contextSection.innerHTML += `<p>The Massachusetts DPU is managing the transition away from natural gas to meet climate goals (Docket 20-80-B). This influences future utility costs and infrastructure decisions.</p>`;
    reportDetailsEl.appendChild(contextSection);

    // Lifetime Financial Outlook
    const lifetimeSection = document.createElement('div');
    lifetimeSection.className = 'report-section';
    lifetimeSection.innerHTML = `<h3 class="report-section-title">15-Year Financial Outlook</h3>`;
    lifetimeSection.innerHTML += `
        <p>
            Assumes a ${state.gasPriceEscalation}% annual increase in gas prices and a ${state.electricityPriceEscalation}% increase in electricity prices.
        </p>
        <div class="chart-container">
            <canvas id="lifetime-chart"></canvas>
        </div>
    `;
    reportDetailsEl.appendChild(lifetimeSection);
    lifetimeChart = createLifetimeChart('lifetime-chart', lifetimeAnalysis, lifetimeChart);

     // Next Steps
    const nextStepsSection = document.createElement('div');
    nextStepsSection.className = 'report-section';
    nextStepsSection.innerHTML = `
        <h3 class="report-section-title">Next Steps</h3>
        <ol class="next-steps-list">
            <li><strong>Get Professional Quotes:</strong> Contact licensed electricians and HVAC professionals for detailed, binding quotes.</li>
            <li><strong>Explore Incentives:</strong> Visit the <a href="https://www.masssave.com/rebates" target="_blank" rel="noopener noreferrer">Mass Save website</a> for available rebates.</li>
            <li><strong>Consider a Home Energy Audit:</strong> An audit can identify other efficiency opportunities.</li>
        </ol>
    `;
    reportDetailsEl.appendChild(nextStepsSection);
}

function renderFullReport(planningAnalysis: PlanningAnalysis, costAnalysis: CostAnalysis, financialAnalysis: FinancialAnalysis, environmentalImpact: EnvironmentalImpact, lifetimeAnalysis: LifetimeAnalysis, distributionImpact: DistributionImpactAnalysis) {
    const includedAppliances = state.appliances.filter(a => a.included);
    if (includedAppliances.length === 0) {
        fullReportContentEl.innerHTML = '<p class="empty-state">Select one or more appliances on the Planner tab to generate a full report.</p>';
        return;
    }
    
    const address = [state.addressNumber, state.streetName, state.town].filter(Boolean).join(' ');
    const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    const climateZoneText = state.climateZone === 'Zone5' ? "Zone 5 (Most of MA)" : "Zone 6 (Berkshires)";

    let inputsHtml = `
        <table class="full-report-table">
            <tr><th>Facility Type</th><td>${state.facilityType}</td></tr>
            <tr><th>Climate Zone</th><td>${climateZoneText}</td></tr>
            <tr><th>Heating Gas Usage</th><td>${state.annualHeatingTherms.toLocaleString()} therms/yr</td></tr>
            <tr><th>Non-Heating Gas Usage</th><td>${state.annualNonHeatingTherms.toLocaleString()} therms/yr</td></tr>
            <tr><th>Main Panel</th><td>${state.panelAmps}A @ ${state.voltage}V</td></tr>
            <tr><th>Available Breaker Spaces</th><td>${state.breakerSpaces}</td></tr>
            <tr><th>Efficiency Target</th><td>${state.efficiencyTier}</td></tr>
            <tr><th>Gas Price</th><td>${formatCurrency(state.gasPricePerTherm)} / therm</td></tr>
            <tr><th>Electricity Price</th><td>${formatCurrency(state.electricityPricePerKwh, false)} / kWh</td></tr>
        </table>
    `;

    const replacedCategories = new Set(planningAnalysis.recommendations.map(rec => {
        const appliance = state.appliances.find(a => a.id === rec.id)!;
        return APPLIANCE_DEFINITIONS[appliance.key].category;
    }));
    const hasFurnace = replacedCategories.has('Furnace');
    const hasBoiler = replacedCategories.has('Boiler');

    let planHtml = `
        <table class="full-report-table">
            <thead><tr><th>Gas Appliance</th><th>BTU/hr</th><th>Efficiency</th><th>Electric Replacement</th></tr></thead>
            <tbody>
            ${planningAnalysis.recommendations.map(rec => {
                const appliance = state.appliances.find(a => a.id === rec.id)!;
                const definition = APPLIANCE_DEFINITIONS[appliance.key];
                return `<tr><td>${definition.name}</td><td>${appliance.btu.toLocaleString()}</td><td>${appliance.efficiency}%</td><td>${rec.text.replace(/<strong>/g, '').replace(/<\/strong>/g, '')}</td></tr>`;
            }).join('')}
            </tbody>
        </table>
    `;

    if (hasFurnace) {
        planHtml += `<p class="report-note"><strong>Note on Furnaces:</strong> Ductwork condition and sizing must be field-verified. If replacement or major modifications are required, add contingency cost of $30,000–$50,000.</p>`;
    }
    if (hasBoiler) {
        planHtml += `<p class="report-note"><strong>Note on Boilers:</strong> Replacing a boiler (which uses hot water/steam pipes) with a standard heat pump (which produces hot air) requires a new heat distribution system. The cost estimate includes the installation of a ductless mini-split system, a common solution. Other options like air-to-water heat pumps may be possible and require consultation with an HVAC professional.</p>`;
    }
    
    let electricalHtml = `
        <p>The total new electrical load is estimated based on the peak power demand of the new appliances.</p>
        <div class="formula">$$Amps = \\frac{kW \\times 1000}{Volts}$$</div>
        <table class="full-report-table">
             <thead><tr><th>Appliance</th><th>Peak kW</th><th>Est. Amps @ ${state.voltage}V</th><th>Breaker Spaces</th></tr></thead>
             <tbody>
                ${planningAnalysis.recommendations.map(rec => `<tr><td>${rec.text.replace(/<strong>/g, '').replace(/<\/strong>/g, '')}</td><td>${rec.kw.toFixed(2)} kW</td><td>${rec.amps.toFixed(1)} A</td><td>${rec.breakerSpaces}</td></tr>`).join('')}
                <tr class="total-row"><td><strong>Total</strong></td><td></td><td><strong>${planningAnalysis.totalNewAmps.toFixed(1)} A</strong></td><td><strong>${planningAnalysis.requiredBreakerSpaces}</strong></td></tr>
             </tbody>
        </table>
        <p style="margin-top: 1rem;"><strong>Breaker Status:</strong> ${planningAnalysis.breakerStatus}<br><strong>Panel Status:</strong> ${planningAnalysis.panelStatus}</p>
    `;

    let distributionHtml = `
        <p>This analysis estimates the project's impact on the local electric distribution system, which is crucial for utility planning as more buildings electrify.</p>
        <table class="full-report-table">
             <thead><tr><th>Metric</th><th>Value</th></tr></thead>
             <tbody>
                <tr>
                    <td>
                        <strong>Peak Demand Increase</strong>
                        <p style="font-size: 0.85em; color: var(--text-muted); margin: 0.5rem 0 0 0;">The maximum potential load added to the grid if all new appliances run simultaneously. This informs capacity planning for extreme weather events.</p>
                    </td>
                    <td style="vertical-align: middle; text-align: center; font-weight: 600;">${distributionImpact.peakDemandKw.toFixed(1)} kW</td>
                </tr>
                <tr>
                    <td>
                        <strong>Non-Heating Demand Increase</strong>
                        <p style="font-size: 0.85em; color: var(--text-muted); margin: 0.5rem 0 0 0;">The baseline load increase from appliances used year-round (e.g., water heaters, ranges). This informs planning for base load grid capacity.</p>
                    </td>
                    <td style="vertical-align: middle; text-align: center; font-weight: 600;">${distributionImpact.nonHeatingDemandKw.toFixed(1)} kW</td>
                </tr>
             </tbody>
        </table>
    `;
    
    // Methane breakdown for environmental report
    const { heatingTherms, nonHeatingTherms } = getThermsForSelection(includedAppliances, state.appliances, state);
    const totalTherms = heatingTherms + nonHeatingTherms;
    const upstreamCh4Kg = totalTherms * KG_CH4_IN_THERM * UPSTREAM_LEAKAGE_RATE;
    const onsiteCh4Kg = totalTherms * KG_CH4_IN_THERM * ONSITE_SLIPPAGE_RATE;

    const environmentalHtml = `
        <p>This analysis includes direct carbon dioxide (CO₂) emissions from gas combustion and accounts for methane (CH₄) leakage from the gas supply chain and on-site appliance slippage. Methane is a potent greenhouse gas (GHG), and its impact is measured in CO₂ equivalent (CO₂e) using its Global Warming Potential (GWP) over different time horizons.</p>
        <div class="formula">$$Total\\ CO_2e = (CO_{2\\ combustion}) + (CH_{4\\ leaked} \\times GWP_{CH_4})$$</div>
        <table class="full-report-table">
            <thead>
                <tr><th>Impact Horizon</th><th>Annual Gas Emissions (kg CO₂e)</th><th>Annual Electric Emissions (kg CO₂e)</th><th>Net Annual Reduction (kg CO₂e)</th></tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>20-Year Impact (GWP₂₀ for CH₄ = ${GWP20_CH4})</strong></td>
                    <td>${environmentalImpact.currentAnnualGasCo2e20Kg.toLocaleString(undefined, {maximumFractionDigits: 0})}</td>
                    <td rowspan="2" style="vertical-align: middle; text-align: center;">${environmentalImpact.projectedAnnualElecCo2Kg.toLocaleString(undefined, {maximumFractionDigits: 0})}</td>
                    <td class="total-row"><strong>${environmentalImpact.annualGhgReductionCo2e20Kg.toLocaleString(undefined, {maximumFractionDigits: 0})}</strong></td>
                </tr>
                <tr>
                    <td><strong>100-Year Impact (GWP₁₀₀ for CH₄ = ${GWP100_CH4})</strong></td>
                    <td>${environmentalImpact.currentAnnualGasCo2e100Kg.toLocaleString(undefined, {maximumFractionDigits: 0})}</td>
                    <td class="total-row"><strong>${environmentalImpact.annualGhgReductionCo2e100Kg.toLocaleString(undefined, {maximumFractionDigits: 0})}</strong></td>
                </tr>
            </tbody>
        </table>
        <p class="report-note">
            This project's annual greenhouse gas reduction is equivalent to taking <strong>${environmentalImpact.ghgReductionCarsOffRoad100yr.toFixed(1)}</strong> cars off the road for a year.
        </p>
        <p style="margin-top: 1rem; font-size: 0.9em; color: var(--text-muted);">
            Gas emissions breakdown: <strong>${environmentalImpact.currentAnnualGasCo2Kg.toLocaleString(undefined, {maximumFractionDigits: 0})} kg of CO₂</strong> from combustion, <strong>${upstreamCh4Kg.toLocaleString(undefined, {maximumFractionDigits: 1})} kg of CH₄</strong> from upstream supply chain leaks, and <strong>${onsiteCh4Kg.toLocaleString(undefined, {maximumFractionDigits: 1})} kg of CH₄</strong> from on-site appliance slippage (unburned gas).
        </p>
    `;

    // --- Start of New Cost Analysis Rendering ---
    const categorizedCosts = new Map<string, CostItem[]>();
    const categoryOrder = [
        'Equipment',
        'Installation & Materials',
        'Heat Distribution System',
        'Decommissioning',
        'Electrical Upgrades'
    ];

    const addToCategory = (category: string, item: CostItem) => {
        if (!categorizedCosts.has(category)) {
            categorizedCosts.set(category, []);
        }
        categorizedCosts.get(category)!.push(item);
    };

    costAnalysis.applianceCosts.forEach((items, id) => {
        const rec = planningAnalysis.recommendations.find(r => r.id === id);
        if (!rec) return;

        const recommendationText = rec.text.replace(/<strong>/g, '').replace(/<\/strong>/g, '');

        items.forEach(item => {
            const newItemWithName = { ...item, name: recommendationText };
            if (item.name === 'Equipment') {
                addToCategory('Equipment', newItemWithName);
            } else if (item.name === 'Installation & Materials') {
                addToCategory('Installation & Materials', newItemWithName);
            } else if (item.name.includes('Distribution System') || item.name.includes('Ductwork')) {
                addToCategory('Heat Distribution System', newItemWithName);
            } else if (item.name === 'Gas Decommissioning') {
                addToCategory('Decommissioning', newItemWithName);
            }
        });
    });

    if (costAnalysis.electricalCosts.length > 0) {
        categorizedCosts.set('Electrical Upgrades', costAnalysis.electricalCosts);
    }
    
    let costTableBodyHtml = '';
    categoryOrder.forEach(category => {
        if (categorizedCosts.has(category)) {
            const items = categorizedCosts.get(category)!;
            if (items.length === 0) return;

            // First row for the category with rowspan
            costTableBodyHtml += `
                <tr>
                    <td rowspan="${items.length}" class="category-cell"><strong>${category}</strong></td>
                    <td>${items[0].name}</td>
                    <td style="text-align: right;">${formatCurrency(items[0].low)}</td>
                    <td style="text-align: right;">${formatCurrency(items[0].high)}</td>
                </tr>
            `;
            // Subsequent rows
            for (let i = 1; i < items.length; i++) {
                costTableBodyHtml += `
                    <tr>
                        <td>${items[i].name}</td>
                        <td style="text-align: right;">${formatCurrency(items[i].low)}</td>
                        <td style="text-align: right;">${formatCurrency(items[i].high)}</td>
                    </tr>
                `;
            }
        }
    });

    let costTableFooterHtml = `
        <tr class="total-row">
            <td colspan="2"><strong>Subtotal</strong></td>
            <td style="text-align: right;"><strong>${formatCurrency(costAnalysis.totalLow)}</strong></td>
            <td style="text-align: right;"><strong>${formatCurrency(costAnalysis.totalHigh)}</strong></td>
        </tr>
    `;

    if (costAnalysis.rebates.length > 0) {
        const items = costAnalysis.rebates;
        costTableFooterHtml += `
            <tr>
                <td rowspan="${items.length}" class="category-cell"><strong>Rebates</strong></td>
                <td>${items[0].name}</td>
                <td style="text-align: right;">(${formatCurrency(items[0].amount)})</td>
                <td style="text-align: right;">(${formatCurrency(items[0].amount)})</td>
            </tr>
        `;
        for (let i = 1; i < items.length; i++) {
            costTableFooterHtml += `
                <tr>
                    <td>${items[i].name}</td>
                    <td style="text-align: right;">(${formatCurrency(items[i].amount)})</td>
                    <td style="text-align: right;">(${formatCurrency(items[i].amount)})</td>
                </tr>
            `;
        }
    }

    costTableFooterHtml += `
        <tr class="total-row">
            <td colspan="2"><strong>Net Project Cost</strong></td>
            <td style="text-align: right;"><strong>${formatCurrency(costAnalysis.netLow)}</strong></td>
            <td style="text-align: right;"><strong>${formatCurrency(costAnalysis.netHigh)}</strong></td>
        </tr>
    `;

    let costHtml = `
        <p>This estimate covers equipment, installation, electrical work, and necessary changes to the heat distribution system. All costs are for planning purposes and should be confirmed with professional quotes.</p>
        <table class="full-report-table">
            <thead>
                <tr>
                    <th>Category</th>
                    <th>Item</th>
                    <th style="text-align: right;">Low Estimate</th>
                    <th style="text-align: right;">High Estimate</th>
                </tr>
            </thead>
            <tbody>
                ${costTableBodyHtml}
            </tbody>
            <tfoot>
                ${costTableFooterHtml}
            </tfoot>
        </table>
    `;
    // --- End of New Cost Analysis Rendering ---

    let rebatesHtml = '';
    if (costAnalysis.rebates.length > 0) {
        let rebateListItems = '';
        if (costAnalysis.rebates.some(r => r.name.includes('Heat Pump'))) {
            rebateListItems += `<li><strong>Whole-Home Heat Pump Rebate (${formatCurrency(REBATES['Heat Pump'])}):</strong> This is the largest incentive, typically available when you install a heat pump system that serves as the sole source of heating for your entire home, replacing 100% of your fossil fuel system.</li>`;
        }
        if (costAnalysis.rebates.some(r => r.name.includes('HPWH'))) {
            rebateListItems += `<li><strong>Heat Pump Water Heater (HPWH) Rebate (${formatCurrency(REBATES['HPWH'])}):</strong> An incentive for replacing a gas or standard electric water heater with a high-efficiency heat pump model.</li>`;
        }
        if (costAnalysis.rebates.some(r => r.name.includes('Panel Upgrade'))) {
            rebateListItems += `<li><strong>Panel Upgrade Rebate (${formatCurrency(REBATES['Panel Upgrade'])}):</strong> This incentive helps offset the cost of upgrading your main electrical panel when it's required for a heat pump installation.</li>`;
        }

        rebatesHtml = `
            <p>The cost analysis includes estimated rebates available through programs like Mass Save. The following provides context on the major incentives included in this plan:</p>
            <ul class="explanation-list">
                ${rebateListItems}
            </ul>
            <p class="report-note"><strong>Disclaimer:</strong> Rebate amounts and eligibility requirements change frequently. This estimate is for planning purposes only. You must confirm your eligibility and apply for rebates directly through the <a href="https://www.masssave.com/rebates" target="_blank" rel="noopener noreferrer">Mass Save website</a> or your utility provider.</p>
        `;
    }


    fullReportContentEl.innerHTML = `
        <div class="full-report-section">
            <h2 class="full-report-title">Electrification Report</h2>
            <p class="full-report-subtitle">For: ${address || 'N/A'}<br>Generated on: ${today}</p>
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">1. User Inputs Summary</h3>
            ${inputsHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">2. Electrification Plan</h3>
            ${planHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">3. Electrical Load Analysis</h3>
            ${electricalHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">4. Impact on Electric Grid</h3>
            ${distributionHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">5. Cost Analysis</h3>
            ${costHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">6. Rebate & Incentive Details</h3>
            ${rebatesHtml || '<p>No rebates were applicable to this specific electrification plan.</p>'}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">7. 15-Year Financial Outlook</h3>
             <p>
                Assumes a ${state.gasPriceEscalation}% annual increase in gas prices and a ${state.electricityPriceEscalation}% increase in electricity prices.
            </p>
            <div class="chart-container" style="height: 400px;">
                <canvas id="full-report-lifetime-chart"></canvas>
            </div>
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">8. Environmental Impact Analysis</h3>
            ${environmentalHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">9. Summary of Revisions</h3>
            <p>This report has been updated based on a review of initial assumptions. The following corrections have been applied:</p>
            <ul class="explanation-list">
                <li><strong>Electricity Price:</strong> Corrected to a realistic average commercial rate of ${formatCurrency(state.electricityPricePerKwh, false)}/kWh for Massachusetts.</li>
                <li><strong>Fuel Escalation Rates:</strong> Annual price increase projections were updated to ${state.gasPriceEscalation}% for natural gas and ${state.electricityPriceEscalation}% for electricity to better reflect market conditions.</li>
                <li><strong>Panel Upgrade:</strong> The electrical analysis now recommends an upgrade to a 600A service to meet the projected load, with associated commercial-scale costs.</li>
                <li><strong>Cost Estimates:</strong> Installation and electrical costs have been revised to align with current Massachusetts commercial project pricing, including increased costs for heat pump installation and new circuits.</li>
                <li><strong>Ductwork Contingency:</strong> A note has been added to the Electrification Plan highlighting the need to field-verify ductwork and to budget for a potential contingency for major modifications or replacement.</li>
                <li><strong>Environmental Recalculations:</strong> All environmental impact figures have been recalculated based on the corrected energy usage and cost inputs.</li>
            </ul>
        </div>
    `;

    fullReportLifetimeChart = createLifetimeChart('full-report-lifetime-chart', lifetimeAnalysis, fullReportLifetimeChart);

    if (window.MathJax) {
        window.MathJax.typesetPromise([fullReportContentEl]).catch((err: any) => console.error('MathJax typesetting error:', err));
    }
}


function renderAppliances() {
  appliancesListEl.innerHTML = ''; // Clear existing rows

  if (state.appliances.length === 0) {
    appliancesListEl.innerHTML = '<p class="empty-state">No appliances added yet.</p>';
    return;
  }
  
  const planningAnalysis = calculateAnalysis(state.appliances.filter(a => a.included), state);

  state.appliances.forEach(appliance => {
    const definition = APPLIANCE_DEFINITIONS[appliance.key];
    if (!definition) return; // Should not happen

    const id = appliance.id;
    const templateClone = applianceRowTemplate.content.cloneNode(true) as DocumentFragment;
    const applianceEntryEl = templateClone.querySelector('.appliance-entry') as HTMLDivElement;
    applianceEntryEl.dataset.id = id;

    const includeCheckbox = applianceEntryEl.querySelector('.include-in-plan') as HTMLInputElement;
    const includeLabel = applianceEntryEl.querySelector('label[for^="include-appliance-"]') as HTMLLabelElement;
    includeCheckbox.id = `include-appliance-${id}`;
    includeLabel.htmlFor = includeCheckbox.id;
    includeCheckbox.checked = appliance.included;

    const nameEl = applianceEntryEl.querySelector('.appliance-name') as HTMLDivElement;
    const btuInput = applianceEntryEl.querySelector('.btu-rating') as HTMLInputElement;
    const efficiencyInput = applianceEntryEl.querySelector('.efficiency-rating') as HTMLInputElement;
    const recommendationEl = applianceEntryEl.querySelector('.recommendation-details') as HTMLDivElement;

    nameEl.textContent = definition.name;
    btuInput.value = appliance.btu > 0 ? appliance.btu.toString() : '';
    efficiencyInput.value = appliance.efficiency > 0 ? appliance.efficiency.toString() : '';

    const recommendation = planningAnalysis.recommendations.find(r => r.id === appliance.id);
    if (recommendation) {
        recommendationEl.innerHTML = `${recommendation.text} <span>(Est. ${recommendation.amps.toFixed(1)} A / ${recommendation.breakerSpaces} spaces)</span>`;
    } else {
        recommendationEl.innerHTML = '';
    }

    appliancesListEl.appendChild(applianceEntryEl);
  });
}

function render() {
  const includedAppliances = state.appliances.filter(a => a.included);
  const planningAnalysis = calculateAnalysis(includedAppliances, state);
  const costAnalysis = calculateCostAnalysis(planningAnalysis, includedAppliances, state);
  const projectedAnnualKwh = calculateProjectedAnnualKwh(includedAppliances, state);
  const financialAnalysis = calculateFinancialAnalysis(costAnalysis, includedAppliances, state, projectedAnnualKwh);
  const environmentalImpact = calculateEnvironmentalImpact(includedAppliances, state, projectedAnnualKwh);
  const lifetimeAnalysis = calculateLifetimeAnalysis(financialAnalysis, costAnalysis, state);
  const distributionImpact = calculateDistributionImpact(planningAnalysis, state);
  
  renderAppliances();
  renderReport(planningAnalysis, costAnalysis, financialAnalysis, environmentalImpact, lifetimeAnalysis);
  renderFullReport(planningAnalysis, costAnalysis, financialAnalysis, environmentalImpact, lifetimeAnalysis, distributionImpact);
}

// --- EVENT HANDLERS ---

/** Reads all values from the UI inputs and updates the state object. */
function updateStateFromUI() {
  const newAppliances: Appliance[] = [];
  const applianceRows = appliancesListEl.querySelectorAll('.appliance-entry');

  applianceRows.forEach(row => {
    const id = (row as HTMLElement).dataset.id;
    const existingAppliance = state.appliances.find(a => a.id === id);
    if (!id || !existingAppliance) return;

    const btu = parseInt((row.querySelector('.btu-rating') as HTMLInputElement).value, 10) || 0;
    const efficiency = parseInt((row.querySelector('.efficiency-rating') as HTMLInputElement).value, 10) || 0;
    const included = (row.querySelector('.include-in-plan') as HTMLInputElement).checked;

    newAppliances.push({ id, key: existingAppliance.key, btu, included, efficiency });
  });

  state = {
    facilityType: facilityTypeEl.value as FacilityType,
    climateZone: climateZoneEl.value as ClimateZone,
    addressNumber: addressNumberEl.value,
    streetName: streetNameEl.value,
    town: townEl.value,
    annualHeatingTherms: parseInt(annualHeatingThermsEl.value, 10) || 0,
    annualNonHeatingTherms: parseInt(annualNonHeatingThermsEl.value, 10) || 0,
    panelAmps: parseInt(panelAmperageEl.value, 10),
    voltage: parseInt(voltageEl.value, 10) || 240,
    breakerSpaces: parseInt(breakerSpacesEl.value, 10) || 0,
    efficiencyTier: efficiencyTierEl.value as EfficiencyTier,
    appliances: newAppliances,
    gasPricePerTherm: parseFloat(gasPriceEl.value) || 0,
    electricityPricePerKwh: parseFloat(electricityPriceEl.value) || 0,
    gasPriceEscalation: parseFloat(gasPriceEscalationEl.value) || 0,
    electricityPriceEscalation: parseFloat(electricityPriceEscalationEl.value) || 0,
    highRiskGasEscalation: parseFloat(highRiskGasEscalationEl.value) || 0,
  };
}

/** Updates all UI input values based on the current state object. */
function updateUIFromState() {
    facilityTypeEl.value = state.facilityType;
    climateZoneEl.value = state.climateZone;
    addressNumberEl.value = state.addressNumber;
    streetNameEl.value = state.streetName;
    townEl.value = state.town;
    annualHeatingThermsEl.value = state.annualHeatingTherms.toString();
    annualNonHeatingThermsEl.value = state.annualNonHeatingTherms.toString();
    panelAmperageEl.value = state.panelAmps.toString();
    voltageEl.value = state.voltage.toString();
    breakerSpacesEl.value = state.breakerSpaces.toString();
    efficiencyTierEl.value = state.efficiencyTier;
    gasPriceEl.value = state.gasPricePerTherm.toFixed(2);
    electricityPriceEl.value = state.electricityPricePerKwh.toFixed(2);
    gasPriceEscalationEl.value = state.gasPriceEscalation.toString();
    electricityPriceEscalationEl.value = state.electricityPriceEscalation.toString();
    highRiskGasEscalationEl.value = state.highRiskGasEscalation.toString();
}


function updateStateAndRender() {
  updateStateFromUI();
  render();
}

function handleFacilityTypeChange() {
    const facilityType = facilityTypeEl.value as FacilityType;
    const defaults = DEFAULT_USAGE_BY_FACILITY[facilityType];
    if (defaults) {
        annualHeatingThermsEl.value = defaults.heating.toString();
        annualNonHeatingThermsEl.value = defaults.nonHeating.toString();
    }
    updateStateAndRender();
}

function addApplianceByKey(key: string) {
    const definition = APPLIANCE_DEFINITIONS[key];
    if (!definition) return;

    state.appliances.push({
        id: crypto.randomUUID(),
        key: key,
        btu: definition.defaultBtu,
        included: true,
        efficiency: definition.defaultEfficiency,
    });
    // We need to re-render everything, including the appliance list itself
    render();
}


function handleRemoveAppliance(event: Event) {
    const target = event.target as HTMLElement;
    const button = target.closest('.remove-btn');
    if (button) {
        const row = button.closest('.appliance-entry') as HTMLElement;
        const idToRemove = row.dataset.id;
        if (idToRemove) {
            state.appliances = state.appliances.filter(app => app.id !== idToRemove);
            render();
        }
    }
}

// --- MODAL LOGIC ---
function openAddApplianceModal() {
    modalTitleEl.textContent = `Add Appliance for ${state.facilityType}`;
    modalApplianceGridEl.innerHTML = ''; // Clear previous buttons

    const applianceKeys = FACILITY_APPLIANCE_MAP[state.facilityType];
    applianceKeys.forEach(key => {
        const definition = APPLIANCE_DEFINITIONS[key];
        const button = document.createElement('button');
        button.className = 'modal-appliance-btn';
        button.textContent = definition.name;
        button.dataset.key = key;
        modalApplianceGridEl.appendChild(button);
    });

    modalOverlayEl.classList.remove('hidden');
    applianceModalEl.classList.remove('hidden');
}

function closeAddApplianceModal() {
    modalOverlayEl.classList.add('hidden');
    applianceModalEl.classList.add('hidden');
}

function handleModalGridClick(event: MouseEvent) {
    const target = event.target as HTMLElement;
    const button = target.closest<HTMLElement>('.modal-appliance-btn');
    if (button && button.dataset.key) {
        addApplianceByKey(button.dataset.key);
        closeAddApplianceModal();
    }
}

// --- TAB LOGIC ---
function handleTabClick(event: MouseEvent) {
    const clickedButton = event.currentTarget as HTMLButtonElement;
    
    if (clickedButton.classList.contains('active')) return;
    
    const allTabs = [tabPlannerBtn, tabReportBtn, tabSaveLoadBtn];
    const allContent = [plannerContentEl, reportContentWrapperEl, saveLoadContentEl];

    allTabs.forEach(tab => tab.classList.remove('active'));
    allContent.forEach(content => content.classList.remove('active'));
    
    clickedButton.classList.add('active');

    if (clickedButton === tabPlannerBtn) plannerContentEl.classList.add('active');
    else if (clickedButton === tabReportBtn) reportContentWrapperEl.classList.add('active');
    else if (clickedButton === tabSaveLoadBtn) saveLoadContentEl.classList.add('active');
    
    // Re-render charts if the report tab is now active, in case it wasn't visible before
    if (clickedButton === tabReportBtn) render();
}

// --- SAVE/LOAD/EXPORT LOGIC ---

function handleExportPdf() {
    window.print();
}

function handleSaveProject() {
    updateStateFromUI(); // Ensure state is current before saving
    const stateJson = JSON.stringify(state, null, 2);
    const blob = new Blob([stateJson], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'electrification-project.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function handleLoadProject(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) {
        loadedFilenameEl.textContent = 'No file selected.';
        return;
    }

    const file = input.files[0];
    loadedFilenameEl.textContent = file.name;
    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const text = e.target?.result;
            if (typeof text !== 'string') {
                throw new Error('File could not be read as text.');
            }
            const loadedState = JSON.parse(text);

            // Basic validation
            if (!loadedState.facilityType || !Array.isArray(loadedState.appliances)) {
                throw new Error('Invalid project file format.');
            }

            state = loadedState as State;
            updateUIFromState();
            render();
            alert('Project loaded successfully!');

        } catch (error) {
            console.error('Failed to load or parse project file:', error);
            alert(`Error: Could not load project file. Please ensure it is a valid, unmodified project file.\n\n(${(error as Error).message})`);
            loadedFilenameEl.textContent = 'Load failed. Please try again.';
        }
    };

    reader.onerror = () => {
        alert('An error occurred while reading the file.');
        loadedFilenameEl.textContent = 'Load failed. Please try again.';
    };

    reader.readAsText(file);
}

// --- INITIALIZATION ---
const allInputs = document.querySelectorAll('#inputs input, #inputs select');
allInputs.forEach(input => {
    input.addEventListener('input', updateStateAndRender);
    input.addEventListener('change', updateStateAndRender);
});
facilityTypeEl.addEventListener('change', handleFacilityTypeChange);


// Modal listeners
addApplianceBtn.addEventListener('click', openAddApplianceModal);
modalCloseBtn.addEventListener('click', closeAddApplianceModal);
modalOverlayEl.addEventListener('click', closeAddApplianceModal);
modalApplianceGridEl.addEventListener('click', handleModalGridClick);

// Tab Listeners
tabPlannerBtn.addEventListener('click', handleTabClick);
tabReportBtn.addEventListener('click', handleTabClick);
tabSaveLoadBtn.addEventListener('click', handleTabClick);

// Save/Load/Export Listeners
exportPdfBtn.addEventListener('click', handleExportPdf);
saveProjectBtn.addEventListener('click', handleSaveProject);
loadProjectInput.addEventListener('change', handleLoadProject);

appliancesListEl.addEventListener('click', (e) => {
    const target = e.target as HTMLElement;
    if (target.closest('.remove-btn')) {
        handleRemoveAppliance(e);
    } else if (target.classList.contains('include-in-plan') || (e.target as HTMLInputElement).type === 'number') {
        // No need to call updateStateAndRender for number inputs, as 'input' event already handles it.
        // This just handles the checkbox click specifically.
        updateStateAndRender();
    }
});

appliancesListEl.addEventListener('input', (e) => {
    const target = e.target as HTMLElement;
    if (target.classList.contains('btu-rating') || target.classList.contains('efficiency-rating')) {
        updateStateAndRender();
    }
});


// Initial population of default values and render
updateUIFromState();
handleFacilityTypeChange();