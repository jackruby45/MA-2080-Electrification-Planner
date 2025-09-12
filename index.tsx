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
type GhgUnit = 'kg' | 't' | 'lb' | 'short-ton' | 'homes-electricity' | 'oil-barrels' | 'trash-bags' | 'gasoline-gallons' | 'smartphones-charged' | 'forest-acres-sequestered' | 'coal-pounds-burned' | 'propane-cylinders' | 'natural-gas-homes';


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
  ghgUnit: GhgUnit;
  // --- CONFIGURABLE DATA ---
  equipmentCosts: Record<FacilityType, Partial<Record<ApplianceCategory, Record<EfficiencyTier, number[]>>>>;
  installationCosts: Record<FacilityType, Partial<Record<ApplianceCategory, number[]>>>;
  decommissioningCosts: Record<FacilityType, Partial<Record<ApplianceCategory, number[]>>>;
  distributionSystemCosts: Record<FacilityType, Record<string, number[]>>;
  electricalUpgradeCosts: Record<FacilityType, Record<string, number[]>>;
  rebates: Record<string, number>;
  annualMaintenanceCosts: Record<FacilityType, Record<string, number[]>>;
  copLookup: Record<ApplianceCategory, Record<EfficiencyTier, number>>;
  emissionsFactors: {
    co2KgPerThermGas: number;
    upstreamLeakageRate: number;
    onsiteSlippageRate: number;
    kgCh4InTherm: number;
    gwp20ch4: number;
    gwp100ch4: number;
    co2KgPerKwhMa: number;
  };
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
    diversifiedTotalAmps: number;
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

const FACILITY_TYPES: FacilityType[] = ['Residential', 'Small Commercial', 'Large Commercial', 'Industrial', 'Restaurant', 'Medical', 'Nursing Home'];

const FACILITY_APPLIANCE_MAP: Record<FacilityType, string[]> = {
    'Residential': ['res-furnace', 'res-boiler', 'res-tank-wh', 'res-tankless-wh', 'res-range', 'res-dryer', 'res-pool-heater', 'res-fireplace'],
    'Small Commercial': ['comm-rtu', 'comm-boiler', 'comm-wh'],
    'Large Commercial': ['comm-rtu', 'comm-boiler', 'comm-wh', 'laundry-dryer'],
    'Industrial': ['comm-rtu', 'comm-boiler', 'comm-wh', 'laundry-dryer'],
    'Restaurant': ['rest-range', 'rest-convection-oven', 'rest-fryer', 'rest-griddle', 'comm-rtu', 'comm-wh', 'rest-booster-heater'],
    'Medical': ['comm-rtu', 'comm-boiler', 'comm-wh', 'laundry-dryer', 'rest-range', 'rest-convection-oven'],
    'Nursing Home': ['comm-rtu', 'comm-boiler', 'comm-wh', 'laundry-dryer', 'rest-range', 'rest-convection-oven'],
};

const getApplianceCategoriesForFacility = (facilityType: FacilityType): ApplianceCategory[] => {
    const categories = new Set<ApplianceCategory>();
    const applianceKeys = FACILITY_APPLIANCE_MAP[facilityType] || [];
    for (const key of applianceKeys) {
        const definition = APPLIANCE_DEFINITIONS[key];
        if (definition) {
            categories.add(definition.category);
        }
    }
    return Array.from(categories);
};

const BTU_PER_THERM = 100000;
const BTU_TO_KW = 0.000293071;
const PANEL_SIZES = [100, 125, 150, 200, 400, 600];
const LIFETIME_YEARS = 15;
const DIVERSITY_FACTOR = 0.6; // 60% diversity factor for remaining loads

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

const KG_CO2E_PER_CAR_YEAR = 4600; // Source: EPA, avg passenger vehicle emits 4.6 metric tons CO2e/year
const GHG_CONVERSIONS: Record<GhgUnit, { factor: number, label: string }> = {
    'kg': { factor: 1, label: 'kg CO₂e' },
    't': { factor: 0.001, label: 't CO₂e' },
    'lb': { factor: 2.20462, label: 'lb CO₂e' },
    'short-ton': { factor: 0.00110231, label: 'US tons CO₂e' },
    'homes-electricity': { factor: 1 / 8930, label: 'homes\' electricity use / year' }, // Source: EPA, 8,930 kg/home/yr
    'oil-barrels': { factor: 1 / 430, label: 'barrels of oil consumed' }, // Source: EPA, 0.43 t/barrel
    'trash-bags': { factor: 1 / 8.03, label: 'trash bags recycled' }, // Source: EPA, 8.03 kg/bag
    'gasoline-gallons': { factor: 1 / 8.89, label: 'gallons of gasoline' }, // Source: EPA, 8.89 kg/gallon
    'smartphones-charged': { factor: 1 / 0.008, label: 'smartphones charged' }, // Source: EPA, 8g/charge
    'forest-acres-sequestered': { factor: 1 / 840, label: 'acres of U.S. forest' }, // Source: EPA, 0.84 t/acre/yr
    'coal-pounds-burned': { factor: 2.20462 / 2.05, label: 'lbs of coal burned' }, // Source: EPA, 2.05 lbs CO2/lb coal
    'propane-cylinders': { factor: 1 / 23.2, label: 'propane cylinders' }, // Source: EPA, 23.2 kg/cylinder
    'natural-gas-homes': { factor: 1 / 5150, label: 'homes\' natural gas use / year' }, // Source: EPA, 5.15 t/home/yr
};

const APPLIANCE_LIFESPANS: Record<string, { years: number, color: string }> = {
    'Gas Furnace/Boiler': { years: 18, color: '#dc3545' },
    'Heat Pump (replaces Furnace/Boiler)': { years: 18, color: '#005fcc' },
    'Gas Tank Water Heater': { years: 10, color: '#dc3545' },
    'Heat Pump Water Heater': { years: 13, color: '#005fcc' },
};

const GRID_DECARBONIZATION_RATE_PER_YEAR = 0.025; // 2.5% annual reduction in grid emissions factor

// --- DOM ELEMENT SELECTORS ---
const facilityTypeEl = document.getElementById('facility-type') as HTMLSelectElement;
const climateZoneEl = document.getElementById('climate-zone') as HTMLSelectElement;
const addressNumberEl = document.getElementById('address-number') as HTMLInputElement;
const streetNameEl = document.getElementById('street-name') as HTMLInputElement;
const townEl = document.getElementById('town') as HTMLInputElement;

const ghgUnitEl = document.getElementById('ghg-unit') as HTMLSelectElement;
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
const tabMethodologyBtn = document.getElementById('tab-methodology') as HTMLButtonElement;
const tabSaveLoadBtn = document.getElementById('tab-save-load') as HTMLButtonElement;
const plannerContentEl = document.getElementById('planner-content') as HTMLDivElement;
const reportContentWrapperEl = document.getElementById('report-content-wrapper') as HTMLDivElement;
const fullReportContentEl = document.getElementById('full-report-content') as HTMLDivElement;
const methodologyContentEl = document.getElementById('methodology-content') as HTMLDivElement;
const saveLoadContentEl = document.getElementById('save-load-content') as HTMLDivElement;


// New Save/Load/Export elements
const exportPdfBtn = document.getElementById('export-pdf-btn') as HTMLButtonElement;


// --- CHART INSTANCES ---
let fullReportLifetimeChart: any | null = null;
let costBreakdownChart: any | null = null;
let emissionsBreakdownChart: any | null = null;
let annualCostsChart: any | null = null;
let paybackChart: any | null = null;
let loadContributionChart: any | null = null;
let operatingCostChart: any | null = null;
let gridDecarbonizationChart: any | null = null;


// --- APPLICATION STATE ---
let state: State = {
  facilityType: 'Residential',
  climateZone: 'Zone5',
  addressNumber: '123',
  streetName: 'Elm Street',
  town: 'Anytown',
  annualHeatingTherms: 700,
  annualNonHeatingTherms: 250,
  panelAmps: 100,
  voltage: 240,
  breakerSpaces: 8,
  efficiencyTier: 'High',
  appliances: [
      { id: crypto.randomUUID(), key: 'res-boiler', btu: 100000, included: true, efficiency: 80 },
      { id: crypto.randomUUID(), key: 'res-tank-wh', btu: 40000, included: true, efficiency: 62 },
      { id: crypto.randomUUID(), key: 'res-dryer', btu: 22000, included: true, efficiency: 75 },
      { id: crypto.randomUUID(), key: 'res-range', btu: 60000, included: true, efficiency: 40 },
  ],
  gasPricePerTherm: 1.50,
  electricityPricePerKwh: 0.22,
  gasPriceEscalation: 3,
  electricityPriceEscalation: 2,
  highRiskGasEscalation: 7,
  ghgUnit: 'kg',

  equipmentCosts: {
    'Residential': {
        'Furnace': { 'Standard': [6000, 8000], 'High': [8000, 12000], 'Premium': [12000, 18000] },
        'Boiler': { 'Standard': [7000, 9000], 'High': [9000, 13000], 'Premium': [13000, 19000] },
        'Water Heater': { 'Standard': [1500, 2500], 'High': [2500, 4000], 'Premium': [4000, 6000] },
        'Range': { 'Standard': [1000, 2000], 'High': [2000, 3500], 'Premium': [3500, 5000] },
        'Dryer': { 'Standard': [800, 1200], 'High': [1200, 1800], 'Premium': [1800, 2500] },
        'Space Heater': { 'Standard': [1500, 2500], 'High': [2500, 4000], 'Premium': [4000, 5500] },
    },
    'Small Commercial': {
        'Furnace': { 'Standard': [10000, 15000], 'High': [15000, 22000], 'Premium': [22000, 30000] },
        'Boiler': { 'Standard': [12000, 18000], 'High': [18000, 26000], 'Premium': [26000, 35000] },
        'Water Heater': { 'Standard': [4000, 6000], 'High': [6000, 9000], 'Premium': [9000, 12000] },
    },
    'Large Commercial': {
        'Furnace': { 'Standard': [25000, 40000], 'High': [40000, 60000], 'Premium': [60000, 80000] },
        'Boiler': { 'Standard': [40000, 60000], 'High': [60000, 90000], 'Premium': [90000, 120000] },
        'Water Heater': { 'Standard': [8000, 12000], 'High': [12000, 18000], 'Premium': [18000, 25000] },
        'Dryer': { 'Standard': [3000, 5000], 'High': [5000, 7000], 'Premium': [7000, 9000] },
    },
    'Industrial': {
        'Furnace': { 'Standard': [30000, 50000], 'High': [50000, 75000], 'Premium': [75000, 100000] },
        'Boiler': { 'Standard': [50000, 80000], 'High': [80000, 120000], 'Premium': [120000, 160000] },
        'Water Heater': { 'Standard': [10000, 15000], 'High': [15000, 22000], 'Premium': [22000, 30000] },
        'Dryer': { 'Standard': [4000, 6000], 'High': [6000, 8000], 'Premium': [8000, 10000] },
    },
    'Restaurant': {
        'Furnace': { 'Standard': [10000, 15000], 'High': [15000, 22000], 'Premium': [22000, 30000] },
        'Water Heater': { 'Standard': [5000, 8000], 'High': [8000, 12000], 'Premium': [12000, 16000] },
        'Range': { 'Standard': [4000, 7000], 'High': [7000, 11000], 'Premium': [11000, 15000] },
    },
    'Medical': {
        'Furnace': { 'Standard': [25000, 40000], 'High': [40000, 60000], 'Premium': [60000, 80000] },
        'Boiler': { 'Standard': [40000, 60000], 'High': [60000, 90000], 'Premium': [90000, 120000] },
        'Water Heater': { 'Standard': [8000, 12000], 'High': [12000, 18000], 'Premium': [18000, 25000] },
        'Dryer': { 'Standard': [3000, 5000], 'High': [5000, 7000], 'Premium': [7000, 9000] },
        'Range': { 'Standard': [2000, 4000], 'High': [4000, 6000], 'Premium': [6000, 8000] },
    },
    'Nursing Home': {
        'Furnace': { 'Standard': [25000, 40000], 'High': [40000, 60000], 'Premium': [60000, 80000] },
        'Boiler': { 'Standard': [40000, 60000], 'High': [60000, 90000], 'Premium': [90000, 120000] },
        'Water Heater': { 'Standard': [8000, 12000], 'High': [12000, 18000], 'Premium': [18000, 25000] },
        'Dryer': { 'Standard': [3000, 5000], 'High': [5000, 7000], 'Premium': [7000, 9000] },
        'Range': { 'Standard': [2000, 4000], 'High': [4000, 6000], 'Premium': [6000, 8000] },
    },
  },
  installationCosts: {
    'Residential': { 'Furnace': [9000, 18000], 'Boiler': [9000, 18000], 'Water Heater': [3000, 6000], 'Range': [300, 600], 'Dryer': [300, 600], 'Space Heater': [7000, 14000] },
    'Small Commercial': { 'Furnace': [12000, 22000], 'Boiler': [15000, 25000], 'Water Heater': [5000, 8000] },
    'Large Commercial': { 'Furnace': [30000, 50000], 'Boiler': [50000, 80000], 'Water Heater': [10000, 15000], 'Dryer': [2000, 4000] },
    'Industrial': { 'Furnace': [40000, 60000], 'Boiler': [60000, 100000], 'Water Heater': [12000, 18000], 'Dryer': [2500, 5000] },
    'Restaurant': { 'Furnace': [12000, 22000], 'Water Heater': [6000, 9000], 'Range': [1000, 2000] },
    'Medical': { 'Furnace': [30000, 50000], 'Boiler': [50000, 80000], 'Water Heater': [10000, 15000], 'Dryer': [2000, 4000], 'Range': [500, 1000] },
    'Nursing Home': { 'Furnace': [30000, 50000], 'Boiler': [50000, 80000], 'Water Heater': [10000, 15000], 'Dryer': [2000, 4000], 'Range': [500, 1000] },
  },
  decommissioningCosts: {
    'Residential': { 'Furnace': [250, 500], 'Boiler': [250, 500], 'Water Heater': [150, 300], 'Range': [100, 200], 'Dryer': [0, 0], 'Space Heater': [50, 100] },
    'Small Commercial': { 'Furnace': [500, 1000], 'Boiler': [600, 1200], 'Water Heater': [300, 600] },
    'Large Commercial': { 'Furnace': [1000, 2000], 'Boiler': [1500, 3000], 'Water Heater': [500, 1000], 'Dryer': [150, 300] },
    'Industrial': { 'Furnace': [1200, 2500], 'Boiler': [2000, 4000], 'Water Heater': [600, 1200], 'Dryer': [200, 400] },
    'Restaurant': { 'Furnace': [500, 1000], 'Water Heater': [400, 800], 'Range': [200, 400] },
    'Medical': { 'Furnace': [1000, 2000], 'Boiler': [1500, 3000], 'Water Heater': [500, 1000], 'Dryer': [150, 300], 'Range': [150, 300] },
    'Nursing Home': { 'Furnace': [1000, 2000], 'Boiler': [1500, 3000], 'Water Heater': [500, 1000], 'Dryer': [150, 300], 'Range': [150, 300] },
  },
  distributionSystemCosts: {
    'Residential': { 'Ductwork Modification': [500, 2000], 'New Heat Distribution System': [4000, 10000] },
    'Small Commercial': { 'Ductwork Modification': [2000, 8000], 'New Heat Distribution System': [8000, 20000] },
    'Large Commercial': { 'Ductwork Modification': [10000, 30000], 'New Heat Distribution System': [25000, 60000] },
    'Industrial': { 'Ductwork Modification': [15000, 50000], 'New Heat Distribution System': [30000, 80000] },
    'Restaurant': { 'Ductwork Modification': [2000, 8000], 'New Heat Distribution System': [8000, 20000] },
    'Medical': { 'Ductwork Modification': [10000, 30000], 'New Heat Distribution System': [25000, 60000] },
    'Nursing Home': { 'Ductwork Modification': [10000, 30000], 'New Heat Distribution System': [25000, 60000] },
  },
  electricalUpgradeCosts: {
    'Residential': { 'New Circuits (10)': [1500, 3000], 'Sub-panel': [1500, 2500], 'Panel Upgrade to 125A': [2000, 3500], 'Panel Upgrade to 150A': [2500, 4000], 'Panel Upgrade to 200A': [3000, 5000], 'Panel Upgrade to 400A': [8000, 12000], 'Panel Upgrade to 600A': [15000, 25000] },
    'Small Commercial': { 'New Circuits (10)': [2500, 5000], 'Sub-panel': [2000, 3500], 'Panel Upgrade to 125A': [2500, 4000], 'Panel Upgrade to 150A': [3000, 5000], 'Panel Upgrade to 200A': [4000, 6000], 'Panel Upgrade to 400A': [10000, 15000], 'Panel Upgrade to 600A': [18000, 30000] },
    'Large Commercial': { 'New Circuits (10)': [5000, 10000], 'Sub-panel': [3000, 5000], 'Panel Upgrade to 125A': [3000, 5000], 'Panel Upgrade to 150A': [4000, 6000], 'Panel Upgrade to 200A': [5000, 8000], 'Panel Upgrade to 400A': [12000, 18000], 'Panel Upgrade to 600A': [20000, 35000] },
    'Industrial': { 'New Circuits (10)': [6000, 12000], 'Sub-panel': [4000, 7000], 'Panel Upgrade to 125A': [3500, 5500], 'Panel Upgrade to 150A': [4500, 7000], 'Panel Upgrade to 200A': [6000, 10000], 'Panel Upgrade to 400A': [15000, 25000], 'Panel Upgrade to 600A': [25000, 45000] },
    'Restaurant': { 'New Circuits (10)': [3000, 6000], 'Sub-panel': [2500, 4000], 'Panel Upgrade to 125A': [2800, 4500], 'Panel Upgrade to 150A': [3500, 5500], 'Panel Upgrade to 200A': [4500, 7000], 'Panel Upgrade to 400A': [11000, 16000], 'Panel Upgrade to 600A': [19000, 32000] },
    'Medical': { 'New Circuits (10)': [7000, 15000], 'Sub-panel': [4000, 7000], 'Panel Upgrade to 125A': [4000, 6000], 'Panel Upgrade to 150A': [5000, 8000], 'Panel Upgrade to 200A': [7000, 11000], 'Panel Upgrade to 400A': [16000, 26000], 'Panel Upgrade to 600A': [28000, 50000] },
    'Nursing Home': { 'New Circuits (10)': [7000, 15000], 'Sub-panel': [4000, 7000], 'Panel Upgrade to 125A': [4000, 6000], 'Panel Upgrade to 150A': [5000, 8000], 'Panel Upgrade to 200A': [7000, 11000], 'Panel Upgrade to 400A': [16000, 26000], 'Panel Upgrade to 600A': [28000, 50000] },
  },
  rebates: {
    'Heat Pump': 10000, 'HPWH': 750, 'Panel Upgrade': 1500,
  },
  annualMaintenanceCosts: {
    'Residential': { 'Gas System': [200, 400], 'Electric System': [100, 200] },
    'Small Commercial': { 'Gas System': [500, 1000], 'Electric System': [250, 500] },
    'Large Commercial': { 'Gas System': [2000, 4000], 'Electric System': [1000, 2000] },
    'Industrial': { 'Gas System': [5000, 10000], 'Electric System': [2500, 5000] },
    'Restaurant': { 'Gas System': [800, 1500], 'Electric System': [400, 750] },
    'Medical': { 'Gas System': [3000, 6000], 'Electric System': [1500, 3000] },
    'Nursing Home': { 'Gas System': [3000, 6000], 'Electric System': [1500, 3000] },
  },
  copLookup: {
    'Furnace':      { 'Standard': 2.5, 'High': 3.0, 'Premium': 3.5 },
    'Boiler':       { 'Standard': 2.5, 'High': 3.0, 'Premium': 3.5 },
    'Water Heater': { 'Standard': 2.8, 'High': 3.5, 'Premium': 4.0 },
    'Space Heater': { 'Standard': 2.5, 'High': 3.0, 'Premium': 3.5 },
    'Range':        { 'Standard': 1, 'High': 1, 'Premium': 1 },
    'Dryer':        { 'Standard': 1.5, 'High': 2.0, 'Premium': 2.5 },
  },
  emissionsFactors: {
    co2KgPerThermGas: 5.3,
    upstreamLeakageRate: 0.015,
    onsiteSlippageRate: 0.01,
    kgCh4InTherm: 2.36,
    gwp20ch4: 84,
    gwp100ch4: 28,
    co2KgPerKwhMa: 0.26,
  }
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
      return (appliance.btu * BTU_TO_KW) / state.copLookup[category][efficiency];
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

    // DIVERSITY CALCULATION: 100% of largest load + DIVERSITY_FACTOR of the rest.
    let diversifiedTotalAmps = 0;
    if (recommendations.length > 0) {
        const ampsList = recommendations.map(r => r.amps);
        const maxAmps = Math.max(...ampsList);
        const sumOfRemainingAmps = totalNewAmps - maxAmps;
        diversifiedTotalAmps = maxAmps + (sumOfRemainingAmps * DIVERSITY_FACTOR);
    }

    const assumedExistingLoad = currentState.panelAmps * 0.8 * 0.5;
    // Use the more realistic DIVERSIFIED load for panel sizing calculations
    const totalCalculatedLoad = assumedExistingLoad + diversifiedTotalAmps;
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

    return { recommendations, totalNewAmps, diversifiedTotalAmps, requiredBreakerSpaces, breakerStatus, panelStatus };
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
    
    if (recommendedApplianceTypes.has('Heat Pump') && state.rebates['Heat Pump']) {
        rebates.push({ name: 'Heat Pump Rebate', amount: state.rebates['Heat Pump'] });
    }
    if (recommendedApplianceTypes.has('HPWH') && state.rebates['HPWH']) {
        rebates.push({ name: 'HPWH Rebate', amount: state.rebates['HPWH'] });
    }
    if (analysis.panelStatus.startsWith('Upgrade') && recommendedApplianceTypes.has('Heat Pump') && state.rebates['Panel Upgrade']) {
         rebates.push({ name: 'Panel Upgrade Rebate', amount: state.rebates['Panel Upgrade'] });
    }

    return rebates;
}

function calculateCostAnalysis(analysis: PlanningAnalysis, includedAppliances: Appliance[], state: State): CostAnalysis {
    const applianceCosts = new Map<string, CostItem[]>();
    let totalLow = 0, totalHigh = 0;

    const facilityEquipCosts = state.equipmentCosts[state.facilityType];
    const facilityInstallCosts = state.installationCosts[state.facilityType];
    const facilityDecomCosts = state.decommissioningCosts[state.facilityType];
    const facilityDistCosts = state.distributionSystemCosts[state.facilityType];
    const facilityElecCosts = state.electricalUpgradeCosts[state.facilityType];

    analysis.recommendations.forEach(rec => {
        const appliance = state.appliances.find(a => a.id === rec.id)!;
        const definition = APPLIANCE_DEFINITIONS[appliance.key];
        const category = definition.category;
        const costs: CostItem[] = [];

        const [equipLow, equipHigh] = facilityEquipCosts[category]?.[state.efficiencyTier] ?? [0, 0];
        costs.push({ name: 'Equipment', low: equipLow, high: equipHigh });

        let [installLow, installHigh] = facilityInstallCosts[category] ?? [0, 0];
        let installName = 'Installation & Materials';

        // For boiler replacements, the 'installation' includes creating the new mini-split distribution system.
        if (category === 'Boiler') {
            const [distLow, distHigh] = facilityDistCosts['New Heat Distribution System'];
            installLow += distLow;
            installHigh += distHigh;
            installName = 'Installation (incl. Mini-Split Distribution)';
        }
        costs.push({ name: installName, low: installLow, high: installHigh });
        
        const [decomLow, decomHigh] = facilityDecomCosts[category] ?? [0, 0];
        if (decomHigh > 0) {
            costs.push({ name: 'Gas Decommissioning', low: decomLow, high: decomHigh });
        }

        if (category === 'Furnace') {
            const [low, high] = facilityDistCosts['Ductwork Modification'];
            costs.push({ name: 'Ductwork Modification', low, high });
        }

        applianceCosts.set(rec.id, costs);
    });

    const electricalCosts: CostItem[] = [];
    if (analysis.panelStatus.startsWith('Upgrade')) {
        const size = analysis.panelStatus.match(/\d+/)?.[0];
        const key = `Panel Upgrade to ${size}A`;
        if (facilityElecCosts[key]) {
            const [low, high] = facilityElecCosts[key];
            electricalCosts.push({ name: 'Main Panel Upgrade', low, high });
        }
    } else if (analysis.breakerStatus === 'Sub-panel Recommended') {
        const [low, high] = facilityElecCosts['Sub-panel'];
        electricalCosts.push({ name: 'Sub-panel Installation', low, high });
    }

    if (analysis.recommendations.length > 0) {
        const [low, high] = facilityElecCosts['New Circuits (10)'];
        electricalCosts.push({ name: `${analysis.recommendations.length} New Circuit(s)`, low, high });
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
            
            let cop = state.copLookup[definition.category][state.efficiencyTier];
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
    const { emissionsFactors: ef } = state;
    const totalLeakageRate = ef.upstreamLeakageRate + ef.onsiteSlippageRate;

    // Current gas emissions
    const currentAnnualGasCo2Kg = totalTherms * ef.co2KgPerThermGas;
    const currentAnnualCh4Kg = totalTherms * ef.kgCh4InTherm * totalLeakageRate;
    const currentAnnualGasCo2e20Kg = currentAnnualGasCo2Kg + (currentAnnualCh4Kg * ef.gwp20ch4);
    const currentAnnualGasCo2e100Kg = currentAnnualGasCo2Kg + (currentAnnualCh4Kg * ef.gwp100ch4);

    // Projected electric emissions
    const projectedAnnualElecCo2Kg = projectedAnnualKwh * ef.co2KgPerKwhMa;

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

function formatUnitPrice(value: number): string {
    // Show 2 decimal places for unit prices like $/therm or $/kWh
    return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
}

function getFormattedGhgParts(valueKg: number, unit: GhgUnit): {value: string, label: string} {
    const conversion = GHG_CONVERSIONS[unit];
    const convertedValue = valueKg * conversion.factor;
    // Adjust precision based on value
    const fractionDigits = convertedValue < 10 ? 2 : (convertedValue < 100 ? 1 : 0);
    return {
        value: convertedValue.toLocaleString(undefined, {maximumFractionDigits: fractionDigits}),
        label: conversion.label
    };
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

// Donut Chart for Cost Breakdown
function createCostBreakdownDonutChart(canvasId: string, breakdown: Record<string, number>, netCost: number, chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;

    if (chartInstance) {
        chartInstance.destroy();
    }

    const data = {
        labels: Object.keys(breakdown),
        datasets: [{
            data: Object.values(breakdown),
            backgroundColor: ['#003f5c', '#58508d', '#bc5090', '#ff6361', '#ffa600'],
            borderColor: 'white',
            borderWidth: 2,
        }]
    };

    return new Chart(ctx, {
        type: 'doughnut',
        data: data,
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' },
                title: { display: true, text: `Net Project Cost: ${formatCurrency(netCost)}` },
                tooltip: {
                    callbacks: {
                        label: function(context: any) {
                            let label = context.label || '';
                            if (label) { label += ': '; }
                            label += formatCurrency(context.raw);
                            return label;
                        }
                    }
                }
            }
        },
    });
}

// Stacked Bar Chart for Emissions
function createEmissionsBarChart(canvasId: string, gasData: number[], elecData: number, chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;

    if (chartInstance) { chartInstance.destroy(); }

    return new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Current Gas System', 'Projected Electric System'],
            datasets: [
                { label: 'CO₂ (Combustion)', data: [gasData[0], 0], backgroundColor: '#58508d' },
                { label: 'CH₄ (Upstream Leakage CO₂e)', data: [gasData[1], 0], backgroundColor: '#bc5090' },
                { label: 'CH₄ (On-site Slippage CO₂e)', data: [gasData[2], 0], backgroundColor: '#ff6361' },
                { label: 'Grid Electricity CO₂e', data: [0, elecData], backgroundColor: '#003f5c' },
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                tooltip: { callbacks: { label: (c: any) => `${c.dataset.label}: ${c.raw.toFixed(0)} kg CO₂e` } }
            },
            scales: {
                x: { stacked: true },
                y: { stacked: true, beginAtZero: true, title: { display: true, text: 'kg CO₂e / year' } }
            }
        }
    });
}

// Simple Bar chart for annual costs
function createAnnualCostsBarChart(canvasId: string, gasCost: number, elecCost: number, chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;

    if (chartInstance) { chartInstance.destroy(); }

    return new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Current Gas Cost', 'Projected Electric Cost'],
            datasets: [{
                data: [gasCost, elecCost],
                backgroundColor: ['#dc3545', '#005fcc'],
                borderColor: ['#dc3545', '#005fcc'],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { callbacks: { label: (c: any) => formatCurrency(c.raw) } }
            },
            scales: { y: { beginAtZero: true, ticks: { callback: (v: any) => formatCurrency(v) } } }
        }
    });
}

// Line chart for payback
function createPaybackLineChart(canvasId: string, cashFlowData: number[], chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;
    
    if (chartInstance) { chartInstance.destroy(); }

    const labels = Array.from({ length: cashFlowData.length }, (_, i) => `Year ${i}`);

    return new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Cumulative Cash Flow',
                data: cashFlowData,
                borderColor: '#005fcc',
                backgroundColor: 'rgba(0, 95, 204, 0.1)',
                fill: true,
                tension: 0.1,
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                tooltip: { callbacks: { label: (c: any) => `Cumulative: ${formatCurrency(c.raw)}` } }
            },
            scales: { y: { ticks: { callback: (v: any) => formatCurrency(v) } } }
        }
    });
}

// Horizontal bar chart for load contribution
function createLoadContributionBarChart(canvasId: string, recommendations: ApplianceRecommendation[], chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;

    if (chartInstance) { chartInstance.destroy(); }

    const labels = recommendations.map(r => r.text.replace(/<strong>/g, '').replace(/<\/strong>/g, ''));
    const data = recommendations.map(r => r.amps);

    return new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Amps',
                data: data,
                backgroundColor: '#005fcc'
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { callbacks: { label: (c: any) => `${c.raw.toFixed(1)} Amps` } }
            },
            scales: { x: { beginAtZero: true, title: { display: true, text: 'Amps' } } }
        }
    });
}

// Line chart for operating costs
function createOperatingCostLineChart(canvasId: string, gasCosts: number[], elecCosts: number[], chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;

    if (chartInstance) { chartInstance.destroy(); }
    
    const labels = Array.from({ length: gasCosts.length }, (_, i) => `Year ${i}`);

    return new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Cumulative Electric Operating Cost',
                    data: elecCosts,
                    borderColor: 'rgb(0, 95, 204)',
                    backgroundColor: 'rgba(0, 95, 204, 0.1)',
                    fill: true,
                    tension: 0.1,
                },
                {
                    label: 'Cumulative Gas Operating Cost',
                    data: gasCosts,
                    borderColor: 'rgb(220, 53, 69)',
                    backgroundColor: 'rgba(220, 53, 69, 0.1)',
                    fill: true,
                    tension: 0.1,
                }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                tooltip: { callbacks: { label: (c: any) => `${c.dataset.label}: ${formatCurrency(c.raw)}` } }
            },
            scales: { y: { beginAtZero: true, ticks: { callback: (v: any) => formatCurrency(v) } } }
        }
    });
}

// Line chart for grid decarbonization
function createGridDecarbonizationChart(canvasId: string, data: number[], chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;
    
    if (chartInstance) { chartInstance.destroy(); }

    const labels = Array.from({ length: data.length }, (_, i) => `Year ${i}`);

    return new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Projected Grid Emissions Factor',
                data: data,
                borderColor: '#198754', // Green for environmental
                backgroundColor: 'rgba(25, 135, 84, 0.1)',
                fill: true,
                tension: 0.1,
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                tooltip: { callbacks: { label: (c: any) => `${c.raw.toFixed(3)} kg CO₂e / kWh` } }
            },
            scales: { 
                y: { 
                    beginAtZero: true, 
                    title: { display: true, text: 'kg CO₂e / kWh' }
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
    const formattedGhg20yr = getFormattedGhgParts(ghgReduction20yr, state.ghgUnit);
    summaryGhgReduction20yrEl.textContent = `${formattedGhg20yr.value} ${formattedGhg20yr.label}`;
    summaryGhgReduction20yrEl.className = 'summary-metric-value';
    if(ghgReduction20yr > 0) summaryGhgReduction20yrEl.classList.add('status-sufficient');
    else if (ghgReduction20yr < 0) summaryGhgReduction20yrEl.classList.add('status-danger');

    const ghgReduction100yr = environmentalImpact.annualGhgReductionCo2e100Kg;
    const formattedGhg100yr = getFormattedGhgParts(ghgReduction100yr, state.ghgUnit);
    summaryGhgReduction100yrEl.textContent = `${formattedGhg100yr.value} ${formattedGhg100yr.label}`;
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
        return;
    }

    // Regulatory Context
    const contextSection = document.createElement('div');
    contextSection.className = 'report-section';
    contextSection.innerHTML = `<h3 class="report-section-title">Regulatory Context</h3>`;
    contextSection.innerHTML += `<p>The Massachusetts DPU is managing the transition away from natural gas to meet climate goals (Docket 20-80-B). This influences future utility costs and infrastructure decisions.</p>`;
    reportDetailsEl.appendChild(contextSection);

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
    // Destroy previous chart instances
    if (fullReportLifetimeChart) fullReportLifetimeChart.destroy();
    if (costBreakdownChart) costBreakdownChart.destroy();
    if (emissionsBreakdownChart) emissionsBreakdownChart.destroy();
    if (annualCostsChart) annualCostsChart.destroy();
    if (paybackChart) paybackChart.destroy();
    if (loadContributionChart) loadContributionChart.destroy();
    if (operatingCostChart) operatingCostChart.destroy();
    if (gridDecarbonizationChart) gridDecarbonizationChart.destroy();

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
            <tr><th>Gas Price</th><td>${formatUnitPrice(state.gasPricePerTherm)} / therm</td></tr>
            <tr><th>Electricity Price</th><td>${formatUnitPrice(state.electricityPricePerKwh)} / kWh</td></tr>
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
        planHtml += `<p class="report-note"><strong>Note on Boilers:</strong> Replacing a boiler (which uses hot water/steam pipes) with a heat pump (which produces hot air) requires a new heat distribution system. This estimate includes a ductless mini-split system, a common solution. The 'Installation' cost item in Section 6 combines the labor and materials for both the heat pump unit and the new mini-split distribution system (indoor heads, linesets, etc.) into a single, comprehensive estimate.</p>`;
    }
    
    // --- NEW SECTION: Health, Comfort & Performance ---
    const healthComfortHtml = `
        <h4 class="full-report-subsection-title">Indoor Air Quality (IAQ) Benefits</h4>
        <p>Electrification eliminates the on-site combustion of fossil fuels, a primary source of indoor air pollution linked to negative health outcomes.</p>
        <div class="info-section">
            <div class="icon-container">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M3.75 12h16.5m-16.5 3.75h16.5M3.75 19.5h16.5M5.625 4.5h12.75a1.875 1.875 0 010 3.75H5.625a1.875 1.875 0 010-3.75z" /></svg>
            </div>
            <div class="info-section-content">
                <h5>Eliminates Combustion Byproducts</h5>
                <p>Gas appliances, especially ranges, can release nitrogen dioxide (NO₂), carbon monoxide (CO), and formaldehyde, which can exacerbate respiratory conditions like asthma.</p>
            </div>
        </div>
        <div class="info-section">
            <div class="icon-container">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75L11.25 15 15 9.75m-3-7.036A11.959 11.959 0 013.598 6 11.99 11.99 0 003 9.749c0 5.592 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.31-.21-2.571-.598-3.751h-.152c-3.196 0-6.1-1.248-8.25-3.286zm0 13.036h.008v.008h-.008v-.008z" /></svg>
            </div>
            <div class="info-section-content">
                <h5>Improves Safety</h5>
                <p>Removes the risk of gas leaks and carbon monoxide poisoning associated with the transport and combustion of natural gas within the building.</p>
            </div>
        </div>
        <h4 class="full-report-subsection-title">Comfort & Performance Advantages</h4>
        <div class="info-section">
            <div class="icon-container">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M15.362 5.214A8.252 8.252 0 0112 21 8.25 8.25 0 016.038 7.048 8.287 8.287 0 009 9.62a8.983 8.983 0 013.362-3.877A8.252 8.252 0 0115.362 5.214z" /><path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12a7.5 7.5 0 0015 0h-15z" /></svg>
            </div>
            <div class="info-section-content">
                <h5>Consistent, Even Heat</h5>
                <p>Modern inverter-driven heat pumps provide continuous, low-level heat rather than the "blast-on, blast-off" cycles of some furnaces, leading to more stable and comfortable indoor temperatures.</p>
            </div>
        </div>
        <div class="info-section">
            <div class="icon-container">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M9 6.75V15m6-6v8.25m.5-10.5h-7a.5.5 0 00-.5.5v12a.5.5 0 00.5.5h7a.5.5 0 00.5-.5v-12a.5.5 0 00-.5-.5z" /></svg>
            </div>
            <div class="info-section-content">
                <h5>Zoning Capabilities</h5>
                <p>Systems like ductless mini-splits allow for room-by-room temperature control, increasing occupant comfort and reducing energy waste by only heating or cooling occupied areas.</p>
            </div>
        </div>
        <div class="info-section">
            <div class="icon-container">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M2.25 12l8.954-8.955c.44-.439 1.152-.439 1.591 0L21.75 12M2.25 12a8.963 8.963 0 0117.5 0M2.25 12a8.963 8.963 0 0017.5 0m-17.5 0h17.5" /></svg>
            </div>
            <div class="info-section-content">
                <h5>High-Efficiency Air Conditioning</h5>
                <p>A new heat pump system is also a top-tier central air conditioner. This project provides a significant cooling upgrade in addition to clean heating.</p>
            </div>
        </div>
    `;

    let electricalHtml = `
        <p>The total new electrical load is estimated based on the peak power demand of the new appliances. A standard diversity factor is applied to estimate the actual peak load, as it is unlikely all appliances will run at maximum power simultaneously.</p>
        <div class="formula">$$Amps = \\frac{kW \\times 1000}{Volts}$$</div>
        <table class="full-report-table">
             <thead><tr><th>Appliance</th><th>Peak kW</th><th>Est. Amps @ ${state.voltage}V</th><th>Breaker Spaces</th></tr></thead>
             <tbody>
                ${planningAnalysis.recommendations.map(rec => `<tr><td>${rec.text.replace(/<strong>/g, '').replace(/<\/strong>/g, '')}</td><td>${rec.kw.toFixed(2)} kW</td><td>${rec.amps.toFixed(1)} A</td><td>${rec.breakerSpaces}</td></tr>`).join('')}
                <tr class="total-row">
                    <td colspan="2" style="vertical-align: middle;"><strong>Total Connected Load</strong></td>
                    <td style="vertical-align: middle;"><strong>${planningAnalysis.totalNewAmps.toFixed(1)} A</strong></td>
                    <td style="vertical-align: middle;"><strong>${planningAnalysis.requiredBreakerSpaces}</strong></td>
                </tr>
                 <tr class="total-row">
                    <td colspan="2" style="vertical-align: middle;">
                        <strong>Estimated Peak Load (with Diversity)</strong>
                        <p style="font-size: 0.85em; color: var(--text-muted); margin: 0.5rem 0 0 0; font-weight: 400;">
                            Assumes not all appliances run at peak power simultaneously. Calculated as 100% of the largest load + ${DIVERSITY_FACTOR * 100}% of remaining loads. This value is used for panel sizing recommendations.
                        </p>
                    </td>
                    <td style="vertical-align: middle;"><strong>${planningAnalysis.diversifiedTotalAmps.toFixed(1)} A</strong></td>
                    <td style="vertical-align: middle;"></td>
                </tr>
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
    const { emissionsFactors: ef } = state;
    const upstreamCh4Kg = totalTherms * ef.kgCh4InTherm * ef.upstreamLeakageRate;
    const onsiteCh4Kg = totalTherms * ef.kgCh4InTherm * ef.onsiteSlippageRate;

    const ghgUnitLabel = GHG_CONVERSIONS[state.ghgUnit].label;
    const environmentalHtml = `
        <p>This analysis includes direct carbon dioxide (CO₂) emissions from gas combustion and accounts for methane (CH₄) leakage from the gas supply chain and on-site appliance slippage. Methane is a potent greenhouse gas (GHG), and its impact is measured in CO₂ equivalent (CO₂e) using its Global Warming Potential (GWP) over different time horizons.</p>
        <div class="formula">$$Total\\ CO_2e = (CO_{2\\ combustion}) + (CH_{4\\ leaked} \\times GWP_{CH_4})$$</div>
        <table class="full-report-table">
            <thead>
                <tr><th>Impact Horizon</th><th>Annual Gas Emissions (${ghgUnitLabel})</th><th>Annual Electric Emissions (${ghgUnitLabel})</th><th>Net Annual Reduction (${ghgUnitLabel})</th></tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>20-Year Impact (GWP₂₀ for CH₄ = ${ef.gwp20ch4})</strong></td>
                    <td>${getFormattedGhgParts(environmentalImpact.currentAnnualGasCo2e20Kg, state.ghgUnit).value}</td>
                    <td rowspan="2" style="vertical-align: middle; text-align: center;">${getFormattedGhgParts(environmentalImpact.projectedAnnualElecCo2Kg, state.ghgUnit).value}</td>
                    <td class="total-row"><strong>${getFormattedGhgParts(environmentalImpact.annualGhgReductionCo2e20Kg, state.ghgUnit).value}</strong></td>
                </tr>
                <tr>
                    <td><strong>100-Year Impact (GWP₁₀₀ for CH₄ = ${ef.gwp100ch4})</strong></td>
                    <td>${getFormattedGhgParts(environmentalImpact.currentAnnualGasCo2e100Kg, state.ghgUnit).value}</td>
                    <td class="total-row"><strong>${getFormattedGhgParts(environmentalImpact.annualGhgReductionCo2e100Kg, state.ghgUnit).value}</strong></td>
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
            } else if (item.name.includes('Installation')) {
                addToCategory('Installation & Materials', newItemWithName);
            } else if (item.name.includes('Ductwork')) {
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
    
    const facilityMaintenanceCosts = state.annualMaintenanceCosts[state.facilityType];
    const maintenanceHtml = `
        <h4 class="full-report-subsection-title">Maintenance & Repair Savings</h4>
        <p>This analysis compares the estimated annual maintenance costs for the existing gas system versus the new, lower-maintenance electric system. Electric appliances generally have fewer failure points than combustion appliances.</p>
        <table class="full-report-table maintenance-table">
            <thead>
                <tr><td>System Type</td><td>Low Est. Annual Cost</td><td>High Est. Annual Cost</td></tr>
            </thead>
            <tbody>
                <tr><td>Current Gas System</td><td>${formatCurrency(facilityMaintenanceCosts['Gas System'][0])}</td><td>${formatCurrency(facilityMaintenanceCosts['Gas System'][1])}</td></tr>
                <tr><td>Projected Electric System</td><td>${formatCurrency(facilityMaintenanceCosts['Electric System'][0])}</td><td>${formatCurrency(facilityMaintenanceCosts['Electric System'][1])}</td></tr>
            </tbody>
        </table>
    `;

    const timelineItemsHtml = Object.entries(APPLIANCE_LIFESPANS).map(([name, data]) => {
        const width = (data.years / 25) * 100; // Assuming a 25-year max view
        return `
            <div class="timeline-item">
                <div class="timeline-label">${name}</div>
                <div class="timeline-bar-container">
                    <div class="timeline-bar" style="width: ${width}%; background-color: ${data.color};">
                        ${data.years} years
                    </div>
                </div>
            </div>
        `;
    }).join('');

    const lifespanHtml = `
        <h4 class="full-report-subsection-title">Appliance Lifespan & Replacement Timeline</h4>
        <p>This timeline provides a long-term view of typical appliance replacement cycles, framing this project within the building's future capital planning. Lifespans are estimates and vary by usage and maintenance.</p>
        <div class="timeline-container">${timelineItemsHtml}</div>
    `;


    let rebatesHtml = '';
    if (costAnalysis.rebates.length > 0) {
        let rebateListItems = '';
        if (costAnalysis.rebates.some(r => r.name.includes('Heat Pump'))) {
            rebateListItems += `<li><strong>Whole-Home Heat Pump Rebate (${formatCurrency(state.rebates['Heat Pump'])}):</strong> This is the largest incentive, typically available when you install a heat pump system that serves as the sole source of heating for your entire home, replacing 100% of your fossil fuel system.</li>`;
        }
        if (costAnalysis.rebates.some(r => r.name.includes('HPWH'))) {
            rebateListItems += `<li><strong>Heat Pump Water Heater (HPWH) Rebate (${formatCurrency(state.rebates['HPWH'])}):</strong> An incentive for replacing a gas or standard electric water heater with a high-efficiency heat pump model.</li>`;
        }
        if (costAnalysis.rebates.some(r => r.name.includes('Panel Upgrade'))) {
            rebateListItems += `<li><strong>Panel Upgrade Rebate (${formatCurrency(state.rebates['Panel Upgrade'])}):</strong> This incentive helps offset the cost of upgrading your main electrical panel when it's required for a heat pump installation.</li>`;
        }

        rebatesHtml = `
            <p>The cost analysis includes estimated rebates available through programs like Mass Save. The following provides context on the major incentives included in this plan:</p>
            <ul class="explanation-list">
                ${rebateListItems}
            </ul>
            <p class="report-note"><strong>Disclaimer:</strong> Rebate amounts and eligibility requirements change frequently. This estimate is for planning purposes only. You must confirm your eligibility and apply for rebates directly through the <a href="https://www.masssave.com/rebates" target="_blank" rel="noopener noreferrer">Mass Save website</a> or your utility provider.</p>
        `;
    }

    // --- NEW SECTION: Future Outlook & Risk ---
    const futureRiskHtml = `
        <h4 class="full-report-subsection-title">Energy Price Volatility & Future Risk</h4>
        <p>This project helps mitigate long-term financial risks associated with remaining on the natural gas system.</p>
         <div class="info-section">
            <div class="icon-container">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M2.25 18L9 11.25l4.306 4.307a11.95 11.95 0 015.814-5.519l2.74-1.22m0 0l-5.94-2.28m5.94 2.28l-2.28 5.941" /></svg>
            </div>
            <div class="info-section-content">
                <h5>Reduced Commodity Price Exposure</h5>
                <p>Natural gas is a global commodity subject to significant price swings from geopolitical events. Electricity prices are more closely tied to local generation and regional markets, offering greater price stability.</p>
            </div>
        </div>
        <div class="info-section">
            <div class="icon-container">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M13.19 8.688a4.5 4.5 0 011.242 7.244l-4.5 4.5a4.5 4.5 0 01-6.364-6.364l1.757-1.757m13.35-.622l1.757-1.757a4.5 4.5 0 00-6.364-6.364l-4.5 4.5a4.5 4.5 0 001.242 7.244" /></svg>
            </div>
            <div class="info-section-content">
                <h5>Mitigating "Stranded Asset" Risk</h5>
                <p>As more customers electrify and leave the gas system, the fixed costs of maintaining the pipeline network are spread across fewer users. This could lead to sharp rate hikes for those who remain, a risk highlighted in the "High-Risk" scenario.</p>
            </div>
        </div>
    `;


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
            <h3 class="full-report-section-title">3. Health, Comfort & Performance</h3>
            ${healthComfortHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">4. Electrical Load Analysis</h3>
            ${electricalHtml}
            <h4 class="full-report-subsection-title">Load Contribution by Appliance</h4>
            <div class="chart-container" style="height: ${Math.max(300, planningAnalysis.recommendations.length * 50)}px;">
                <canvas id="load-contribution-chart"></canvas>
            </div>
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">5. Impact on Electric Grid</h3>
            ${distributionHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">6. Cost & Long-Term Planning</h3>
            ${costHtml}
            <h4 class="full-report-subsection-title">Upfront Project Cost Breakdown (Average)</h4>
            <div class="chart-container" style="height: 400px; max-width: 500px; margin: auto;">
                <canvas id="cost-breakdown-chart"></canvas>
            </div>
            ${maintenanceHtml}
            ${lifespanHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">7. Rebate & Incentive Details</h3>
            ${rebatesHtml || '<p>No rebates were applicable to this specific electrification plan.</p>'}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">8. Financial Analysis</h3>
            <h4 class="full-report-subsection-title">Annual Operating Costs</h4>
            <div class="chart-container" style="height: 300px;">
                <canvas id="annual-costs-chart"></canvas>
            </div>
            <h4 class="full-report-subsection-title">Cumulative Cash Flow & Payback Period</h4>
            <p>This chart tracks the cumulative financial return over 15 years, starting with the initial investment. The point where the line crosses $0 is the payback period.</p>
            <div class="chart-container" style="height: 400px;">
                <canvas id="payback-chart"></canvas>
            </div>
            <h4 class="full-report-subsection-title">15-Year Cumulative Operating Costs</h4>
            <p>
                This chart compares the projected total cost of fuel over 15 years, excluding the initial project investment. It assumes a ${state.gasPriceEscalation}% annual increase in gas prices and a ${state.electricityPriceEscalation}% increase in electricity prices.
            </p>
            <div class="chart-container" style="height: 400px;">
                <canvas id="operating-cost-chart"></canvas>
            </div>
            <h4 class="full-report-subsection-title">15-Year Total Cost Outlook (incl. investment)</h4>
            <p>This chart includes the upfront project cost in the electric total, providing a complete picture of the investment over time. It also shows a high-risk scenario for gas price increases.</p>
            <div class="chart-container" style="height: 400px;">
                <canvas id="full-report-lifetime-chart"></canvas>
            </div>
            ${futureRiskHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">9. Environmental Impact & Future Outlook</h3>
            ${environmentalHtml}
            <h4 class="full-report-subsection-title">Annual Emissions Breakdown (100-Year GWP)</h4>
            <div class="chart-container" style="height: 400px;">
                <canvas id="emissions-breakdown-chart"></canvas>
            </div>
            <h4 class="full-report-subsection-title">Grid Decarbonization Trajectory</h4>
            <p>A key benefit of electrification is that the emissions from your building will automatically decrease over time as the electric grid adds more renewable energy, as mandated by state goals. This chart projects this trend based on a conservative annual improvement in the grid's carbon intensity.</p>
            <div class="chart-container" style="height: 300px;">
                <canvas id="grid-decarbonization-chart"></canvas>
            </div>
        </div>
    `;

    // --- RENDER CHARTS ---

    // 1. DATA & RENDER: COST BREAKDOWN DONUT
    const costBreakdown: Record<string, number> = {
        'Equipment': 0,
        'Installation & Materials': 0,
        'Heat Distribution System': 0,
        'Electrical Upgrades': 0,
    };
    costAnalysis.applianceCosts.forEach(items => {
        items.forEach(item => {
            const avgCost = (item.low + item.high) / 2;
            if (item.name === 'Equipment') costBreakdown['Equipment'] += avgCost;
            else if (item.name.includes('Installation')) costBreakdown['Installation & Materials'] += avgCost;
            else if (item.name.includes('Ductwork')) costBreakdown['Heat Distribution System'] += avgCost;
        });
    });
    costAnalysis.electricalCosts.forEach(item => {
        costBreakdown['Electrical Upgrades'] += (item.low + item.high) / 2;
    });
    costBreakdownChart = createCostBreakdownDonutChart('cost-breakdown-chart', costBreakdown, (costAnalysis.netLow + costAnalysis.netHigh) / 2, costBreakdownChart);

    // 2. DATA & RENDER: EMISSIONS BAR CHART (using 100-year GWP)
    const { heatingTherms: ht, nonHeatingTherms: nht } = getThermsForSelection(includedAppliances, state.appliances, state);
    const totalThermsForSelection = ht + nht;
    const gasCombustionCo2 = totalThermsForSelection * state.emissionsFactors.co2KgPerThermGas;
    const upstreamCh4Co2e = (totalThermsForSelection * state.emissionsFactors.kgCh4InTherm * state.emissionsFactors.upstreamLeakageRate) * state.emissionsFactors.gwp100ch4;
    const onsiteCh4Co2e = (totalThermsForSelection * state.emissionsFactors.kgCh4InTherm * state.emissionsFactors.onsiteSlippageRate) * state.emissionsFactors.gwp100ch4;
    const gasEmissionsData = [gasCombustionCo2, upstreamCh4Co2e, onsiteCh4Co2e];
    const elecEmissionsData = environmentalImpact.projectedAnnualElecCo2Kg;
    emissionsBreakdownChart = createEmissionsBarChart('emissions-breakdown-chart', gasEmissionsData, elecEmissionsData, emissionsBreakdownChart);

    // 3. RENDER: ANNUAL COSTS BAR CHART
    annualCostsChart = createAnnualCostsBarChart('annual-costs-chart', financialAnalysis.currentAnnualGasCost, financialAnalysis.projectedAnnualElecCost, annualCostsChart);
    
    // 4. DATA & RENDER: PAYBACK LINE CHART
    const cumulativeCashFlow: number[] = [-(costAnalysis.netLow + costAnalysis.netHigh) / 2];
    let currentGasCost = financialAnalysis.currentAnnualGasCost;
    let currentElecCost = financialAnalysis.projectedAnnualElecCost;
    const gasEscRate = 1 + (state.gasPriceEscalation / 100);
    const elecEscRate = 1 + (state.electricityPriceEscalation / 100);
    for (let i = 0; i < LIFETIME_YEARS; i++) {
        const netSavingsThisYear = currentGasCost - currentElecCost;
        cumulativeCashFlow.push(cumulativeCashFlow[i] + netSavingsThisYear);
        currentGasCost *= gasEscRate;
        currentElecCost *= elecEscRate;
    }
    paybackChart = createPaybackLineChart('payback-chart', cumulativeCashFlow, paybackChart);

    // 5. RENDER: LOAD CONTRIBUTION BAR CHART
    loadContributionChart = createLoadContributionBarChart('load-contribution-chart', planningAnalysis.recommendations, loadContributionChart);
    
    // 6. DATA & RENDER: OPERATING COST LINE CHART
    const cumulativeElecOperatingCosts = lifetimeAnalysis.cumulativeElecCosts.map(c => c - lifetimeAnalysis.cumulativeElecCosts[0]);
    operatingCostChart = createOperatingCostLineChart('operating-cost-chart', lifetimeAnalysis.cumulativeGasCosts, cumulativeElecOperatingCosts, operatingCostChart);

    // 7. RENDER: ORIGINAL LIFETIME (TOTAL COST) CHART
    fullReportLifetimeChart = createLifetimeChart('full-report-lifetime-chart', lifetimeAnalysis, fullReportLifetimeChart);

    // 8. DATA & RENDER: GRID DECARBONIZATION CHART
    const gridDecarbData: number[] = [];
    let currentGridFactor = state.emissionsFactors.co2KgPerKwhMa;
    for (let i = 0; i <= LIFETIME_YEARS; i++) {
        gridDecarbData.push(currentGridFactor);
        currentGridFactor *= (1 - GRID_DECARBONIZATION_RATE_PER_YEAR);
    }
    gridDecarbonizationChart = createGridDecarbonizationChart('grid-decarbonization-chart', gridDecarbData, gridDecarbonizationChart);


    if (window.MathJax) {
        window.MathJax.typesetPromise([fullReportContentEl]).catch((err: any) => console.error('MathJax typesetting error:', err));
    }
}

function renderMethodologyContent() {
    methodologyContentEl.innerHTML = `
        <div class="card methodology-card">
            <h2>Methodology & Data Sources</h2>
            <p>This document details the calculation methodologies, assumptions, and data sources used by the Natural Gas to Electrification Tool. Its purpose is to provide transparency into the tool's operations for facility owners, regulators, and utilities.</p>

            <section>
                <h3>1. Core Calculation: Energy & Efficiency</h3>
                <p>The planner's primary function is to determine the amount of electricity required to replace the <strong>useful heat output</strong> of a gas appliance, not just its input energy. This is critical because gas appliances are not 100% efficient; a significant portion of the fuel's energy is lost as waste heat.</p>
                
                <h4>1.1. Gas Appliance Useful Heat Output</h4>
                <p>First, we calculate the actual heat energy delivered by the existing gas system by accounting for its operational efficiency (e.g., AFUE for furnaces, thermal efficiency for water heaters).</p>
                <div class="formula">$$Useful\\ Heat_{BTU} = Therms_{gas} \\times \\frac{100,000\\ BTU}{Therm} \\times Efficiency_{gas}(\\%)$$</div>
                <p>This <code>Useful Heat</code> value becomes the target amount of energy the new electric appliance must deliver to the conditioned space or water to provide the same level of service.</p>

                <h4>1.2. Electric Appliance Efficiency (COP)</h4>
                <p>Electric heat pumps operate on a different principle than gas combustion or electric resistance. They don't create heat; they move it. Their efficiency is measured by the <strong>Coefficient of Performance (COP)</strong>, which is the ratio of heat energy delivered to the electrical energy consumed.</p>
                <div class="formula">$$COP = \\frac{Heat\\ Delivered}{Electricity\\ Consumed}$$</div>
                <p>A COP of 3.0 means the heat pump delivers 3 units of heat energy for every 1 unit of electrical energy it consumes. The planner uses a lookup table (<code>copLookup</code>) based on appliance type and the selected efficiency tier (Standard, High, Premium).</p>
                <p>For Massachusetts, a climate adjustment is applied for heating systems in colder regions (<code>CLIMATE_ZONE_COP_ADJUSTMENT</code>), reducing the assumed seasonal COP by 10% for Zone 6 (Berkshires) to account for lower average winter temperatures.</p>

                <h4>1.3. Projected Electricity Usage</h4>
                <p>Finally, we calculate the required electrical energy input (in kWh) by determining how much electricity is needed for the new, efficient appliance to deliver the same useful heat as the old gas appliance.</p>
                <div class="formula">$$Electric\\ Input_{kWh} = \\frac{Useful\\ Heat_{BTU} \\times 0.000293\\ \\frac{kWh}{BTU}}{Adjusted\\ COP_{electric}}$$</div>
            </section>

            <section>
                <h3>2. Electrical Load Analysis</h3>
                <p>This analysis estimates the <strong>peak power demand (kW)</strong>, not annual energy usage (kWh). It assesses the maximum instantaneous load the new appliances could place on the building's electrical system, which is crucial for safety and capacity planning.</p>
                
                <h4>2.1. Peak Demand & Amperage</h4>
                <p>The peak power draw in kilowatts (kW) is calculated based on the appliance's rated BTU/hr output. This amperage is then calculated using Ohm's Law for power.</p>
                <div class="formula">$$Peak\\ kW = \\frac{Appliance\\ Rating_{BTU/hr} \\times 0.000293\\ \\frac{kW}{BTU/hr}}{COP_{electric}}$$</div>
                <div class="formula">$$Amps = \\frac{Peak\\ kW \\times 1000}{Voltage}$$</div>

                <h4>2.2. Panel Capacity Assessment</h4>
                <p>The model uses a simplified method to assess if the existing panel is sufficient. It assumes a baseline existing load of 50% of the panel's safe capacity (80% of its rating). The new load is added to this baseline.</p>
                <div class="formula">$$Total\\ Load = (Panel\\ Amps \\times 0.8 \\times 0.5) + New\\ Appliance\\ Amps$$</div>
                <p>If the <code>Total Load</code> exceeds 80% of the panel's rating, an upgrade to the next standard size is recommended. The number of required breaker spaces is also checked against the available spaces to determine if a sub-panel is needed.</p>
            </section>
            
            <section>
                <h3>3. Environmental Impact Analysis</h3>
                <p>The environmental analysis accounts for emissions from both the displaced natural gas and the new electricity consumption. A key feature is the inclusion of methane (CH₄), a potent greenhouse gas.</p>
                
                <h4>3.1. Gas Emissions: CO₂ and Methane</h4>
                <p>Gas emissions are calculated from two components:</p>
                <ul>
                    <li><strong>Carbon Dioxide (CO₂)</strong> from direct combustion.</li>
                    <li><strong>Methane (CH₄)</strong> from upstream leaks in the supply chain and on-site "slippage" (unburned gas).</li>
                </ul>
                
                <h4>3.2. Global Warming Potential (GWP)</h4>
                <p>To compare the impact of different greenhouse gases, a <strong>Global Warming Potential (GWP)</strong> is used. GWP measures how much energy the emissions of 1 ton of a gas will absorb over a given period, relative to 1 ton of CO₂. Methane's GWP is much higher than CO₂'s, especially over shorter time horizons.</p>
                <p>This model uses GWP values from the IPCC Sixth Assessment Report (AR6) for both 20-year and 100-year horizons to show both the short-term and long-term climate impact.</p>
                <div class="formula">$$Total\\ CO_2e = (CO_{2\\ from\\ combustion}) + (CH_{4\\ leaked} \\times GWP_{CH_4})$$</div>
            </section>

            <section>
                <h3>4. Data Sources & References</h3>
                <p>Default values for costs, efficiencies, and emissions are pre-configured within the tool.</p>
                <div class="data-source">
                    <p><strong>Gas Combustion Emissions (CO₂):</strong> U.S. Environmental Protection Agency (EPA), "Greenhouse Gas Inventory Guidance."</p>
                    <p><strong>Methane Leakage Rates:</strong> Based on analysis from the Massachusetts Climate Action Plan.</p>
                    <p><strong>Global Warming Potentials (GWP):</strong> Intergovernmental Panel on Climate Change (IPCC), Sixth Assessment Report (AR6).</p>
                    <p><strong>Grid Emissions Factor (CO₂):</strong> ISO New England 2022 data for the Massachusetts grid.</p>
                    <p><strong>Cost & Rebate Data:</strong> Estimates are based on 2023-2024 market analysis for commercial projects in Massachusetts and Mass Save rebate schedules. <em>These are for planning purposes only and subject to change.</em></p>
                    <p><strong>Regulatory Context:</strong> Massachusetts Department of Public Utilities (DPU) Docket 20-80, "Investigation by the Department on its own Motion into the role of gas local distribution companies as the Commonwealth achieves its 2050 climate goals."</p>
                </div>
            </section>
            
            <section>
                <h3>5. Disclaimer</h3>
                <p class="report-note">This tool is designed for high-level planning and estimation. It is not a substitute for professional engineering analysis, site-specific assessments, or formal quotes from licensed contractors. All costs, savings, and environmental figures are estimates based on the inputs provided and the methodologies described herein. Users must perform their own due diligence and consult with qualified professionals before making any financial or construction decisions.</p>
            </section>
        </div>
    `;
    if (window.MathJax) {
        window.MathJax.typesetPromise([methodologyContentEl]).catch((err: any) => console.error('MathJax typesetting error:', err));
    }
}

function renderSaveLoadContent() {
    saveLoadContentEl.innerHTML = `
        <div class="card">
            <h3 class="card-title">Save/Load Full Project</h3>
            <p>Save or load your entire session, including all facility inputs, appliance selections, and all underlying cost and emissions data. Uses <strong>.json</strong> files.</p>
            <div class="table-actions-grid">
                <button id="save-project-btn" class="button-primary">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor" width="20" height="20"><path d="M10.75 2.75a.75.75 0 00-1.5 0v8.614L6.295 8.235a.75.75 0 10-1.09 1.03l4.25 4.5a.75.75 0 001.09 0l4.25-4.5a.75.75 0 00-1.09-1.03l-2.955 3.129V2.75z" /><path d="M3.5 12.75a.75.75 0 00-1.5 0v2.5A2.75 2.75 0 004.75 18h10.5A2.75 2.75 0 0018 15.25v-2.5a.75.75 0 00-1.5 0v2.5c0 .69-.56 1.25-1.25 1.25H4.75c-.69 0-1.25-.56-1.25-1.25v-2.5z" /></svg>
                    <span>Save Project File</span>
                </button>
                <div class="form-field">
                    <label for="load-project-input" class="button-secondary file-input-label">
                        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor" width="20" height="20"><path d="M9.25 1.75a.75.75 0 01.75.75v8.614l2.955-3.129a.75.75 0 011.09 1.03l-4.25 4.5a.75.75 0 01-1.09 0l-4.25-4.5a.75.75 0 011.09-1.03L9.25 11.114V2.5a.75.75 0 01.75-.75z" /><path d="M3.5 12.75a.75.75 0 00-1.5 0v2.5A2.75 2.75 0 004.75 18h10.5A2.75 2.75 0 0018 15.25v-2.5a.75.75 0 00-1.5 0v2.5c0 .69-.56 1.25-1.25 1.25H4.75c-.69 0-1.25-.56-1.25-1.25v-2.5z" /></svg>
                        <span>Load Project File...</span>
                    </label>
                    <input type="file" id="load-project-input" accept=".json" class="sr-only">
                    <span id="loaded-project-filename" class="loaded-filename">No file selected.</span>
                </div>
            </div>
        </div>
    `;
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

    const currentState = {
        ...state,
        facilityType: facilityTypeEl.value as FacilityType,
        climateZone: climateZoneEl.value as ClimateZone,
        addressNumber: addressNumberEl.value,
        streetName: streetNameEl.value,
        town: townEl.value,
        ghgUnit: ghgUnitEl.value as GhgUnit,
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
    state = currentState;
}

/** Updates all UI input values based on the current state object. */
function updateUIFromState() {
    facilityTypeEl.value = state.facilityType;
    climateZoneEl.value = state.climateZone;
    addressNumberEl.value = state.addressNumber;
    streetNameEl.value = state.streetName;
    townEl.value = state.town;
    ghgUnitEl.value = state.ghgUnit;
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
    
    const allTabs = [tabPlannerBtn, tabReportBtn, tabMethodologyBtn, tabSaveLoadBtn];
    const allContent = [plannerContentEl, reportContentWrapperEl, methodologyContentEl, saveLoadContentEl];

    allTabs.forEach(tab => tab.classList.remove('active'));
    allContent.forEach(content => content.classList.remove('active'));
    
    clickedButton.classList.add('active');

    if (clickedButton === tabPlannerBtn) plannerContentEl.classList.add('active');
    else if (clickedButton === tabReportBtn) reportContentWrapperEl.classList.add('active');
    else if (clickedButton === tabMethodologyBtn) methodologyContentEl.classList.add('active');
    else if (clickedButton === tabSaveLoadBtn) saveLoadContentEl.classList.add('active');
    
    if (clickedButton === tabReportBtn) render();

    if (clickedButton === tabMethodologyBtn && window.MathJax) {
         window.MathJax.typesetPromise([methodologyContentEl]).catch((err: any) => console.error('MathJax typesetting error:', err));
    }
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
    const loadedFilenameEl = document.getElementById('loaded-project-filename') as HTMLSpanElement | null;
    
    if (!input.files || input.files.length === 0) {
        if(loadedFilenameEl) loadedFilenameEl.textContent = 'No file selected.';
        return;
    }

    const file = input.files[0];
    if(loadedFilenameEl) loadedFilenameEl.textContent = file.name;
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
            if(loadedFilenameEl) loadedFilenameEl.textContent = 'Load failed. Please try again.';
        }
    };

    reader.onerror = () => {
        alert('An error occurred while reading the file.');
        if(loadedFilenameEl) loadedFilenameEl.textContent = 'Load failed. Please try again.';
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
tabMethodologyBtn.addEventListener('click', handleTabClick);
tabSaveLoadBtn.addEventListener('click', handleTabClick);

// Export Listener
exportPdfBtn.addEventListener('click', handleExportPdf);

// Save/Load Listeners (delegated)
saveLoadContentEl.addEventListener('click', (e: Event) => {
    const target = e.target as HTMLElement;
    if (target.closest('#save-project-btn')) {
        handleSaveProject();
    }
});
saveLoadContentEl.addEventListener('change', (e: Event) => {
    const target = e.target as HTMLInputElement;
    if (target.id === 'load-project-input') {
        handleLoadProject(e);
    }
});

// Appliance List Listeners
appliancesListEl.addEventListener('click', (e) => {
    const target = e.target as HTMLElement;
    if (target.closest('.remove-btn')) {
        handleRemoveAppliance(e);
    } else if (target.classList.contains('include-in-plan') || (e.target as HTMLInputElement).type === 'number') {
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
renderMethodologyContent();
renderSaveLoadContent();
render();