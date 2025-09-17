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
type ClimateZone = 'Zone5'; // Hardcoded to MA Climate Zone 5
type Region = 'MA'; // Hardcoded to MA region
type EVChargerType = 'none' | '32A' | '40A' | '48A';
type EnvironmentalUnit = 'kg' | 'tons' | 'cars' | 'gasoline' | 'miles' | 'oil' | 'trees';

interface Appliance {
  id: string;
  key: string; // Key into APPLIANCE_DEFINITIONS
  btu: number;
  included: boolean;
  efficiency: number;
  zones?: number; // Number of indoor heads for ductless system (for boilers)
  supplementalHeaters?: number; // Number of electric heaters for small rooms
}

interface State {
  facilityType: FacilityType;
  region: Region;
  climateZone: ClimateZone;
  addressNumber: string;
  streetName: string;
  town: string;
  annualHeatingTherms: number;
  annualNonHeatingTherms: number;
  annualExistingElecKwh: number;
  existingPeakLoadAmps: number;
  panelAmps: number;
  voltage: number;
  breakerSpaces: number;
  evCharger: EVChargerType;
  commercialEVStations: number;
  efficiencyTier: EfficiencyTier;
  appliances: Appliance[];
  gasPricePerTherm: number;
  electricityPricePerKwh: number;
  gasPriceEscalation: number;
  electricityPriceEscalation: number;
  highRiskGasEscalation: number;
  gridDecarbonizationRate: number;
  useTouRates: boolean;
  peakPrice: number;
  offPeakPrice: number;
  offPeakUsagePercent: number;
  environmentalUnit: EnvironmentalUnit;
}

interface ApplianceRecommendation {
    id: string;
    text: string;
    size: string; // e.g., "3.0-ton", "50-gallon"
    btu: number; // BTU/hr output rating of the replacement
    kw: number; // Peak kW for load calculation
    amps: number;
    breakerSpaces: number;
}

interface PlanningAnalysis {
    recommendations: ApplianceRecommendation[];
    sumOfPeakAmps: number;
    diversifiedPeakAmps: number;
    requiredBreakerSpaces: number;
    breakerStatus: 'Sufficient' | 'Sub-panel Recommended' | 'Panel Upgrade Recommended';
    panelStatus: 'Not Required' | `Upgrade to ${number}A Recommended`;
}

interface CostItem {
    name: string;
    low: number;
    medium: number;
    high: number;
}

interface RebateItem {
    name:string;
    amount: number;
}

interface CostAnalysis {
    applianceCosts: Map<string, CostItem[]>; // Map<applianceId, CostItem[]>
    electricalCosts: CostItem[];
    rebates: RebateItem[];
    totalLow: number;
    totalMedium: number;
    totalHigh: number;
    totalRebates: number;
    netLow: number;
    netMedium: number;
    netHigh: number;
}

interface FinancialAnalysis {
    currentAnnualGasCost: number;
    currentAnnualElecCost: number;
    currentAnnualMaintenanceCost: number;
    totalCurrentAnnualCost: number;
    projectedAnnualElecCost: number;
    projectedAnnualMaintenanceCost: number;
    totalProjectedAnnualCost: number;
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
    // For decarbonization scenario
    projectedYear15ElecCo2Kg: number;
    annualGhgReductionCo2e100KgYear15: number;
}

interface LifetimeAnalysis {
    totalGasCost15yr: number;
    totalElecCost15yr: number;
    totalSavings15yr: number;
    cumulativeGasCosts: number[];
    cumulativeElecCostsLow: number[];
    cumulativeElecCostsMedium: number[];
    cumulativeElecCostsHigh: number[];
    totalHighRiskSavings15yr: number;
    cumulativeHighRiskGasCosts: number[];
    annualGasCosts: number[];
    annualElecCosts: number[];
}

interface EnergyCostOutlook {
    cumulativeGasCosts: number[];
    cumulativeElecCosts: number[];
    cumulativeHighRiskGasCosts: number[];
}

interface DistributionImpactAnalysis {
    peakDemandKw: number;
    diversifiedPeakDemandKw: number;
    nonHeatingDemandKw: number;
    diversifiedNonHeatingDemandKw: number;
}


// --- CONSTANTS ---

// --- APPLIANCE DATA LIBRARY ---
const APPLIANCE_DEFINITIONS: Record<string, {name: string, category: ApplianceCategory, defaultBtu: number, defaultEfficiency: number}> = {
    // Residential
    'res-furnace': { name: 'Forced Air Furnace', category: 'Furnace', defaultBtu: 80000, defaultEfficiency: 80 },
    'res-boiler': { name: 'Boiler (Hot Water)', category: 'Boiler', defaultBtu: 100000, defaultEfficiency: 80 },
    'res-steam-boiler': { name: 'Steam Boiler', category: 'Boiler', defaultBtu: 120000, defaultEfficiency: 75 },
    'res-tank-wh': { name: 'Tank Water Heater', category: 'Water Heater', defaultBtu: 40000, defaultEfficiency: 62 },
    'res-tankless-wh': { name: 'Tankless Water Heater', category: 'Water Heater', defaultBtu: 150000, defaultEfficiency: 82 },
    'res-range': { name: 'Gas Range', category: 'Range', defaultBtu: 60000, defaultEfficiency: 40 },
    'res-dryer': { name: 'Gas Dryer', category: 'Dryer', defaultBtu: 22000, defaultEfficiency: 75 },
    'res-pool-heater': { name: 'Pool Heater', category: 'Space Heater', defaultBtu: 250000, defaultEfficiency: 82 },
    'res-fireplace': { name: 'Gas Fireplace Insert', category: 'Space Heater', defaultBtu: 30000, defaultEfficiency: 75 },
    // Commercial / Industrial / Medical / Nursing
    'comm-rtu': { name: 'Rooftop Unit (RTU)', category: 'Furnace', defaultBtu: 240000, defaultEfficiency: 80 },
    'comm-boiler': { name: 'Commercial Boiler (Hot Water)', category: 'Boiler', defaultBtu: 500000, defaultEfficiency: 80 },
    'comm-steam-boiler': { name: 'Commercial Steam Boiler', category: 'Boiler', defaultBtu: 600000, defaultEfficiency: 75 },
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
    'Residential': ['res-furnace', 'res-boiler', 'res-steam-boiler', 'res-tank-wh', 'res-tankless-wh', 'res-range', 'res-dryer', 'res-pool-heater', 'res-fireplace'],
    'Small Commercial': ['comm-rtu', 'comm-boiler', 'comm-steam-boiler', 'comm-wh'],
    'Large Commercial': ['comm-rtu', 'comm-boiler', 'comm-steam-boiler', 'comm-wh', 'laundry-dryer'],
    'Industrial': ['comm-rtu', 'comm-boiler', 'comm-steam-boiler', 'comm-wh', 'laundry-dryer'],
    'Restaurant': ['rest-range', 'rest-convection-oven', 'rest-fryer', 'rest-griddle', 'comm-rtu', 'comm-boiler', 'comm-steam-boiler', 'comm-wh', 'rest-booster-heater'],
    'Medical': ['comm-rtu', 'comm-boiler', 'comm-steam-boiler', 'comm-wh', 'laundry-dryer', 'rest-range', 'rest-convection-oven'],
    'Nursing Home': ['comm-rtu', 'comm-boiler', 'comm-steam-boiler', 'comm-wh', 'laundry-dryer', 'rest-range', 'rest-convection-oven'],
};


const BTU_PER_THERM = 100000;
const BTU_TO_KW = 0.000293071;
const PANEL_SIZES_BY_FACILITY: Record<'Residential' | 'Small Commercial' | 'Large Commercial', number[]> = {
    'Residential': [100, 200, 400],
    'Small Commercial': [200, 400, 600, 800],
    'Large Commercial': [1000, 2000, 2500, 3000, 4000],
};
const LIFETIME_YEARS = 15;
const COLD_CLIMATE_COP_PENALTY = 0.85; // 15% reduction for MA winter performance
const COP_LOOKUP: Record<ApplianceCategory, Record<EfficiencyTier, number>> = {
    'Furnace':      { 'Standard': 2.5, 'High': 3.0, 'Premium': 3.5 },
    'Boiler':       { 'Standard': 2.5, 'High': 3.0, 'Premium': 3.5 },
    'Water Heater': { 'Standard': 2.8, 'High': 3.5, 'Premium': 4.0 },
    'Space Heater': { 'Standard': 2.5, 'High': 3.0, 'Premium': 3.5 }, // Replaced by ductless mini-split
    'Range':        { 'Standard': 1, 'High': 1, 'Premium': 1 }, // Represents efficiency gain over gas
    'Dryer':        { 'Standard': 1.5, 'High': 2.0, 'Premium': 2.5 }, // Represents HP dryer efficiency
};
const CLIMATE_ZONE_COP_ADJUSTMENT: Record<ClimateZone, number> = {
    'Zone5': 1.0,  // Baseline for MA; specific penalty applied in calculation now
};
const APPLIANCE_KW_LOOKUP: Partial<Record<ApplianceCategory, number>> = { Range: 9.6, Dryer: 1.8 }; // Peak kW for non-BTU appliances
const BREAKER_SPACE_REQUIREMENTS: Record<ApplianceCategory, number> = {
    'Furnace': 2, 'Boiler': 2, 'Water Heater': 2, 'Range': 2, 'Dryer': 2, 'Space Heater': 2
};

const EV_CHARGER_SPECS: Record<EVChargerType, { amps: number, kw: number, breakerSpaces: number }> = {
    'none': { amps: 0, kw: 0, breakerSpaces: 0 },
    '32A': { amps: 32, kw: 7.7, breakerSpaces: 2 },
    '40A': { amps: 40, kw: 9.6, breakerSpaces: 2 },
    '48A': { amps: 48, kw: 11.5, breakerSpaces: 2 },
};
const COMMERCIAL_EV_CHARGER_SPEC = { amps: 40, kw: 9.6, breakerSpaces: 2 };

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

const DEFAULT_EXISTING_ELEC_USAGE_BY_FACILITY: Record<FacilityType, number> = {
    'Residential': 8000,
    'Small Commercial': 30000,
    'Large Commercial': 200000,
    'Industrial': 800000,
    'Restaurant': 60000,
    'Medical': 250000,
    'Nursing Home': 220000,
};

const DEFAULT_EV_STATIONS_BY_FACILITY: Record<FacilityType, number> = {
    'Residential': 0, // Not used for residential, but good to be explicit
    'Small Commercial': 2,
    'Large Commercial': 10,
    'Industrial': 4,
    'Restaurant': 4,
    'Medical': 8,
    'Nursing Home': 4,
};


// --- EMISSIONS & REGIONAL DATA ---
const CO2_KG_PER_THERM_GAS = 5.3; // Source: EPA, Combustion only
const UPSTREAM_LEAKAGE_RATE = 0.015; // 1.5% upstream leakage
const ONSITE_SLIPPAGE_RATE = 0.01; // 1.0% on-site slippage/unburned methane
const TOTAL_LEAKAGE_RATE = UPSTREAM_LEAKAGE_RATE + ONSITE_SLIPPAGE_RATE;
const KG_CH4_IN_THERM = 2.36; // Approx. kg of CH4 in 1 therm of natural gas
const GWP20_CH4 = 84; // 20-year Global Warming Potential of Methane (IPCC AR6)
const GWP100_CH4 = 28; // 100-year Global Warming Potential of Methane (IPCC AR6)
const MILES_DRIVEN_PER_KG_CO2 = 2.48; // Source: EPA, Avg passenger vehicle
const KG_CO2E_PER_CAR_YEAR = 4600; // Source: EPA, avg passenger vehicle emits 4.6 metric tons CO2e/year
const GHG_CONVERSION_FACTORS = {
    KG_TO_TONS: 0.001,
    KG_CO2_PER_GALLON_GASOLINE: 8.89,
    KG_CO2E_PER_BARREL_OIL: 430, // Source: EPA
    KG_CO2E_PER_TREE_10_YEARS: 60, // Source: EPA
};

// --- EDITABLE DEFAULT INPUTS ---
// This object holds all the cost and projection data that can be edited in the "Input Tables" tab.
let defaultInputs = {
    REGIONAL_DATA: {
        'MA': {
            name: 'Massachusetts',
            gasPrice: 1.85,
            elecPrice: 0.31,
            co2KgPerKwh: 0.26, // ISO New England
            massSaveRebates: { 'Heat Pump': 10000, 'HPWH': 1000, 'Panel Upgrade': 1500 }
        },
    } as Record<Region, { name: string, gasPrice: number, elecPrice: number, co2KgPerKwh: number, massSaveRebates: Record<string, number> }>,
    FEDERAL_REBATES: {
        'Heat Pump': 2000, // IRA 25C Tax Credit, capped
        'HPWH': 2000,      // IRA 25C Tax Credit, capped
        'Panel Upgrade': 600, // IRA 25C Tax Credit
    } as Record<string, number>,
    EQUIPMENT_COSTS: {
        'Furnace': { 'Standard': [4500, 5750, 7000], 'High': [7000, 8750, 10500], 'Premium': [10500, 13250, 16000] },
        'Boiler': { 'Standard': [4500, 5750, 7000], 'High': [7000, 8750, 10500], 'Premium': [10500, 13250, 16000] },
        'Water Heater': { 'Standard': [2300, 2900, 3500], 'High': [3500, 4050, 4600], 'Premium': [4600, 5200, 5800] },
        'Range': { 'Standard': [1200, 1650, 2100], 'High': [2100, 2800, 3500], 'Premium': [3500, 4650, 5800] },
        'Dryer': { 'Standard': [900, 1150, 1400], 'High': [1400, 1750, 2100], 'Premium': [2100, 2500, 2900] },
        'Space Heater': { 'Standard': [2300, 2900, 3500], 'High': [3500, 4350, 5200], 'Premium': [5200, 6100, 7000] },
    } as Record<ApplianceCategory, Record<EfficiencyTier, number[]>>,
    INSTALLATION_COSTS: {
        'Furnace': [7500, 11250, 15000],
        'Boiler': [7500, 11250, 15000],
        'Water Heater': [1800, 2650, 3500],
        'Range': [500, 750, 1000],
        'Dryer': [500, 750, 1000],
        'Space Heater': [1800, 2500, 3200],
    } as Record<ApplianceCategory, number[]>,
    GAS_APPLIANCE_MAINTENANCE_COSTS: {
        'Furnace': 125, 'Boiler': 200, 'Water Heater': 75, 'Space Heater': 125, 'Range': 20, 'Dryer': 20,
    } as Record<ApplianceCategory, number>,
    ELECTRIC_APPLIANCE_MAINTENANCE_COSTS: {
        'Furnace': 125, 'Boiler': 125, 'Water Heater': 60, 'Space Heater': 125, 'Range': 10, 'Dryer': 50,
    } as Record<ApplianceCategory, number>,
    DISTRIBUTION_SYSTEM_COSTS: {
        'Ductwork Modification': [750, 1625, 2500],
    } as Record<string, number[]>,
    DUCTLESS_MINI_SPLIT_COSTS: {
        base: [5000, 6750, 8500],
        perAdditionalZone: [1800, 2400, 3000],
    } as Record<string, number[]>,
    DECOMMISSIONING_COSTS: {
        'Furnace': [300, 450, 600], 'Boiler': [300, 450, 600], 'Water Heater': [180, 265, 350],
        'Range': [120, 180, 240], 'Dryer': [0, 0, 0], 'Space Heater': [50, 75, 100],
    } as Record<ApplianceCategory, number[]>,
    ELECTRICAL_UPGRADE_COSTS: {
        'Permitting & Fees': [250, 500, 750], 'New Circuit': [700, 1000, 1300], 'EV Charger Equipment': [500, 650, 800],
        'EV Charger Circuit': [1000, 1400, 1800], 'Sub-panel': [1800, 2400, 3000], 'Panel Upgrade to 125A': [2500, 3350, 4200],
        'Panel Upgrade to 150A': [3000, 4000, 5000], 'Panel Upgrade to 200A': [3800, 4900, 6000],
        'Panel Upgrade to 400A': [6000, 8000, 10000], 'Supplemental Heater': [250, 375, 500]
    } as Record<string, number[]>,
};


// --- DOM ELEMENT SELECTORS ---
const facilityTypeEl = document.getElementById('facility-type') as HTMLSelectElement;
const addressNumberEl = document.getElementById('address-number') as HTMLInputElement;
const streetNameEl = document.getElementById('street-name') as HTMLInputElement;
const townEl = document.getElementById('town') as HTMLInputElement;

const annualHeatingThermsEl = document.getElementById('annual-heating-therms') as HTMLInputElement;
const annualNonHeatingThermsEl = document.getElementById('annual-non-heating-therms') as HTMLInputElement;
const annualExistingElecKwhEl = document.getElementById('annual-existing-elec-kwh') as HTMLInputElement;
const existingPeakLoadAmpsEl = document.getElementById('existing-peak-load-amps') as HTMLInputElement;
const panelAmperageEl = document.getElementById('panel-amperage') as HTMLSelectElement;
const voltageEl = document.getElementById('voltage') as HTMLInputElement;
const breakerSpacesEl = document.getElementById('breaker-spaces') as HTMLInputElement;
const efficiencyTierEl = document.getElementById('efficiency-tier') as HTMLSelectElement;
const environmentalUnitEl = document.getElementById('environmental-unit') as HTMLSelectElement;

// EV Charger Inputs
const residentialEvChargerFieldEl = document.getElementById('residential-ev-charger-field') as HTMLDivElement;
const evChargerEl = document.getElementById('ev-charger') as HTMLSelectElement;
const commercialEvStationsFieldEl = document.getElementById('commercial-ev-stations-field') as HTMLDivElement;
const commercialEvStationsEl = document.getElementById('commercial-ev-stations') as HTMLInputElement;


const gasPriceEl = document.getElementById('gas-price') as HTMLInputElement;
const electricityPriceEl = document.getElementById('electricity-price') as HTMLInputElement;
const gasPriceEscalationEl = document.getElementById('gas-price-escalation') as HTMLInputElement;
const electricityPriceEscalationEl = document.getElementById('electricity-price-escalation') as HTMLInputElement;
const highRiskGasEscalationEl = document.getElementById('high-risk-gas-escalation') as HTMLInputElement;
const gridDecarbonizationRateEl = document.getElementById('grid-decarbonization-rate') as HTMLInputElement;
const useTouRatesEl = document.getElementById('use-tou-rates') as HTMLInputElement;
const touRatesGridEl = document.getElementById('tou-rates-grid') as HTMLDivElement;
const peakPriceEl = document.getElementById('peak-price') as HTMLInputElement;
const offPeakPriceEl = document.getElementById('off-peak-price') as HTMLInputElement;
const offPeakUsagePercentEl = document.getElementById('off-peak-percent') as HTMLInputElement;

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
const tabInputTablesBtn = document.getElementById('tab-input-tables') as HTMLButtonElement;
const tabSaveLoadBtn = document.getElementById('tab-save-load') as HTMLButtonElement;
const plannerContentEl = document.getElementById('planner-content') as HTMLDivElement;
const reportContentWrapperEl = document.getElementById('report-content-wrapper') as HTMLDivElement;
const fullReportContentEl = document.getElementById('full-report-content') as HTMLDivElement;
const methodologyContentEl = document.getElementById('methodology-content') as HTMLDivElement;
const inputTablesContentEl = document.getElementById('input-tables-content') as HTMLDivElement;
const saveLoadContentEl = document.getElementById('save-load-content') as HTMLDivElement;

// Project Save/Load/Export elements
const exportPdfBtn = document.getElementById('export-pdf-btn') as HTMLButtonElement;
const saveProjectBtn = document.getElementById('save-project-btn') as HTMLButtonElement;
const loadProjectInput = document.getElementById('load-project-input') as HTMLInputElement;
const loadedFilenameEl = document.getElementById('loaded-filename') as HTMLSpanElement;

// Defaults Save/Load elements
const saveDefaultsBtn = document.getElementById('save-defaults-btn') as HTMLButtonElement;
const loadDefaultsInput = document.getElementById('load-defaults-input') as HTMLInputElement;
const loadedDefaultsFilenameEl = document.getElementById('loaded-defaults-filename') as HTMLSpanElement;


let lifetimeChart: any | null = null;
let fullReportLifetimeChart: any | null = null;
let fullReportEnergyCostChart: any | null = null;
let fullReportCostBreakdownChart: any | null = null;
let fullReportAnnualOperatingCostChart: any | null = null;
let fullReportLoadContributionChart: any | null = null;
let fullReportGhgSourceChart: any | null = null;


// --- APPLICATION STATE ---
let state: State = {
  facilityType: 'Residential',
  region: 'MA',
  climateZone: 'Zone5',
  addressNumber: '45',
  streetName: 'Main Street',
  town: 'Gardner',
  annualHeatingTherms: 700,
  annualNonHeatingTherms: 250,
  annualExistingElecKwh: 8000,
  existingPeakLoadAmps: 30,
  panelAmps: 200,
  voltage: 240,
  breakerSpaces: 4,
  evCharger: 'none',
  commercialEVStations: 0,
  efficiencyTier: 'High',
  appliances: [],
  gasPricePerTherm: 1.85,
  electricityPricePerKwh: 0.31,
  gasPriceEscalation: 2,
  electricityPriceEscalation: 2,
  highRiskGasEscalation: 7,
  gridDecarbonizationRate: 2,
  useTouRates: false,
  peakPrice: 0.35,
  offPeakPrice: 0.18,
  offPeakUsagePercent: 50,
  environmentalUnit: 'kg',
};

// --- CORE LOGIC ---
/**
 * Formats a number for display, limiting it to a maximum number of decimal places.
 * @param value The number to format.
 * @param decimals The maximum number of decimal places.
 * @returns A formatted string or 'N/A'.
 */
function formatNumber(value: number | null | undefined, decimals = 2): string {
    if (value === null || value === undefined || isNaN(value)) {
        return 'N/A';
    }
    return value.toLocaleString(undefined, {
        minimumFractionDigits: 0,
        maximumFractionDigits: decimals,
    });
}


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
            size = `${formatNumber(Math.round(tons * 2) / 2, 1)}-ton`;
            break;
        case 'Boiler':
            typeName = 'Ductless Mini-Split System';
            electricType = 'Heat Pump';
            const boilerTons = appliance.btu / 12000;
            size = `${formatNumber(Math.round(boilerTons * 2) / 2, 1)}-ton`;
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
            size = `${formatNumber(Math.round(spaceHeaterTons * 2) / 2, 1)}-ton`;
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

        let btu = 0;
        if (size.endsWith('-ton')) {
            const tons = parseFloat(size);
            btu = tons * 12000;
        } else if (definition.category === 'Water Heater') {
            // For HPWH, rating is based on output. Let's use the useful heat of the gas unit it replaces.
            btu = app.btu * (app.efficiency / 100);
        } else if (definition.category === 'Range' || definition.category === 'Dryer') {
            // For resistive or other appliances, convert peak kW to BTU/hr
            btu = kw * 3412.14;
        }

        return { id: app.id, text, size, btu, kw, amps, breakerSpaces };
    });

    if (currentState.facilityType === 'Residential' && currentState.evCharger !== 'none') {
        const chargerSpec = EV_CHARGER_SPECS[currentState.evCharger];
        recommendations.push({
            id: 'ev-charger',
            text: `Level 2 EV Charger (${currentState.evCharger})`,
            size: 'N/A',
            btu: 0, // EV chargers don't have a BTU rating for heating/cooling
            kw: chargerSpec.kw,
            amps: chargerSpec.amps,
            breakerSpaces: chargerSpec.breakerSpaces
        });
    } else if (currentState.facilityType !== 'Residential' && currentState.commercialEVStations > 0) {
        for (let i = 0; i < currentState.commercialEVStations; i++) {
            recommendations.push({
                id: `ev-station-${i + 1}`,
                text: `Commercial EV Station #${i + 1} (40A)`,
                size: 'N/A',
                btu: 0,
                kw: COMMERCIAL_EV_CHARGER_SPEC.kw,
                amps: COMMERCIAL_EV_CHARGER_SPEC.amps,
                breakerSpaces: COMMERCIAL_EV_CHARGER_SPEC.breakerSpaces,
            });
        }
    }

    const sumOfPeakAmps = recommendations.reduce((sum, r) => sum + r.amps, 0);
    const requiredBreakerSpaces = recommendations.reduce((sum, r) => sum + r.breakerSpaces, 0);

    // Calculate diversified load
    let diversifiedPeakAmps = 0;
    if (currentState.facilityType === 'Residential' && recommendations.length > 0) {
        // Simplified residential diversity: 100% of largest load + 50% of the rest
        const sortedRecs = [...recommendations].sort((a, b) => b.amps - a.amps);
        const largestLoad = sortedRecs[0];
        const otherLoadsSum = sortedRecs.slice(1).reduce((sum, r) => sum + r.amps, 0);
        diversifiedPeakAmps = largestLoad.amps + (otherLoadsSum * 0.50);
    } else {
        // Simplified commercial diversity: 100% of first 80A, 75% of the rest
        if (sumOfPeakAmps <= 80) {
            diversifiedPeakAmps = sumOfPeakAmps;
        } else {
            diversifiedPeakAmps = 80 + (sumOfPeakAmps - 80) * 0.75;
        }
    }


    const totalCalculatedLoad = currentState.existingPeakLoadAmps + diversifiedPeakAmps;
    const panelCapacity = currentState.panelAmps * 0.8;

    let panelStatus: PlanningAnalysis['panelStatus'] = 'Not Required';
    if (totalCalculatedLoad > panelCapacity) {
        const requiredPanelSize = totalCalculatedLoad / 0.8;
        
        let availableSizes: number[];
        switch (currentState.facilityType) {
            case 'Residential':
                availableSizes = PANEL_SIZES_BY_FACILITY['Residential'];
                break;
            case 'Small Commercial':
            case 'Restaurant':
                availableSizes = PANEL_SIZES_BY_FACILITY['Small Commercial'];
                break;
            case 'Large Commercial':
            case 'Industrial':
            case 'Medical':
            case 'Nursing Home':
            default:
                availableSizes = PANEL_SIZES_BY_FACILITY['Large Commercial'];
                break;
        }

        const recommendedPanel = availableSizes.find(size => size >= requiredPanelSize) || availableSizes.slice(-1)[0];
        panelStatus = `Upgrade to ${recommendedPanel}A Recommended`;
    }

    let breakerStatus: PlanningAnalysis['breakerStatus'] = 'Sufficient';
    if (panelStatus !== 'Not Required') breakerStatus = 'Panel Upgrade Recommended';
    else if (requiredBreakerSpaces > currentState.breakerSpaces) breakerStatus = 'Sub-panel Recommended';

    return { recommendations, sumOfPeakAmps, diversifiedPeakAmps, requiredBreakerSpaces, breakerStatus, panelStatus };
}

function calculateRebates(analysis: PlanningAnalysis, includedAppliances: Appliance[], state: State): RebateItem[] {
    const rebates: RebateItem[] = [];
    if (includedAppliances.length === 0) return rebates;

    const massSaveRebates = defaultInputs.REGIONAL_DATA[state.region].massSaveRebates;
    const federalRebates = defaultInputs.FEDERAL_REBATES;

    const recommendedApplianceTypes = new Set(
        analysis.recommendations
            .filter(rec => !rec.id.startsWith('ev-')) // Exclude EV chargers
            .map(rec => {
                const appliance = state.appliances.find(a => a.id === rec.id)!;
                return getApplianceRecommendation(appliance, state.efficiencyTier).electricType;
            })
    );

    // Mass Save Rebates
    if (recommendedApplianceTypes.has('Heat Pump') && massSaveRebates['Heat Pump']) {
        rebates.push({ name: 'Mass Save Heat Pump Rebate', amount: massSaveRebates['Heat Pump'] });
    }
    if (recommendedApplianceTypes.has('HPWH') && massSaveRebates['HPWH']) {
        rebates.push({ name: 'Mass Save HPWH Rebate', amount: massSaveRebates['HPWH'] });
    }
    if (analysis.panelStatus.startsWith('Upgrade') && recommendedApplianceTypes.has('Heat Pump') && massSaveRebates['Panel Upgrade']) {
         rebates.push({ name: 'Mass Save Panel Upgrade Rebate', amount: massSaveRebates['Panel Upgrade'] });
    }

    // Federal IRA Rebates (25C Tax Credits)
    if (recommendedApplianceTypes.has('Heat Pump') && federalRebates['Heat Pump']) {
        rebates.push({ name: 'Federal IRA Heat Pump Tax Credit', amount: federalRebates['Heat Pump'] });
    }
    if (recommendedApplianceTypes.has('HPWH') && federalRebates['HPWH']) {
        rebates.push({ name: 'Federal IRA HPWH Tax Credit', amount: federalRebates['HPWH'] });
    }
    if (analysis.panelStatus.startsWith('Upgrade') && recommendedApplianceTypes.has('Heat Pump') && federalRebates['Panel Upgrade']) {
         rebates.push({ name: 'Federal IRA Panel Upgrade Tax Credit', amount: federalRebates['Panel Upgrade'] });
    }


    return rebates;
}

function calculateCostAnalysis(analysis: PlanningAnalysis, includedAppliances: Appliance[], state: State): CostAnalysis {
    const applianceCosts = new Map<string, CostItem[]>();
    let totalLow = 0, totalMedium = 0, totalHigh = 0;

    const applianceRecs = analysis.recommendations.filter(r => !r.id.startsWith('ev-'));

    applianceRecs.forEach(rec => {
        const appliance = state.appliances.find(a => a.id === rec.id)!;
        const definition = APPLIANCE_DEFINITIONS[appliance.key];
        const category = definition.category;
        const costs: CostItem[] = [];

        const [equipLow, equipMed, equipHigh] = defaultInputs.EQUIPMENT_COSTS[category][state.efficiencyTier];
        costs.push({ name: 'Equipment', low: equipLow, medium: equipMed, high: equipHigh });

        const [installLow, installMed, installHigh] = defaultInputs.INSTALLATION_COSTS[category];
        costs.push({ name: 'Installation & Materials', low: installLow, medium: installMed, high: installHigh });

        const [decomLow, decomMed, decomHigh] = defaultInputs.DECOMMISSIONING_COSTS[category];
        if (decomHigh > 0) {
            costs.push({ name: 'Gas Decommissioning', low: decomLow, medium: decomMed, high: decomHigh });
        }

        // Add costs for heat distribution system based on original appliance
        if (category === 'Furnace') {
            const [low, medium, high] = defaultInputs.DISTRIBUTION_SYSTEM_COSTS['Ductwork Modification'];
            costs.push({ name: 'Ductwork Modification', low, medium, high });
        } else if (category === 'Boiler' || category === 'Space Heater') {
            const zones = Math.max(1, appliance.zones || 1); // Ensure at least 1 zone
            const [baseLow, baseMed, baseHigh] = defaultInputs.DUCTLESS_MINI_SPLIT_COSTS.base;
            const [perZoneLow, perZoneMed, perZoneHigh] = defaultInputs.DUCTLESS_MINI_SPLIT_COSTS.perAdditionalZone;
        
            const lowCost = baseLow + Math.max(0, zones - 1) * perZoneLow;
            const medCost = baseMed + Math.max(0, zones - 1) * perZoneMed;
            const highCost = baseHigh + Math.max(0, zones - 1) * perZoneHigh;
        
            costs.push({ name: `Ductless Mini-Split System (${zones} zones)`, low: lowCost, medium: medCost, high: highCost });
        }

        applianceCosts.set(rec.id, costs);
    });

    const electricalCosts: CostItem[] = [];

    // Add EV Charger costs first
    const evChargerCount = analysis.recommendations.filter(r => r.id.startsWith('ev-')).length;
    if (evChargerCount > 0) {
        // Equipment hardware cost
        const [equipLow, equipMed, equipHigh] = defaultInputs.ELECTRICAL_UPGRADE_COSTS['EV Charger Equipment'];
        electricalCosts.push({ name: `EV Charging Station(s) (${evChargerCount})`, low: equipLow * evChargerCount, medium: equipMed * evChargerCount, high: equipHigh * evChargerCount });
        // Circuit installation cost
        const [circuitLow, circuitMed, circuitHigh] = defaultInputs.ELECTRICAL_UPGRADE_COSTS['EV Charger Circuit'];
        electricalCosts.push({ name: `EV Charger Circuit(s) (${evChargerCount})`, low: circuitLow * evChargerCount, medium: circuitMed * evChargerCount, high: circuitHigh * evChargerCount });
    }

    // Add panel/sub-panel costs
    if (analysis.panelStatus.startsWith('Upgrade')) {
        const size = analysis.panelStatus.match(/\d+/)?.[0];
        const key = `Panel Upgrade to ${size}A` as keyof typeof defaultInputs.ELECTRICAL_UPGRADE_COSTS;
        const [low, medium, high] = defaultInputs.ELECTRICAL_UPGRADE_COSTS[key];
        electricalCosts.push({ name: 'Main Panel Upgrade', low, medium, high });
    } else if (analysis.breakerStatus === 'Sub-panel Recommended') {
        const [low, medium, high] = defaultInputs.ELECTRICAL_UPGRADE_COSTS['Sub-panel'];
        electricalCosts.push({ name: 'Sub-panel Installation', low, medium, high });
    }

    // Add circuit costs for appliances
    const numApplianceCircuits = applianceRecs.length;
    if (numApplianceCircuits > 0) {
        const [low, medium, high] = defaultInputs.ELECTRICAL_UPGRADE_COSTS['New Circuit'];
        electricalCosts.push({ name: `${numApplianceCircuits} New Appliance Circuit(s)`, low: low * numApplianceCircuits, medium: medium * numApplianceCircuits, high: high * numApplianceCircuits });
    }
    
    // Add supplemental heaters for ductless conversions (boilers, space heaters)
    const appliancesWithHeaters = includedAppliances.filter(app => {
        const category = APPLIANCE_DEFINITIONS[app.key].category;
        return (category === 'Boiler' || category === 'Space Heater');
    });
    const numHeaters = appliancesWithHeaters.reduce((sum, app) => sum + (app.supplementalHeaters || 0), 0);

    if (numHeaters > 0) {
        const [low, medium, high] = defaultInputs.ELECTRICAL_UPGRADE_COSTS['Supplemental Heater'];
        electricalCosts.push({
            name: `Supplemental Heaters (${numHeaters} units)`,
            low: low * numHeaters,
            medium: medium * numHeaters,
            high: high * numHeaters
        });
    }

    // Add permitting costs if any work is being done
    if (applianceRecs.length > 0 || evChargerCount > 0) {
        const [low, medium, high] = defaultInputs.ELECTRICAL_UPGRADE_COSTS['Permitting & Fees'];
        electricalCosts.push({ name: 'Permitting & Fees', low, medium, high });
    }

    // Sum totals
    for(const costs of applianceCosts.values()) {
        costs.forEach(c => { totalLow += c.low; totalMedium += c.medium; totalHigh += c.high; });
    }
    electricalCosts.forEach(c => { totalLow += c.low; totalMedium += c.medium; totalHigh += c.high; });

    // Calculate rebates
    const rebates = calculateRebates(analysis, includedAppliances, state);
    const totalRebates = rebates.reduce((sum, r) => sum + r.amount, 0);

    const netLow = Math.max(0, totalLow - totalRebates);
    const netMedium = Math.max(0, totalMedium - totalRebates);
    const netHigh = Math.max(0, totalHigh - totalRebates);

    return { applianceCosts, electricalCosts, rebates, totalLow, totalMedium, totalHigh, totalRebates, netLow, netMedium, netHigh };
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
                cop *= COLD_CLIMATE_COP_PENALTY; // Apply cold climate penalty for MA
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
    
    // Calculate current costs
    const currentAnnualGasCost = totalTherms * state.gasPricePerTherm;
    // For simplicity, TOU rates are not applied to existing electrical usage, as that usage profile is unknown.
    const currentAnnualElecCost = state.annualExistingElecKwh * state.electricityPricePerKwh;

    // Calculate projected costs
    const totalProjectedAnnualKwh = projectedAnnualKwh + state.annualExistingElecKwh;
    let projectedAnnualElecCost = 0;
    if (state.useTouRates) {
        // A simplifying assumption: the existing load and new load will follow the same peak/off-peak split.
        const offPeakKwh = totalProjectedAnnualKwh * (state.offPeakUsagePercent / 100);
        const peakKwh = totalProjectedAnnualKwh * (1 - (state.offPeakUsagePercent / 100));
        projectedAnnualElecCost = (offPeakKwh * state.offPeakPrice) + (peakKwh * state.peakPrice);
    } else {
        projectedAnnualElecCost = totalProjectedAnnualKwh * state.electricityPricePerKwh;
    }


    // Calculate maintenance costs
    let currentAnnualMaintenanceCost = 0;
    let projectedAnnualMaintenanceCost = 0;
    includedAppliances.forEach(app => {
        const category = APPLIANCE_DEFINITIONS[app.key].category;
        currentAnnualMaintenanceCost += defaultInputs.GAS_APPLIANCE_MAINTENANCE_COSTS[category] || 0;
        projectedAnnualMaintenanceCost += defaultInputs.ELECTRIC_APPLIANCE_MAINTENANCE_COSTS[category] || 0;
    });
    
    // Calculate totals
    const totalCurrentAnnualCost = currentAnnualGasCost + currentAnnualElecCost + currentAnnualMaintenanceCost;
    const totalProjectedAnnualCost = projectedAnnualElecCost + projectedAnnualMaintenanceCost;
    const netAnnualSavings = totalCurrentAnnualCost - totalProjectedAnnualCost;

    let simplePaybackYears: number | null = null;
    if (netAnnualSavings > 0 && costAnalysis.netMedium > 0) {
        simplePaybackYears = costAnalysis.netMedium / netAnnualSavings;
    }

    return {
        currentAnnualGasCost,
        currentAnnualElecCost,
        projectedAnnualElecCost,
        currentAnnualMaintenanceCost,
        projectedAnnualMaintenanceCost,
        totalCurrentAnnualCost,
        totalProjectedAnnualCost,
        netAnnualSavings,
        simplePaybackYears
    };
}

function calculateEnvironmentalImpact(includedAppliances: Appliance[], state: State, projectedAnnualKwh: number): EnvironmentalImpact {
    const { heatingTherms, nonHeatingTherms } = getThermsForSelection(includedAppliances, state.appliances, state);
    const totalTherms = heatingTherms + nonHeatingTherms;

    // Current gas emissions
    const currentAnnualGasCo2Kg = totalTherms * CO2_KG_PER_THERM_GAS;
    const currentAnnualCh4Kg = totalTherms * KG_CH4_IN_THERM * TOTAL_LEAKAGE_RATE;
    const currentAnnualGasCo2e20Kg = currentAnnualGasCo2Kg + (currentAnnualCh4Kg * GWP20_CH4);
    const currentAnnualGasCo2e100Kg = currentAnnualGasCo2Kg + (currentAnnualCh4Kg * GWP100_CH4);

    // Projected electric emissions (Year 1)
    const projectedAnnualElecCo2Kg = projectedAnnualKwh * defaultInputs.REGIONAL_DATA[state.region].co2KgPerKwh;

    // Reductions (Year 1)
    const annualGhgReductionCo2e20Kg = currentAnnualGasCo2e20Kg - projectedAnnualElecCo2Kg;
    const annualGhgReductionCo2e100Kg = currentAnnualGasCo2e100Kg - projectedAnnualElecCo2Kg;

    // Equivalency
    const ghgReductionCarsOffRoad100yr = annualGhgReductionCo2e100Kg / KG_CO2E_PER_CAR_YEAR;

    // Projected electric emissions (Year 15 with decarbonization)
    const decarbonizationFactor = Math.pow(1 - (state.gridDecarbonizationRate / 100), LIFETIME_YEARS - 1);
    const finalGridCo2KgPerKwh = defaultInputs.REGIONAL_DATA[state.region].co2KgPerKwh * decarbonizationFactor;
    const projectedYear15ElecCo2Kg = projectedAnnualKwh * finalGridCo2KgPerKwh;
    const annualGhgReductionCo2e100KgYear15 = currentAnnualGasCo2e100Kg - projectedYear15ElecCo2Kg;

    return { 
        currentAnnualGasCo2Kg, 
        currentAnnualCh4Kg,
        currentAnnualGasCo2e20Kg,
        currentAnnualGasCo2e100Kg,
        projectedAnnualElecCo2Kg,
        annualGhgReductionCo2e20Kg,
        annualGhgReductionCo2e100Kg,
        ghgReductionCarsOffRoad100yr,
        projectedYear15ElecCo2Kg,
        annualGhgReductionCo2e100KgYear15,
    };
}

function calculateLifetimeAnalysis(financialAnalysis: FinancialAnalysis, costAnalysis: CostAnalysis, state: State): LifetimeAnalysis {
    const cumulativeGasCosts: number[] = [0];
    const cumulativeHighRiskGasCosts: number[] = [0];
    const cumulativeElecCostsLow: number[] = [costAnalysis.netLow];
    const cumulativeElecCostsMedium: number[] = [costAnalysis.netMedium];
    const cumulativeElecCostsHigh: number[] = [costAnalysis.netHigh];
    const annualGasCosts: number[] = [];
    const annualElecCosts: number[] = [];

    // Use total annual costs (fuel + maintenance) as the starting point for the escalation calculation.
    let currentYearTotalGasCost = financialAnalysis.totalCurrentAnnualCost;
    let currentYearHighRiskTotalGasCost = financialAnalysis.totalCurrentAnnualCost;
    let currentYearTotalElecCost = financialAnalysis.totalProjectedAnnualCost;

    const gasEscalationRate = 1 + (state.gasPriceEscalation / 100);
    const highRiskGasEscalationRate = 1 + (state.highRiskGasEscalation / 100);
    const elecEscalationRate = 1 + (state.electricityPriceEscalation / 100);

    // We assume maintenance costs escalate at the same rate as the respective utility for simplicity.
    for (let i = 0; i < LIFETIME_YEARS; i++) {
        annualGasCosts.push(currentYearTotalGasCost);
        annualElecCosts.push(currentYearTotalElecCost);

        cumulativeGasCosts.push(cumulativeGasCosts[i] + currentYearTotalGasCost);
        cumulativeHighRiskGasCosts.push(cumulativeHighRiskGasCosts[i] + currentYearHighRiskTotalGasCost);
        cumulativeElecCostsLow.push(cumulativeElecCostsLow[i] + currentYearTotalElecCost);
        cumulativeElecCostsMedium.push(cumulativeElecCostsMedium[i] + currentYearTotalElecCost);
        cumulativeElecCostsHigh.push(cumulativeElecCostsHigh[i] + currentYearTotalElecCost);
        
        currentYearTotalGasCost *= gasEscalationRate;
        currentYearHighRiskTotalGasCost *= highRiskGasEscalationRate;
        currentYearTotalElecCost *= elecEscalationRate;
    }
    
    // Total electric cost for savings calc is operating cost, not including upfront.
    const totalElecOperatingCost15yr = cumulativeElecCostsMedium[LIFETIME_YEARS] - cumulativeElecCostsMedium[0];

    const totalGasCost15yr = cumulativeGasCosts[LIFETIME_YEARS];
    const totalSavings15yr = totalGasCost15yr - totalElecOperatingCost15yr;

    const totalHighRiskGasCost15yr = cumulativeHighRiskGasCosts[LIFETIME_YEARS];
    const totalHighRiskSavings15yr = totalHighRiskGasCost15yr - totalElecOperatingCost15yr;


    return { 
        totalGasCost15yr, 
        totalElecCost15yr: totalElecOperatingCost15yr, 
        totalSavings15yr, 
        cumulativeGasCosts, 
        cumulativeElecCostsLow, 
        cumulativeElecCostsMedium,
        cumulativeElecCostsHigh,
        totalHighRiskSavings15yr,
        cumulativeHighRiskGasCosts,
        annualGasCosts,
        annualElecCosts,
    };
}

function calculateEnergyCostOutlook(financialAnalysis: FinancialAnalysis, state: State): EnergyCostOutlook {
    const cumulativeGasCosts: number[] = [0];
    const cumulativeElecCosts: number[] = [0];
    const cumulativeHighRiskGasCosts: number[] = [0];

    let currentYearGasCost = financialAnalysis.currentAnnualGasCost;
    let currentYearElecCost = financialAnalysis.projectedAnnualElecCost;
    let currentYearHighRiskGasCost = financialAnalysis.currentAnnualGasCost; // Starts same as regular

    const gasEscalationRate = 1 + (state.gasPriceEscalation / 100);
    const elecEscalationRate = 1 + (state.electricityPriceEscalation / 100);
    const highRiskGasEscalationRate = 1 + (state.highRiskGasEscalation / 100);

    for (let i = 0; i < LIFETIME_YEARS; i++) {
        cumulativeGasCosts.push(cumulativeGasCosts[i] + currentYearGasCost);
        cumulativeElecCosts.push(cumulativeElecCosts[i] + currentYearElecCost);
        cumulativeHighRiskGasCosts.push(cumulativeHighRiskGasCosts[i] + currentYearHighRiskGasCost);
        
        currentYearGasCost *= gasEscalationRate;
        currentYearElecCost *= elecEscalationRate;
        currentYearHighRiskGasCost *= highRiskGasEscalationRate;
    }

    return {
        cumulativeGasCosts,
        cumulativeElecCosts,
        cumulativeHighRiskGasCosts,
    };
}


function calculateDistributionImpact(planningAnalysis: PlanningAnalysis, state: State): DistributionImpactAnalysis {
    const nonHeatingCategories: ApplianceCategory[] = ['Water Heater', 'Range', 'Dryer'];

    const heatingRecommendations = planningAnalysis.recommendations.filter(rec => {
        const appliance = state.appliances.find(a => a.id === rec.id);
        return appliance && !nonHeatingCategories.includes(APPLIANCE_DEFINITIONS[appliance.key].category);
    });

    const nonHeatingRecommendations = planningAnalysis.recommendations.filter(rec => {
        const appliance = state.appliances.find(a => a.id === rec.id);
        return appliance && nonHeatingCategories.includes(APPLIANCE_DEFINITIONS[appliance.key].category);
    });

    const peakDemandKw = planningAnalysis.recommendations.reduce((sum, r) => sum + r.kw, 0);
    const nonHeatingDemandKw = nonHeatingRecommendations.reduce((sum, r) => sum + r.kw, 0);

    // Apply diversity factors
    let diversifiedPeakDemandKw = 0;
    if (state.facilityType === 'Residential' && planningAnalysis.recommendations.length > 0) {
        const sortedRecs = [...planningAnalysis.recommendations].sort((a, b) => b.kw - a.kw);
        const largestLoad = sortedRecs[0];
        const otherLoadsSum = sortedRecs.slice(1).reduce((sum, r) => sum + r.kw, 0);
        diversifiedPeakDemandKw = largestLoad.kw + (otherLoadsSum * 0.50);
    } else {
        const kva = peakDemandKw; // Assuming power factor of 1 for simplicity (kW = kVA)
        if (kva <= 20) {
            diversifiedPeakDemandKw = kva;
        } else {
            diversifiedPeakDemandKw = 20 + (kva - 20) * 0.75;
        }
    }

    let diversifiedNonHeatingDemandKw = 0;
     if (state.facilityType === 'Residential' && nonHeatingRecommendations.length > 0) {
        const sortedNonHeating = [...nonHeatingRecommendations].sort((a, b) => b.kw - a.kw);
        const largestLoad = sortedNonHeating[0];
        const otherLoadsSum = sortedNonHeating.slice(1).reduce((sum, r) => sum + r.kw, 0);
        diversifiedNonHeatingDemandKw = largestLoad.kw + (otherLoadsSum * 0.40); // Stricter diversity for non-heating
    } else {
        diversifiedNonHeatingDemandKw = nonHeatingDemandKw * 0.80; // Simple 80% factor for commercial non-heating
    }

    return { peakDemandKw, diversifiedPeakDemandKw, nonHeatingDemandKw, diversifiedNonHeatingDemandKw };
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
                    label: 'Electric System Cost (Likely)',
                    data: analysis.cumulativeElecCostsMedium,
                    borderColor: 'rgb(0, 95, 204)',
                    backgroundColor: 'rgba(0, 95, 204, 0.1)',
                    fill: false,
                    tension: 0.1,
                    borderWidth: 2,
                },
                {
                    label: 'Electric System Cost (High Estimate)',
                    data: analysis.cumulativeElecCostsHigh,
                    borderColor: 'rgba(0, 95, 204, 0.4)',
                    backgroundColor: 'rgba(0, 95, 204, 0.1)',
                    fill: false,
                    tension: 0.1,
                    borderDash: [5, 5],
                    pointRadius: 0,
                },
                 {
                    label: 'Electric System Cost (Low Estimate)',
                    data: analysis.cumulativeElecCostsLow,
                    borderColor: 'rgba(0, 95, 204, 0.4)',
                    backgroundColor: 'rgba(0, 95, 204, 0.1)',
                    fill: '-1',
                    tension: 0.1,
                    borderDash: [5, 5],
                    pointRadius: 0,
                },
                {
                    label: 'Gas System Operating Cost',
                    data: analysis.cumulativeGasCosts,
                    borderColor: 'rgb(220, 53, 69)',
                    backgroundColor: 'rgba(220, 53, 69, 0.1)',
                    fill: false,
                    tension: 0.1,
                },
                {
                    label: 'Gas System Cost (High-Risk)',
                    data: analysis.cumulativeHighRiskGasCosts,
                    borderColor: 'rgb(220, 53, 69)',
                    borderDash: [5, 5],
                    backgroundColor: 'rgba(220, 53, 69, 0.05)',
                    fill: false,
                    tension: 0.1,
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                 tooltip: {
                    mode: 'index',
                    intersect: false,
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

function createEnergyCostChart(canvasId: string, analysis: EnergyCostOutlook, chartInstance: any | null): any {
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
                    label: 'Cumulative Electricity Cost',
                    data: analysis.cumulativeElecCosts,
                    borderColor: 'rgb(0, 95, 204)',
                    backgroundColor: 'rgba(0, 95, 204, 0.1)',
                    fill: true,
                    tension: 0.1,
                },
                {
                    label: 'Cumulative Gas Cost',
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

function createLoadContributionChart(canvasId: string, planningAnalysis: PlanningAnalysis, chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;

    if (chartInstance) {
        chartInstance.destroy();
    }
    
    const labels = planningAnalysis.recommendations.map(r => r.text.replace(/<strong>/g, '').replace(/<\/strong>/g, ''));
    const data = planningAnalysis.recommendations.map(r => r.amps);
    const totalAmps = planningAnalysis.sumOfPeakAmps;

    return new Chart(ctx, {
        type: 'pie',
        data: {
            labels: labels,
            datasets: [{
                label: 'Peak Amps',
                data: data,
                backgroundColor: [
                    'rgba(255, 99, 132, 0.8)', 'rgba(54, 162, 235, 0.8)',
                    'rgba(255, 206, 86, 0.8)', 'rgba(75, 192, 192, 0.8)',
                    'rgba(153, 102, 255, 0.8)', 'rgba(255, 159, 64, 0.8)'
                ],
                borderColor: '#fff',
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: { display: true, text: 'Contribution to Total Peak Load (Sum of Peaks)' },
                tooltip: {
                    callbacks: {
                        label: (context: any) => {
                            const value = context.parsed;
                            if (value === null) return '';
                            const percentage = totalAmps > 0 ? ((value / totalAmps) * 100) : 0;
                            return `${context.label}: ${formatNumber(value, 1)} A (${formatNumber(percentage, 1)}%)`;
                        }
                    }
                }
            }
        }
    });
}

function createCostBreakdownChart(canvasId: string, costAnalysis: CostAnalysis, chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;

    if (chartInstance) {
        chartInstance.destroy();
    }

    const costCategories: Record<string, number> = {
        'Equipment': 0,
        'Installation & Labor': 0,
        'Heat Distribution': 0,
        'Electrical Upgrades': 0,
        'Decommissioning': 0,
    };

    costAnalysis.applianceCosts.forEach((items) => {
        items.forEach(item => {
            const avgCost = item.medium;
            if (item.name === 'Equipment') {
                costCategories['Equipment'] += avgCost;
            } else if (item.name === 'Installation & Materials') {
                costCategories['Installation & Labor'] += avgCost;
            } else if (item.name.includes('Ductless Mini-Split System') || item.name.includes('Ductwork')) {
                costCategories['Heat Distribution'] += avgCost;
            } else if (item.name === 'Gas Decommissioning') {
                costCategories['Decommissioning'] += avgCost;
            }
        });
    });

    costAnalysis.electricalCosts.forEach(item => {
        costCategories['Electrical Upgrades'] += item.medium;
    });

    const labels = Object.keys(costCategories).filter(key => costCategories[key] > 0);
    const data = labels.map(key => costCategories[key]);
    const totalCost = data.reduce((sum, value) => sum + value, 0);

    return new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: [
                    'rgba(0, 95, 204, 0.8)', 'rgba(0, 158, 255, 0.8)',
                    'rgba(83, 166, 255, 0.8)', 'rgba(100, 181, 246, 0.8)',
                    'rgba(144, 202, 249, 0.8)'
                ],
                borderColor: '#fff',
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: { display: true, text: 'Estimated Project Gross Cost Breakdown' },
                tooltip: {
                    callbacks: {
                         label: (context: any) => {
                            const value = context.parsed;
                             if (value === null) return '';
                            const percentage = totalCost > 0 ? ((value / totalCost) * 100) : 0;
                            return `${context.label}: ${formatCurrency(value)} (${formatNumber(percentage, 1)}%)`;
                        }
                    }
                }
            }
        }
    });
}

function createAnnualOperatingCostChart(canvasId: string, lifetimeAnalysis: LifetimeAnalysis, chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;

    if (chartInstance) {
        chartInstance.destroy();
    }

    const labels = Array.from({ length: LIFETIME_YEARS }, (_, i) => `Year ${i + 1}`);

    return new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Gas System Annual Cost',
                    data: lifetimeAnalysis.annualGasCosts,
                    backgroundColor: 'rgba(220, 53, 69, 0.7)',
                    borderColor: 'rgb(220, 53, 69)',
                    borderWidth: 1
                },
                {
                    label: 'Electric System Annual Cost',
                    data: lifetimeAnalysis.annualElecCosts,
                    backgroundColor: 'rgba(0, 95, 204, 0.7)',
                    borderColor: 'rgb(0, 95, 204)',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: { display: true, text: 'Projected Annual Operating Costs (Energy + Maintenance)' },
                tooltip: { callbacks: { label: (context: any) => `${context.dataset.label}: ${formatCurrency(context.parsed.y)}` } }
            },
            scales: {
                x: { stacked: false },
                y: {
                    stacked: false,
                    ticks: { callback: (value: any) => formatCurrency(value) },
                    title: { display: true, text: 'Annual Cost' }
                }
            }
        }
    });
}

function createGhgSourceChart(canvasId: string, environmentalImpact: EnvironmentalImpact, chartInstance: any | null): any {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    if (!canvas) return chartInstance;
    const ctx = canvas.getContext('2d');
    if (!ctx) return chartInstance;

    if (chartInstance) {
        chartInstance.destroy();
    }

    const gasCo2 = environmentalImpact.currentAnnualGasCo2Kg;
    const gasCh4Co2e = environmentalImpact.currentAnnualCh4Kg * GWP20_CH4;
    const elecCo2e = environmentalImpact.projectedAnnualElecCo2Kg;

    return new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Current Gas System', 'Projected Electric System'],
            datasets: [
                {
                    label: 'CO₂ (from combustion)',
                    data: [gasCo2, 0],
                    stack: 'gas',
                    backgroundColor: 'rgba(108, 117, 125, 0.7)',
                },
                {
                    label: 'CH₄ leakage (CO₂e, 20-yr GWP)',
                    data: [gasCh4Co2e, 0],
                    stack: 'gas',
                    backgroundColor: 'rgba(220, 53, 69, 0.7)',
                },
                {
                    label: 'Grid Emissions (CO₂e)',
                    data: [0, elecCo2e],
                    stack: 'electric',
                    backgroundColor: 'rgba(0, 95, 204, 0.7)',
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: { display: true, text: 'Annual Greenhouse Gas Emissions by Source' },
                tooltip: {
                    callbacks: {
                         label: (context: any) => `${context.dataset.label}: ${formatNumber(context.parsed.y, 0)} kg CO₂e`
                    }
                }
            },
            scales: {
                x: { stacked: true },
                y: {
                    stacked: true,
                    title: { display: true, text: 'kg CO₂e' }
                }
            }
        }
    });
}



function renderReport(planningAnalysis: PlanningAnalysis, costAnalysis: CostAnalysis, financialAnalysis: FinancialAnalysis, environmentalImpact: EnvironmentalImpact, lifetimeAnalysis: LifetimeAnalysis) {
    // Summary
    summaryTotalCostEl.textContent = `${formatCurrency(costAnalysis.netMedium)} (Range: ${formatCurrency(costAnalysis.netLow)} - ${formatCurrency(costAnalysis.netHigh)})`;
    summaryTotalRebatesEl.textContent = formatCurrency(costAnalysis.totalRebates);
    summaryAnnualSavingsEl.textContent = formatCurrency(financialAnalysis.netAnnualSavings, true);
    summaryAnnualSavingsEl.className = 'summary-metric-value';
    if(financialAnalysis.netAnnualSavings > 0) summaryAnnualSavingsEl.classList.add('status-sufficient');
    else summaryAnnualSavingsEl.classList.add('status-danger');

    const ghgReduction20yr = environmentalImpact.annualGhgReductionCo2e20Kg;
    const ghgReduction100yr = environmentalImpact.annualGhgReductionCo2e100Kg;

    let ghg20yrValue: number, ghg100yrValue: number, ghgUnitLabel: string;
    
    switch (state.environmentalUnit) {
        case 'tons':
            ghg20yrValue = ghgReduction20yr * GHG_CONVERSION_FACTORS.KG_TO_TONS;
            ghg100yrValue = ghgReduction100yr * GHG_CONVERSION_FACTORS.KG_TO_TONS;
            ghgUnitLabel = 't CO₂e';
            break;
        case 'cars':
            ghg20yrValue = ghgReduction20yr / KG_CO2E_PER_CAR_YEAR;
            ghg100yrValue = ghgReduction100yr / KG_CO2E_PER_CAR_YEAR;
            ghgUnitLabel = 'cars off road';
            break;
        case 'gasoline':
            ghg20yrValue = ghgReduction20yr / GHG_CONVERSION_FACTORS.KG_CO2_PER_GALLON_GASOLINE;
            ghg100yrValue = ghgReduction100yr / GHG_CONVERSION_FACTORS.KG_CO2_PER_GALLON_GASOLINE;
            ghgUnitLabel = 'gallons of gas';
            break;
        case 'miles':
            ghg20yrValue = ghgReduction20yr * MILES_DRIVEN_PER_KG_CO2;
            ghg100yrValue = ghgReduction100yr * MILES_DRIVEN_PER_KG_CO2;
            ghgUnitLabel = 'miles driven';
            break;
        case 'oil':
            ghg20yrValue = ghgReduction20yr / GHG_CONVERSION_FACTORS.KG_CO2E_PER_BARREL_OIL;
            ghg100yrValue = ghgReduction100yr / GHG_CONVERSION_FACTORS.KG_CO2E_PER_BARREL_OIL;
            ghgUnitLabel = 'barrels of oil';
            break;
        case 'trees':
            ghg20yrValue = ghgReduction20yr / GHG_CONVERSION_FACTORS.KG_CO2E_PER_TREE_10_YEARS;
            ghg100yrValue = ghgReduction100yr / GHG_CONVERSION_FACTORS.KG_CO2E_PER_TREE_10_YEARS;
            ghgUnitLabel = 'tree seedlings grown for 10 yrs';
            break;
        case 'kg':
        default:
            ghg20yrValue = ghgReduction20yr;
            ghg100yrValue = ghgReduction100yr;
            ghgUnitLabel = 'kg CO₂e';
            break;
    }

    summaryGhgReduction20yrEl.textContent = `${formatNumber(ghg20yrValue, 2)} ${ghgUnitLabel}`;
    summaryGhgReduction20yrEl.className = 'summary-metric-value';
    if(ghgReduction20yr > 0) summaryGhgReduction20yrEl.classList.add('status-sufficient');
    else if (ghgReduction20yr < 0) summaryGhgReduction20yrEl.classList.add('status-danger');

    summaryGhgReduction100yrEl.textContent = `${formatNumber(ghg100yrValue, 2)} ${ghgUnitLabel}`;
    summaryGhgReduction100yrEl.className = 'summary-metric-value';
    if(ghgReduction100yr > 0) summaryGhgReduction100yrEl.classList.add('status-sufficient');
    else if (ghgReduction100yr < 0) summaryGhgReduction100yrEl.classList.add('status-danger');

    const ghgEquivalent = environmentalImpact.ghgReductionCarsOffRoad100yr;
    summaryGhgEquivalentEl.textContent = `${formatNumber(ghgEquivalent, 2)} cars`;
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
        summaryPaybackPeriodEl.textContent = `${formatNumber(financialAnalysis.simplePaybackYears, 1)} years`;
    } else {
        summaryPaybackPeriodEl.textContent = 'N/A';
    }

    // Details
    reportDetailsEl.innerHTML = '';
    const includedAppliances = state.appliances.filter(a => a.included);
    if (includedAppliances.length === 0 && state.commercialEVStations === 0 && state.evCharger === 'none') {
        reportDetailsEl.innerHTML = '<p class="empty-state">Select one or more appliances in the "Plan" column to generate a report.</p>';
        if (fullReportLifetimeChart) fullReportLifetimeChart.destroy();
        if (fullReportEnergyCostChart) fullReportEnergyCostChart.destroy();
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
            Assumes a ${state.gasPriceEscalation}% annual increase in gas prices and a ${state.electricityPriceEscalation}% increase in electricity prices. This chart shows the total cost of ownership, including upfront costs and ongoing operating expenses, with a range for cost uncertainty.
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
            <li><strong>Explore Incentives:</strong> Visit the <a href="https://www.masssave.com/rebates" target="_blank" rel="noopener noreferrer">Mass Save website</a> for state rebates and research federal IRA tax credits.</li>
            <li><strong>Consider a Home Energy Audit:</strong> An audit can identify other efficiency opportunities.</li>
        </ol>
    `;
    reportDetailsEl.appendChild(nextStepsSection);
}

function renderFullReport(planningAnalysis: PlanningAnalysis, costAnalysis: CostAnalysis, financialAnalysis: FinancialAnalysis, environmentalImpact: EnvironmentalImpact, lifetimeAnalysis: LifetimeAnalysis, distributionImpact: DistributionImpactAnalysis, energyCostOutlook: EnergyCostOutlook) {
    const includedAppliances = state.appliances.filter(a => a.included);
    if (includedAppliances.length === 0 && state.evCharger === 'none' && state.commercialEVStations === 0) {
        fullReportContentEl.innerHTML = '<p class="empty-state">Select one or more appliances on the Planner tab to generate a full report.</p>';
        if (fullReportLifetimeChart) fullReportLifetimeChart.destroy();
        if (fullReportEnergyCostChart) fullReportEnergyCostChart.destroy();
        if (fullReportCostBreakdownChart) fullReportCostBreakdownChart.destroy();
        if (fullReportAnnualOperatingCostChart) fullReportAnnualOperatingCostChart.destroy();
        if (fullReportLoadContributionChart) fullReportLoadContributionChart.destroy();
        if (fullReportGhgSourceChart) fullReportGhgSourceChart.destroy();
        return;
    }
    
    const address = [state.addressNumber, state.streetName, state.town].filter(Boolean).join(' ');
    const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    
    let inputsHtml = `
        <table class="full-report-table">
            <tr><th>Facility Type</th><td>${state.facilityType}</td></tr>
            <tr><th>State/Region</th><td>Massachusetts</td></tr>
            <tr><th>Climate Zone</th><td>Zone 5 (Statewide)</td></tr>
            <tr><th>Heating Gas Usage</th><td>${formatNumber(state.annualHeatingTherms, 0)} therms/yr</td></tr>
            <tr><th>Non-Heating Gas Usage</th><td>${formatNumber(state.annualNonHeatingTherms, 0)} therms/yr</td></tr>
            <tr><th>Existing Annual Electric Usage</th><td>${formatNumber(state.annualExistingElecKwh, 0)} kWh/yr</td></tr>
            <tr><th>Main Panel</th><td>${state.panelAmps}A @ ${state.voltage}V</td></tr>
            <tr><th>Existing Peak Load</th><td>${formatNumber(state.existingPeakLoadAmps, 0)} Amps</td></tr>
            <tr><th>Available Breaker Spaces</th><td>${state.breakerSpaces}</td></tr>
            ${state.facilityType === 'Residential'
                ? `<tr><th>EV Charger Included</th><td>${state.evCharger === 'none' ? 'None' : `Level 2 (${state.evCharger})`}</td></tr>`
                : `<tr><th>Number of EV Stations</th><td>${state.commercialEVStations}</td></tr>`
            }
            <tr><th>Efficiency Target</th><td>${state.efficiencyTier}</td></tr>
            <tr><th>Gas Price</th><td>$${formatNumber(state.gasPricePerTherm, 2)} / therm</td></tr>
            <tr><th>Electricity Price</th><td>${state.useTouRates ? `TOU Rates (Peak: $${formatNumber(state.peakPrice, 2)}, Off-Peak: $${formatNumber(state.offPeakPrice, 2)})` : `$${formatNumber(state.electricityPricePerKwh, 2)} / kWh`}</td></tr>
        </table>
    `;

    const replacedCategories = new Set(planningAnalysis.recommendations.map(rec => {
        const appliance = state.appliances.find(a => a.id === rec.id);
        return appliance ? APPLIANCE_DEFINITIONS[appliance.key].category : '';
    }));
    const hasHeating = replacedCategories.has('Furnace') || replacedCategories.has('Boiler') || replacedCategories.has('Space Heater');
    
    const ductlessAppliances = includedAppliances.filter(a => {
        const category = APPLIANCE_DEFINITIONS[a.key].category;
        return category === 'Boiler' || category === 'Space Heater';
    });
    const needsDuctlessNote = ductlessAppliances.some(a => (a.supplementalHeaters || 0) > 0);


    let planHtml = `
        <table class="full-report-table">
            <thead><tr><th>Gas Appliance</th><th>BTU/hr</th><th>Efficiency</th><th>Electric Replacement</th><th>BTU/hr Output</th></tr></thead>
            <tbody>
            ${planningAnalysis.recommendations
                .filter(rec => !rec.id.startsWith('ev-'))
                .map(rec => {
                const appliance = state.appliances.find(a => a.id === rec.id)!;
                const definition = APPLIANCE_DEFINITIONS[appliance.key];
                return `<tr><td>${definition.name}</td><td>${formatNumber(appliance.btu, 0)}</td><td>${appliance.efficiency}%</td><td>${rec.text.replace(/<strong>/g, '').replace(/<\/strong>/g, '')}</td><td>${rec.btu > 0 ? formatNumber(rec.btu, 0) : 'N/A'}</td></tr>`;
            }).join('')}
            </tbody>
        </table>
    `;

    if (hasHeating) {
        planHtml += `
            <div class="report-note">
                <strong>Performance & Comfort Note for Cold Climates:</strong>
                <br><br>
                This analysis accounts for Massachusetts' cold winters by applying a conservative <strong>${formatNumber((1-COLD_CLIMATE_COP_PENALTY)*100, 0)}% reduction</strong> to the rated efficiency (COP) of all new heat pumps. This provides a more realistic estimate of annual energy consumption and operating costs compared to using manufacturer-rated specifications alone.
            </div>`;
    }

    if (needsDuctlessNote) {
        const totalZones = ductlessAppliances.reduce((sum, app) => sum + (app.zones || 0), 0);
        const totalHeaters = ductlessAppliances.reduce((sum, app) => sum + (app.supplementalHeaters || 0), 0);
        planHtml += `
            <div class="report-note">
                <strong>Critical Consideration for Ductless Systems: Bathroom & Small Room Heating</strong>
                <br><br>
                When replacing a primary heat source with a ductless mini-split system, a new method of heat distribution is required. This plan includes a <strong>${totalZones}-zone ductless system</strong> to heat the primary living areas and bedrooms.
                <br><br>
                However, bathrooms and other small rooms often do not receive a dedicated indoor unit and lose heat quickly due to tile surfaces and ventilation. The original heat source (e.g., a radiator or wall furnace) will be removed, and indirect airflow from a hallway unit is often insufficient to maintain comfort.
                <br><br>
                The standard professional solution is to install small, dedicated electric heaters in these spaces. Therefore, this estimate explicitly includes the cost for installing <strong>${totalHeaters} supplemental electric heaters</strong> (e.g., electric baseboards or wall units) to ensure these rooms remain warm.
            </div>`;
    }
    
    const diversityExplanation = state.facilityType === 'Residential'
        ? 'For residential properties, a diversity factor is applied based on National Electrical Code (NEC) principles where not all appliances run at peak power simultaneously. This estimate uses 100% of the largest load (typically the heat pump) and 50% of the sum of all other new loads.'
        : 'For commercial properties, a diversity factor is applied based on standard electrical guidelines for general loads. This estimate uses 100% of the first 80 Amps of new load and 75% of the remaining load.';

    let electricalHtml = `
        <p>A load calculation estimates the total power demand of the new appliances to ensure the existing electrical panel has sufficient capacity. Simply summing the peak demand of each appliance provides a theoretical maximum, while applying a <strong>diversity factor</strong> provides a more realistic estimate of the actual peak load.</p>
        <table class="full-report-table">
             <thead><tr><th>Appliance</th><th>Peak kW</th><th>Peak Amps @ ${state.voltage}V</th><th>Breaker Spaces</th></tr></thead>
             <tbody>
                ${planningAnalysis.recommendations.map(rec => `<tr><td>${rec.text.replace(/<strong>/g, '').replace(/<\/strong>/g, '')}</td><td>${formatNumber(rec.kw, 2)} kW</td><td>${formatNumber(rec.amps, 1)} A</td><td>${rec.breakerSpaces}</td></tr>`).join('')}
                <tr class="subtotal-row">
                    <td colspan="2"><strong>Sum of Peak Loads (Worst-Case)</strong></td>
                    <td><strong>${formatNumber(planningAnalysis.sumOfPeakAmps, 1)} A</strong></td>
                    <td><strong>${planningAnalysis.requiredBreakerSpaces}</strong></td>
                </tr>
                <tr>
                    <td colspan="4" class="explanation-cell">
                        <p><strong>Diversity Factor Application:</strong><br>${diversityExplanation}</p>
                    </td>
                </tr>
                 <tr class="total-row">
                    <td colspan="2"><strong>Estimated Diversified Peak Load (Realistic)</strong></td>
                    <td><strong>${formatNumber(planningAnalysis.diversifiedPeakAmps, 1)} A</strong></td>
                    <td><strong>${planningAnalysis.requiredBreakerSpaces}</strong></td>
                </tr>
             </tbody>
        </table>
        <h4 class="full-report-subsection-title">Panel Capacity Analysis</h4>
        <p>The diversified load of the new appliances is added to your estimated existing peak load to determine the total projected demand on your electrical panel. This total is compared against 80% of your panel's rated capacity, as required by the National Electrical Code (NEC) for continuous loads.</p>
        <table class="full-report-table">
            <tbody>
                <tr>
                    <td><strong>A. Estimated Existing Peak Load</strong></td>
                    <td style="text-align: right;">${formatNumber(state.existingPeakLoadAmps, 1)} A</td>
                </tr>
                <tr>
                    <td><strong>B. Diversified Load of New Appliances</strong></td>
                    <td style="text-align: right;">${formatNumber(planningAnalysis.diversifiedPeakAmps, 1)} A</td>
                </tr>
                <tr class="subtotal-row">
                    <td><strong>Total Projected Peak Load (A + B)</strong></td>
                    <td style="text-align: right;"><strong>${formatNumber(state.existingPeakLoadAmps + planningAnalysis.diversifiedPeakAmps, 1)} A</strong></td>
                </tr>
                <tr>
                    <td><strong>Panel Capacity (80% of ${state.panelAmps}A)</strong></td>
                    <td style="text-align: right;">${formatNumber(state.panelAmps * 0.8, 1)} A</td>
                </tr>
                <tr class="total-row">
                    <td><strong>Status</strong></td>
                    <td style="text-align: right;"><strong>${planningAnalysis.panelStatus}</strong></td>
                </tr>
            </tbody>
        </table>
        <p style="margin-top: 1rem;"><strong>Breaker Status Recommendation:</strong> ${planningAnalysis.breakerStatus}</p>
        <div class="chart-container" style="height: 350px; margin-top: 1.5rem; max-width: 600px; margin-left: auto; margin-right: auto;">
            <canvas id="full-report-load-contribution-chart"></canvas>
        </div>
    `;

    let distributionHtml = `
        <p>This analysis estimates the project's impact on the local electric distribution system. It shows both the theoretical peak (sum of all new loads) and a more realistic diversified peak, which accounts for the unlikelihood of all appliances running at maximum power simultaneously.</p>
        <table class="full-report-table">
             <thead>
                <tr>
                    <th>Metric</th>
                    <th style="text-align: center;">Sum of Peaks (Worst-Case)</th>
                    <th style="text-align: center;">Estimated Diversified Peak (Realistic)</th>
                </tr>
            </thead>
             <tbody>
                <tr>
                    <td>
                        <strong>Total Peak Demand Increase</strong>
                        <p style="font-size: 0.85em; color: var(--text-muted); margin: 0.5rem 0 0 0;">The maximum potential load added to the grid from all new appliances. This informs capacity planning for extreme weather events.</p>
                    </td>
                    <td style="vertical-align: middle; text-align: center; font-weight: 600;">${formatNumber(distributionImpact.peakDemandKw, 1)} kW</td>
                    <td style="vertical-align: middle; text-align: center; font-weight: 600;" class="total-row">${formatNumber(distributionImpact.diversifiedPeakDemandKw, 1)} kW</td>
                </tr>
                <tr>
                    <td>
                        <strong>Non-Heating Demand Increase</strong>
                        <p style="font-size: 0.85em; color: var(--text-muted); margin: 0.5rem 0 0 0;">The baseline load increase from appliances used year-round (e.g., water heaters, ranges). This informs planning for base load grid capacity.</p>
                    </td>
                    <td style="vertical-align: middle; text-align: center; font-weight: 600;">${formatNumber(distributionImpact.nonHeatingDemandKw, 1)} kW</td>
                     <td style="vertical-align: middle; text-align: center; font-weight: 600;" class="total-row">${formatNumber(distributionImpact.diversifiedNonHeatingDemandKw, 1)} kW</td>
                </tr>
             </tbody>
        </table>
        <p class="report-note"><strong>Note on Grid Impact:</strong> The diversified peak demand provides an estimate of the new load on the local transformer and secondary wires. While this tool does not model time-of-day demand curves, it's important to note that clusters of homes electrifying simultaneously can strain local infrastructure, potentially requiring utility-side upgrades. This is a key area of study for ISO-NE and local utilities as they plan for widespread electrification and the adoption of managed charging and flexible loads.</p>
    `;
    
    // Methane breakdown for environmental report
    const { heatingTherms, nonHeatingTherms } = getThermsForSelection(includedAppliances, state.appliances, state);
    const totalTherms = heatingTherms + nonHeatingTherms;
    const upstreamCh4Kg = totalTherms * KG_CH4_IN_THERM * UPSTREAM_LEAKAGE_RATE;
    const onsiteCh4Kg = totalTherms * KG_CH4_IN_THERM * ONSITE_SLIPPAGE_RATE;
    const emissionsFactor = defaultInputs.REGIONAL_DATA[state.region].co2KgPerKwh;
    const regionName = defaultInputs.REGIONAL_DATA[state.region].name;


    let environmentalHtml = `
        <p>This analysis includes direct carbon dioxide (CO₂) emissions from gas combustion and accounts for methane (CH₄) leakage from the gas supply chain and on-site appliance slippage. Methane is a potent greenhouse gas (GHG), and its impact is measured in CO₂ equivalent (CO₂e) using its Global Warming Potential (GWP) over different time horizons.</p>
        <div class="formula">$$Total\\ CO_2e = (CO_{2\\ combustion}) + (CH_{4\\ leaked} \\times GWP_{CH_4})$$</div>
        <table class="full-report-table">
            <thead>
                <tr><th>Impact Horizon</th><th>Annual Gas Emissions (kg CO₂e)</th><th>Annual Electric Emissions (kg CO₂e)</th><th>Net Annual Reduction (kg CO₂e)</th></tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>20-Year Impact (GWP₂₀ for CH₄ = ${GWP20_CH4})</strong></td>
                    <td>${formatNumber(environmentalImpact.currentAnnualGasCo2e20Kg, 0)}</td>
                    <td rowspan="2" style="vertical-align: middle; text-align: center;">${formatNumber(environmentalImpact.projectedAnnualElecCo2Kg, 0)}</td>
                    <td class="total-row"><strong>${formatNumber(environmentalImpact.annualGhgReductionCo2e20Kg, 0)}</strong></td>
                </tr>
                <tr>
                    <td><strong>100-Year Impact (GWP₁₀₀ for CH₄ = ${GWP100_CH4})</strong></td>
                    <td>${formatNumber(environmentalImpact.currentAnnualGasCo2e100Kg, 0)}</td>
                    <td class="total-row"><strong>${formatNumber(environmentalImpact.annualGhgReductionCo2e100Kg, 0)}</strong></td>
                </tr>
            </tbody>
        </table>
        <p class="report-note">
            This project's annual greenhouse gas reduction is equivalent to taking <strong>${formatNumber(environmentalImpact.ghgReductionCarsOffRoad100yr, 2)}</strong> cars off the road for a year.
        </p>
        <div class="report-note">
            <strong>Future Impact with Grid Decarbonization:</strong>
            <br><br>
            As the ISO New England grid incorporates more renewable energy, the emissions from your electric appliances will automatically decrease. Assuming a <strong>${state.gridDecarbonizationRate}% annual reduction</strong> in grid emissions:
            <ul>
                <li>Your Year 1 electric emissions are <strong>${formatNumber(environmentalImpact.projectedAnnualElecCo2Kg, 0)} kg CO₂e</strong>.</li>
                <li>By Year 15, the annual emissions from the same system would fall to just <strong>${formatNumber(environmentalImpact.projectedYear15ElecCo2Kg, 0)} kg CO₂e</strong>.</li>
                <li>This increases your net annual GHG reduction in Year 15 to <strong>${formatNumber(environmentalImpact.annualGhgReductionCo2e100KgYear15, 0)} kg CO₂e</strong> (100-yr GWP).</li>
            </ul>
        </div>
        <p style="margin-top: 1rem; font-size: 0.9em; color: var(--text-muted);">
            Gas emissions breakdown: <strong>${formatNumber(environmentalImpact.currentAnnualGasCo2Kg, 0)} kg of CO₂</strong> from combustion, <strong>${formatNumber(upstreamCh4Kg, 1)} kg of CH₄</strong> from upstream supply chain leaks, and <strong>${formatNumber(onsiteCh4Kg, 1)} kg of CH₄</strong> from on-site appliance slippage (unburned gas). Electric emissions are based on the average grid intensity for ${regionName} (approx. ${emissionsFactor} kg CO₂e/kWh).
        </p>
        <h4 class="full-report-subsection-title">Annual Emissions Breakdown (20-Year GWP)</h4>
        <div class="chart-container" style="height: 400px; margin-top: 1.5rem;">
            <canvas id="full-report-ghg-source-chart"></canvas>
        </div>
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
            } else if (item.name.includes('Ductless Mini-Split System') || item.name.includes('Ductwork')) {
                addToCategory('Heat Distribution System', { ...item, name: item.name }); // Use original name for this category
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
                    <td style="text-align: right;">${formatCurrency(items[0].medium)}</td>
                    <td style="text-align: right;">${formatCurrency(items[0].high)}</td>
                </tr>
            `;
            // Subsequent rows
            for (let i = 1; i < items.length; i++) {
                costTableBodyHtml += `
                    <tr>
                        <td>${items[i].name}</td>
                        <td style="text-align: right;">${formatCurrency(items[i].low)}</td>
                        <td style="text-align: right;">${formatCurrency(items[i].medium)}</td>
                        <td style="text-align: right;">${formatCurrency(items[i].high)}</td>
                    </tr>
                `;
            }
        }
    });

    let costTableFooterHtml = `
        <tr class="subtotal-row">
            <td colspan="2"><strong>Subtotal</strong></td>
            <td style="text-align: right;"><strong>${formatCurrency(costAnalysis.totalLow)}</strong></td>
            <td style="text-align: right;"><strong>${formatCurrency(costAnalysis.totalMedium)}</strong></td>
            <td style="text-align: right;"><strong>${formatCurrency(costAnalysis.totalHigh)}</strong></td>
        </tr>
    `;

    if (costAnalysis.rebates.length > 0) {
        const items = costAnalysis.rebates;
        costTableFooterHtml += `
            <tr>
                <td rowspan="${items.length}" class="category-cell rebate-cell"><strong>Incentives</strong></td>
                <td>${items[0].name}</td>
                <td style="text-align: right; color: var(--success-color);">(${formatCurrency(items[0].amount)})</td>
                <td style="text-align: right; color: var(--success-color);">(${formatCurrency(items[0].amount)})</td>
                <td style="text-align: right; color: var(--success-color);">(${formatCurrency(items[0].amount)})</td>
            </tr>
        `;
        for (let i = 1; i < items.length; i++) {
            costTableFooterHtml += `
                <tr>
                    <td>${items[i].name}</td>
                    <td style="text-align: right; color: var(--success-color);">(${formatCurrency(items[i].amount)})</td>
                    <td style="text-align: right; color: var(--success-color);">(${formatCurrency(items[i].amount)})</td>
                    <td style="text-align: right; color: var(--success-color);">(${formatCurrency(items[i].amount)})</td>
                </tr>
            `;
        }
    }

    costTableFooterHtml += `
        <tr class="total-row">
            <td colspan="2"><strong>Net Project Cost</strong></td>
            <td style="text-align: right;"><strong>${formatCurrency(costAnalysis.netLow)}</strong></td>
            <td style="text-align: right;"><strong>${formatCurrency(costAnalysis.netMedium)}</strong></td>
            <td style="text-align: right;"><strong>${formatCurrency(costAnalysis.netHigh)}</strong></td>
        </tr>
    `;

    const costDisclaimer = `
        <div class="report-note">
            <p><strong>Disclaimer on Cost Estimates:</strong> All costs presented in this report are for planning purposes only and are based on typical Massachusetts market rates.</p>
            <ul>
                <li><strong>Equipment Costs</strong> are based on a 2024 market analysis of national home improvement retailers and manufacturer pricing.</li>
                <li><strong>Installation & Labor Costs</strong> are based on regional average estimates for licensed and insured contractors in Massachusetts.</li>
            </ul>
            <p>These figures do not account for supply chain issues or the specific complexities of your property. For accurate pricing, it is <strong>essential</strong> to obtain multiple, detailed quotes from qualified local HVAC and electrical professionals.</p>
        </div>
    `;

    let costHtml = `
        ${costDisclaimer}
        <table class="full-report-table">
            <thead>
                <tr>
                    <th>Category</th>
                    <th>Item</th>
                    <th style="text-align: right;">Low Estimate</th>
                    <th style="text-align: right;">Medium Estimate</th>
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
        <h4 class="full-report-subsection-title">Project Gross Cost Breakdown (Based on Medium Estimate)</h4>
        <div class="chart-container" style="height: 350px; max-width: 600px; margin: 1.5rem auto 0 auto;">
            <canvas id="full-report-cost-breakdown-chart"></canvas>
        </div>
    `;
    // --- End of New Cost Analysis Rendering ---

    let rebatesHtml = '';
    if (costAnalysis.rebates.length > 0) {
        let rebateListItems = '';
        const massSaveRebates = defaultInputs.REGIONAL_DATA[state.region].massSaveRebates;
        const federalRebates = defaultInputs.FEDERAL_REBATES;

        if (costAnalysis.rebates.some(r => r.name.includes('Mass Save Heat Pump'))) {
            rebateListItems += `<li><strong>Mass Save Whole-Home Heat Pump Rebate (${formatCurrency(massSaveRebates['Heat Pump'])}):</strong> This is the largest state incentive, typically available when you install a heat pump system that serves as the sole source of heating for your entire home.</li>`;
        }
        if (costAnalysis.rebates.some(r => r.name.includes('Federal IRA Heat Pump'))) {
            rebateListItems += `<li><strong>Federal IRA 25C Tax Credit for Heat Pumps (up to ${formatCurrency(federalRebates['Heat Pump'])}):</strong> A federal tax credit available for qualifying high-efficiency heat pump installations.</li>`;
        }
        if (costAnalysis.rebates.some(r => r.name.includes('HPWH'))) {
            rebateListItems += `<li><strong>State & Federal HPWH Incentives:</strong> Combined incentives are available for replacing a gas or standard electric water heater with a high-efficiency heat pump model.</li>`;
        }
        if (costAnalysis.rebates.some(r => r.name.includes('Panel Upgrade'))) {
            rebateListItems += `<li><strong>State & Federal Panel Upgrade Incentives:</strong> These incentives help offset the cost of upgrading your main electrical panel when it's required for a heat pump installation.</li>`;
        }

        rebatesHtml = `
            <p>The cost analysis includes estimated incentives available through state programs like Mass Save and federal programs like the Inflation Reduction Act (IRA). The following provides context on the major incentives included in this plan:</p>
            <ul class="explanation-list">
                ${rebateListItems}
            </ul>
            <p class="report-note"><strong>Disclaimer:</strong> Rebate and tax credit amounts and eligibility requirements change frequently. This estimate is for planning purposes only. You must confirm your eligibility and apply for incentives directly through the <a href="https://www.masssave.com/rebates" target="_blank" rel="noopener noreferrer">Mass Save website</a> and consult a tax professional regarding federal credits.</p>
        `;
    }

    let operatingCostHtml = `
        <p>This analysis compares the total estimated annual utility bills before and after electrification.</p>
        <table class="full-report-table">
            <thead>
                <tr>
                    <th></th>
                    <th style="text-align: right;">Before Electrification</th>
                    <th style="text-align: right;">After Electrification</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>Annual Gas Bill</td>
                    <td style="text-align: right;">${formatCurrency(financialAnalysis.currentAnnualGasCost)}</td>
                    <td style="text-align: right;">$0</td>
                </tr>
                <tr>
                    <td>Annual Electric Bill</td>
                    <td style="text-align: right;">${formatCurrency(financialAnalysis.currentAnnualElecCost)}</td>
                    <td style="text-align: right;">${formatCurrency(financialAnalysis.projectedAnnualElecCost)}</td>
                </tr>
                <tr class="subtotal-row">
                    <td><strong>Total Annual Utility Cost (Energy Only)</strong></td>
                    <td style="text-align: right;"><strong>${formatCurrency(financialAnalysis.currentAnnualGasCost + financialAnalysis.currentAnnualElecCost)}</strong></td>
                    <td style="text-align: right;"><strong>${formatCurrency(financialAnalysis.projectedAnnualElecCost)}</strong></td>
                </tr>
            </tbody>
        </table>
        
        <h4 class="full-report-subsection-title">Operating Cost Component Breakdown</h4>
        <p>This table breaks down the total operating costs, including both energy and estimated routine maintenance, to determine the net annual savings.</p>
        <table class="full-report-table">
            <thead>
                <tr>
                    <th>Cost Component</th>
                    <th style="text-align: right;">Current System (Gas + Electric)</th>
                    <th style="text-align: right;">Projected System (All-Electric)</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>Annual Energy Cost</td>
                    <td style="text-align: right;">${formatCurrency(financialAnalysis.currentAnnualGasCost + financialAnalysis.currentAnnualElecCost)}</td>
                    <td style="text-align: right;">${formatCurrency(financialAnalysis.projectedAnnualElecCost)}</td>
                </tr>
                <tr>
                    <td>Est. Annual Maintenance</td>
                    <td style="text-align: right;">${formatCurrency(financialAnalysis.currentAnnualMaintenanceCost)}</td>
                    <td style="text-align: right;">${formatCurrency(financialAnalysis.projectedAnnualMaintenanceCost)}</td>
                </tr>
                <tr class="subtotal-row">
                    <td><strong>Total Annual Operating Cost</strong></td>
                    <td style="text-align: right;"><strong>${formatCurrency(financialAnalysis.totalCurrentAnnualCost)}</strong></td>
                    <td style="text-align: right;"><strong>${formatCurrency(financialAnalysis.totalProjectedAnnualCost)}</strong></td>
                </tr>
            </tbody>
            <tfoot>
                <tr class="total-row">
                    <td colspan="2"><strong>Net Annual Savings</strong></td>
                    <td style="text-align: right;"><strong>${formatCurrency(financialAnalysis.netAnnualSavings)}</strong></td>
                </tr>
            </tfoot>
        </table>
        <p class="report-note">
            <strong>Note on Maintenance Costs:</strong> Maintenance costs are estimates based on industry averages for annual tune-ups and inspections. Actual costs may vary. Gas appliances typically require annual safety checks of combustion components, while electric heat pumps require regular filter changes and coil cleaning.
        </p>
        <h4 class="full-report-subsection-title">15-Year Annual Operating Cost Projection</h4>
        <div class="chart-container" style="height: 400px; margin-top: 1.5rem;">
           <canvas id="full-report-annual-operating-cost-chart"></canvas>
       </div>
    `;

    const energyCostOutlookHtml = `
        <p>
            This chart isolates the projected cumulative cost of energy over 15 years, comparing only the escalating price of gas versus electricity. It does not include upfront installation or ongoing maintenance costs, providing a clear view of the long-term fuel cost trends.
        </p>
        <div class="chart-container" style="height: 400px;">
            <canvas id="full-report-energy-cost-chart"></canvas>
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
            <h3 class="full-report-section-title">3. Electrical Load Analysis</h3>
            ${electricalHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">4. Impact on Electric Grid</h3>
            ${distributionHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">5. Project Cost Analysis</h3>
            ${costHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">6. Rebate & Incentive Details</h3>
            ${rebatesHtml || '<p>No incentives were applicable to this specific electrification plan.</p>'}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">7. Annual Operating Cost Analysis</h3>
            ${operatingCostHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">8. 15-Year Energy Cost Comparison</h3>
            ${energyCostOutlookHtml}
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">9. 15-Year Financial Outlook (Total Cost of Ownership)</h3>
             <p>
                This chart projects the cumulative cost of ownership over 15 years, including upfront installation costs and ongoing, escalating operating costs (energy and maintenance). It shows a likely cost scenario along with a range representing low and high cost estimates to capture uncertainty.
            </p>
            <div class="chart-container" style="height: 400px;">
                <canvas id="full-report-lifetime-chart"></canvas>
            </div>
        </div>
        <div class="full-report-section">
            <h3 class="full-report-section-title">10. Environmental Impact Analysis</h3>
            ${environmentalHtml}
        </div>
    `;

    fullReportLifetimeChart = createLifetimeChart('full-report-lifetime-chart', lifetimeAnalysis, fullReportLifetimeChart);
    fullReportEnergyCostChart = createEnergyCostChart('full-report-energy-cost-chart', energyCostOutlook, fullReportEnergyCostChart);
    fullReportLoadContributionChart = createLoadContributionChart('full-report-load-contribution-chart', planningAnalysis, fullReportLoadContributionChart);
    fullReportCostBreakdownChart = createCostBreakdownChart('full-report-cost-breakdown-chart', costAnalysis, fullReportCostBreakdownChart);
    fullReportAnnualOperatingCostChart = createAnnualOperatingCostChart('full-report-annual-operating-cost-chart', lifetimeAnalysis, fullReportAnnualOperatingCostChart);
    fullReportGhgSourceChart = createGhgSourceChart('full-report-ghg-source-chart', environmentalImpact, fullReportGhgSourceChart);


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
    
    const heatDistributionFields = applianceEntryEl.querySelector('.heat-distribution-fields') as HTMLDivElement;
    const zonesInput = applianceEntryEl.querySelector('.zones-rating') as HTMLInputElement;
    const supplementalHeatersInput = applianceEntryEl.querySelector('.supplemental-heaters-rating') as HTMLInputElement;

    nameEl.textContent = definition.name;
    btuInput.value = appliance.btu > 0 ? appliance.btu.toString() : '';
    efficiencyInput.value = appliance.efficiency > 0 ? appliance.efficiency.toString() : '';

    if (definition.category === 'Boiler' || definition.category === 'Space Heater') {
        heatDistributionFields.classList.remove('hidden');
        zonesInput.value = (appliance.zones || 3).toString();
        supplementalHeatersInput.value = (appliance.supplementalHeaters ?? 2).toString();
    } else {
        heatDistributionFields.classList.add('hidden');
    }

    const recommendation = planningAnalysis.recommendations.find(r => r.id === appliance.id);
    if (recommendation) {
        recommendationEl.innerHTML = `${recommendation.text} <span>(Est. ${formatNumber(recommendation.amps, 1)} A / ${recommendation.breakerSpaces} spaces)</span>`;
    } else {
        recommendationEl.innerHTML = '';
    }

    appliancesListEl.appendChild(applianceEntryEl);
  });
}

function renderInputTables() {
    const createInput = (value: number, type: 'number' | 'text' = 'number', dataset: Record<string, string> = {}) => {
        const input = document.createElement('input');
        input.type = type;
        input.value = value.toString();
        Object.entries(dataset).forEach(([key, val]) => input.dataset[key] = val);
        return input;
    };

    // Utility Costs
    const utilityTbody = document.getElementById('defaults-utility-tbody') as HTMLTableSectionElement;
    utilityTbody.innerHTML = `
        <tr><td>Gas Price ($/therm)</td><td></td></tr>
        <tr><td>Elec Price ($/kWh)</td><td></td></tr>
        <tr><td>Gas Price Escalation (%/yr)</td><td></td></tr>
        <tr><td>Elec Price Escalation (%/yr)</td><td></td></tr>
        <tr><td>High-Risk Gas Escalation (%/yr)</td><td></td></tr>
        <tr><td>Grid Decarbonization Rate (%/yr)</td><td></td></tr>
    `;
    utilityTbody.rows[0].cells[1].appendChild(createInput(state.gasPricePerTherm, 'number', { key: 'gasPricePerTherm' }));
    utilityTbody.rows[1].cells[1].appendChild(createInput(state.electricityPricePerKwh, 'number', { key: 'electricityPricePerKwh' }));
    utilityTbody.rows[2].cells[1].appendChild(createInput(state.gasPriceEscalation, 'number', { key: 'gasPriceEscalation' }));
    utilityTbody.rows[3].cells[1].appendChild(createInput(state.electricityPriceEscalation, 'number', { key: 'electricityPriceEscalation' }));
    utilityTbody.rows[4].cells[1].appendChild(createInput(state.highRiskGasEscalation, 'number', { key: 'highRiskGasEscalation' }));
    utilityTbody.rows[5].cells[1].appendChild(createInput(state.gridDecarbonizationRate, 'number', { key: 'gridDecarbonizationRate' }));

    // Equipment Costs
    const equipmentTbody = document.getElementById('defaults-equipment-tbody') as HTMLTableSectionElement;
    equipmentTbody.innerHTML = '';
    for (const category in defaultInputs.EQUIPMENT_COSTS) {
        const row = equipmentTbody.insertRow();
        row.insertCell().textContent = category;
        const tiers: EfficiencyTier[] = ['Standard', 'High', 'Premium'];
        tiers.forEach(tier => {
            const [low, medium, high] = defaultInputs.EQUIPMENT_COSTS[category as ApplianceCategory][tier];
            row.insertCell().appendChild(createInput(low, 'number', { category, tier, range: 'low' }));
            row.insertCell().appendChild(createInput(medium, 'number', { category, tier, range: 'medium' }));
            row.insertCell().appendChild(createInput(high, 'number', { category, tier, range: 'high' }));
        });
    }
    
    // Installation Costs
    const installationTbody = document.getElementById('defaults-installation-tbody') as HTMLTableSectionElement;
    installationTbody.innerHTML = '';
    for (const category in defaultInputs.INSTALLATION_COSTS) {
        const row = installationTbody.insertRow();
        row.insertCell().textContent = category;
        const [low, medium, high] = defaultInputs.INSTALLATION_COSTS[category as ApplianceCategory];
        row.insertCell().appendChild(createInput(low, 'number', { category, range: 'low' }));
        row.insertCell().appendChild(createInput(medium, 'number', { category, range: 'medium' }));
        row.insertCell().appendChild(createInput(high, 'number', { category, range: 'high' }));
    }

    // Distribution Costs
    const distributionTbody = document.getElementById('defaults-distribution-tbody') as HTMLTableSectionElement;
    distributionTbody.innerHTML = '';
    Object.entries({ ...defaultInputs.DISTRIBUTION_SYSTEM_COSTS, ...defaultInputs.DUCTLESS_MINI_SPLIT_COSTS }).forEach(([key, [low, medium, high]]) => {
        const row = distributionTbody.insertRow();
        row.insertCell().textContent = key;
        row.insertCell().appendChild(createInput(low, 'number', { key, range: 'low' }));
        row.insertCell().appendChild(createInput(medium, 'number', { key, range: 'medium' }));
        row.insertCell().appendChild(createInput(high, 'number', { key, range: 'high' }));
    });
    
    // Decommissioning Costs
    const decommissioningTbody = document.getElementById('defaults-decommissioning-tbody') as HTMLTableSectionElement;
    decommissioningTbody.innerHTML = '';
    for (const category in defaultInputs.DECOMMISSIONING_COSTS) {
        const row = decommissioningTbody.insertRow();
        row.insertCell().textContent = category;
        const [low, medium, high] = defaultInputs.DECOMMISSIONING_COSTS[category as ApplianceCategory];
        row.insertCell().appendChild(createInput(low, 'number', { category, range: 'low' }));
        row.insertCell().appendChild(createInput(medium, 'number', { category, range: 'medium' }));
        row.insertCell().appendChild(createInput(high, 'number', { category, range: 'high' }));
    }

    // Electrical Costs
    const electricalTbody = document.getElementById('defaults-electrical-tbody') as HTMLTableSectionElement;
    electricalTbody.innerHTML = '';
    for (const key in defaultInputs.ELECTRICAL_UPGRADE_COSTS) {
        const row = electricalTbody.insertRow();
        row.insertCell().textContent = key;
        const [low, medium, high] = defaultInputs.ELECTRICAL_UPGRADE_COSTS[key];
        row.insertCell().appendChild(createInput(low, 'number', { key, range: 'low' }));
        row.insertCell().appendChild(createInput(medium, 'number', { key, range: 'medium' }));
        row.insertCell().appendChild(createInput(high, 'number', { key, range: 'high' }));
    }

    // Incentives
    const incentivesTbody = document.getElementById('defaults-incentives-tbody') as HTMLTableSectionElement;
    incentivesTbody.innerHTML = '';
    const incentiveKeys = Object.keys(defaultInputs.REGIONAL_DATA.MA.massSaveRebates);
    incentiveKeys.forEach(key => {
        const row = incentivesTbody.insertRow();
        row.insertCell().textContent = key;
        row.insertCell().appendChild(createInput(defaultInputs.REGIONAL_DATA.MA.massSaveRebates[key], 'number', { source: 'massSave', key }));
        row.insertCell().appendChild(createInput(defaultInputs.FEDERAL_REBATES[key] || 0, 'number', { source: 'federal', key }));
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
  const energyCostOutlook = calculateEnergyCostOutlook(financialAnalysis, state);
  
  renderAppliances();
  renderReport(planningAnalysis, costAnalysis, financialAnalysis, environmentalImpact, lifetimeAnalysis);
  renderFullReport(planningAnalysis, costAnalysis, financialAnalysis, environmentalImpact, lifetimeAnalysis, distributionImpact, energyCostOutlook);
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

    const heatDistributionFields = row.querySelector('.heat-distribution-fields') as HTMLDivElement;
    let zones: number | undefined = undefined;
    let supplementalHeaters: number | undefined = undefined;

    if (heatDistributionFields && !heatDistributionFields.classList.contains('hidden')) {
        zones = parseInt((row.querySelector('.zones-rating') as HTMLInputElement).value, 10) || undefined;
        supplementalHeaters = parseInt((row.querySelector('.supplemental-heaters-rating') as HTMLInputElement).value, 10);
        if (isNaN(supplementalHeaters)) supplementalHeaters = 0;
    }

    newAppliances.push({ id, key: existingAppliance.key, btu, included, efficiency, zones, supplementalHeaters });
  });

  state = {
    ...state, // Keep hardcoded region and climateZone
    facilityType: facilityTypeEl.value as FacilityType,
    addressNumber: addressNumberEl.value,
    streetName: streetNameEl.value,
    town: townEl.value,
    annualHeatingTherms: parseInt(annualHeatingThermsEl.value, 10) || 0,
    annualNonHeatingTherms: parseInt(annualNonHeatingThermsEl.value, 10) || 0,
    annualExistingElecKwh: parseInt(annualExistingElecKwhEl.value, 10) || 0,
    existingPeakLoadAmps: parseInt(existingPeakLoadAmpsEl.value, 10) || 0,
    panelAmps: parseInt(panelAmperageEl.value, 10),
    voltage: parseInt(voltageEl.value, 10) || 240,
    breakerSpaces: parseInt(breakerSpacesEl.value, 10) || 0,
    evCharger: evChargerEl.value as EVChargerType,
    commercialEVStations: parseInt(commercialEvStationsEl.value, 10) || 0,
    efficiencyTier: efficiencyTierEl.value as EfficiencyTier,
    appliances: newAppliances,
    gasPricePerTherm: parseFloat(gasPriceEl.value) || 0,
    electricityPricePerKwh: parseFloat(electricityPriceEl.value) || 0,
    gasPriceEscalation: parseFloat(gasPriceEscalationEl.value) || 0,
    electricityPriceEscalation: parseFloat(electricityPriceEscalationEl.value) || 0,
    highRiskGasEscalation: parseFloat(highRiskGasEscalationEl.value) || 0,
    gridDecarbonizationRate: parseFloat(gridDecarbonizationRateEl.value) || 0,
    useTouRates: useTouRatesEl.checked,
    peakPrice: parseFloat(peakPriceEl.value) || 0,
    offPeakPrice: parseFloat(offPeakPriceEl.value) || 0,
    offPeakUsagePercent: parseInt(offPeakUsagePercentEl.value, 10) || 0,
    environmentalUnit: environmentalUnitEl.value as EnvironmentalUnit,
  };
}

/** Updates all UI input values based on the current state object. */
function updateUIFromState() {
    facilityTypeEl.value = state.facilityType;
    addressNumberEl.value = state.addressNumber;
    streetNameEl.value = state.streetName;
    townEl.value = state.town;
    annualHeatingThermsEl.value = state.annualHeatingTherms.toString();
    annualNonHeatingThermsEl.value = state.annualNonHeatingTherms.toString();
    annualExistingElecKwhEl.value = state.annualExistingElecKwh.toString();
    existingPeakLoadAmpsEl.value = state.existingPeakLoadAmps.toString();
    panelAmperageEl.value = state.panelAmps.toString();
    voltageEl.value = state.voltage.toString();
    breakerSpacesEl.value = state.breakerSpaces.toString();
    evChargerEl.value = state.evCharger;
    commercialEvStationsEl.value = state.commercialEVStations.toString();
    efficiencyTierEl.value = state.efficiencyTier;
    gasPriceEl.value = formatNumber(state.gasPricePerTherm, 2);
    electricityPriceEl.value = formatNumber(state.electricityPricePerKwh, 2);
    gasPriceEscalationEl.value = state.gasPriceEscalation.toString();
    electricityPriceEscalationEl.value = state.electricityPriceEscalation.toString();
    highRiskGasEscalationEl.value = state.highRiskGasEscalation.toString();
    gridDecarbonizationRateEl.value = state.gridDecarbonizationRate.toString();
    useTouRatesEl.checked = state.useTouRates;
    peakPriceEl.value = formatNumber(state.peakPrice, 2);
    offPeakPriceEl.value = formatNumber(state.offPeakPrice, 2);
    offPeakUsagePercentEl.value = state.offPeakUsagePercent.toString();
    environmentalUnitEl.value = state.environmentalUnit;

    touRatesGridEl.classList.toggle('hidden', !state.useTouRates);

    // Toggle EV input visibility based on loaded state
    const isResidential = state.facilityType === 'Residential';
    residentialEvChargerFieldEl.classList.toggle('hidden', !isResidential);
    commercialEvStationsFieldEl.classList.toggle('hidden', isResidential);
}


function updateStateAndRender() {
  updateStateFromUI();
  render();
}

function handleFacilityTypeChange() {
    const facilityType = facilityTypeEl.value as FacilityType;
    
    // Toggle visibility of EV charger inputs
    const isResidential = facilityType === 'Residential';
    residentialEvChargerFieldEl.classList.toggle('hidden', !isResidential);
    commercialEvStationsFieldEl.classList.toggle('hidden', isResidential);

    // Set default gas usage
    const usageDefaults = DEFAULT_USAGE_BY_FACILITY[facilityType];
    if (usageDefaults) {
        annualHeatingThermsEl.value = usageDefaults.heating.toString();
        annualNonHeatingThermsEl.value = usageDefaults.nonHeating.toString();
    }
    
    // Set default existing electrical usage and peak load
    const elecUsageDefault = DEFAULT_EXISTING_ELEC_USAGE_BY_FACILITY[facilityType];
    if (elecUsageDefault) {
        annualExistingElecKwhEl.value = elecUsageDefault.toString();
        // Estimate peak load from annual kWh. This is a rough heuristic.
        // Assumes a load factor of about 20-25% (8760 hours/year).
        // Peak kW ~= (Annual kWh / 8760) * 4.
        // Amps = (kW * 1000) / voltage
        const avgKw = elecUsageDefault / 8760;
        const peakKw = avgKw * 4; // Peaking factor
        const voltage = parseInt(voltageEl.value, 10) || 240;
        const peakAmps = (peakKw * 1000) / voltage;
        existingPeakLoadAmpsEl.value = Math.round(peakAmps).toString();
    }
    
    // Set default EV stations for non-residential
    if (!isResidential) {
        const stationDefaults = DEFAULT_EV_STATIONS_BY_FACILITY[facilityType] || 0;
        commercialEvStationsEl.value = stationDefaults.toString();
    }

    updateStateAndRender();
}

function handleTouToggle() {
    touRatesGridEl.classList.toggle('hidden', !useTouRatesEl.checked);
    updateStateAndRender();
}


function addApplianceByKey(key: string) {
    const definition = APPLIANCE_DEFINITIONS[key];
    if (!definition) return;

    const newAppliance: Appliance = {
        id: crypto.randomUUID(),
        key: key,
        btu: definition.defaultBtu,
        included: true,
        efficiency: definition.defaultEfficiency,
    };

    if (definition.category === 'Boiler' || definition.category === 'Space Heater') {
        newAppliance.zones = 3; // Default zones for a new ductless system
        newAppliance.supplementalHeaters = 2; // Default heaters for small rooms
    }

    state.appliances.push(newAppliance);
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
    
    const allTabs = [tabPlannerBtn, tabReportBtn, tabMethodologyBtn, tabInputTablesBtn, tabSaveLoadBtn];
    const allContent = [plannerContentEl, reportContentWrapperEl, methodologyContentEl, inputTablesContentEl, saveLoadContentEl];

    allTabs.forEach(tab => tab.classList.remove('active'));
    allContent.forEach(content => content.classList.remove('active'));
    
    clickedButton.classList.add('active');

    if (clickedButton === tabPlannerBtn) plannerContentEl.classList.add('active');
    else if (clickedButton === tabReportBtn) reportContentWrapperEl.classList.add('active');
    else if (clickedButton === tabMethodologyBtn) methodologyContentEl.classList.add('active');
    else if (clickedButton === tabInputTablesBtn) inputTablesContentEl.classList.add('active');
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

function handleSaveDefaults() {
    const defaultsJson = JSON.stringify(defaultInputs, null, 2);
    const blob = new Blob([defaultsJson], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'electrification-defaults.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function handleLoadDefaults(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) {
        loadedDefaultsFilenameEl.textContent = 'No file selected.';
        return;
    }

    const file = input.files[0];
    loadedDefaultsFilenameEl.textContent = file.name;
    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const text = e.target?.result;
            if (typeof text !== 'string') throw new Error('File could not be read as text.');
            
            const loadedDefaults = JSON.parse(text);

            if (!loadedDefaults.EQUIPMENT_COSTS || !loadedDefaults.FEDERAL_REBATES) {
                throw new Error('Invalid defaults file format.');
            }

            // A simple merge to avoid breaking the app if the loaded file is missing new keys
            for (const key in defaultInputs) {
                if (loadedDefaults[key]) {
                    (defaultInputs as any)[key] = loadedDefaults[key];
                }
            }
            
            // Update the main state with the new utility price defaults
            state.gasPricePerTherm = defaultInputs.REGIONAL_DATA.MA.gasPrice;
            state.electricityPricePerKwh = defaultInputs.REGIONAL_DATA.MA.elecPrice;

            renderInputTables();
            updateUIFromState();
            render();
            alert('Default values loaded successfully!');

        } catch (error) {
            console.error('Failed to load or parse defaults file:', error);
            alert(`Error: Could not load defaults file. Please ensure it is a valid file.\n\n(${(error as Error).message})`);
            loadedDefaultsFilenameEl.textContent = 'Load failed. Please try again.';
        }
    };
    reader.readAsText(file);
}

function updateDefaultsFromUI(event: Event) {
    const input = event.target as HTMLInputElement;
    const value = parseFloat(input.value);
    if (isNaN(value)) return;

    const { category, tier, range, key, source } = input.dataset;

    const parentTbody = input.closest('tbody');
    if (!parentTbody) return;

    const updateCostArray = (arr: number[], range: string, value: number) => {
        if (range === 'low') arr[0] = value;
        else if (range === 'medium') arr[1] = value;
        else if (range === 'high') arr[2] = value;
    };

    switch (parentTbody.id) {
        case 'defaults-utility-tbody':
            if (key) (state as any)[key] = value;
            break;
        case 'defaults-equipment-tbody':
            if (category && tier && range) {
                const costArray = defaultInputs.EQUIPMENT_COSTS[category as ApplianceCategory][tier as EfficiencyTier];
                updateCostArray(costArray, range, value);
            }
            break;
        case 'defaults-installation-tbody':
             if (category && range) {
                const costArray = defaultInputs.INSTALLATION_COSTS[category as ApplianceCategory];
                updateCostArray(costArray, range, value);
            }
            break;
        case 'defaults-distribution-tbody':
            if (key && range) {
                const costArray = (defaultInputs.DISTRIBUTION_SYSTEM_COSTS[key] || defaultInputs.DUCTLESS_MINI_SPLIT_COSTS[key]);
                updateCostArray(costArray, range, value);
            }
            break;
        case 'defaults-decommissioning-tbody':
             if (category && range) {
                const costArray = defaultInputs.DECOMMISSIONING_COSTS[category as ApplianceCategory];
                updateCostArray(costArray, range, value);
            }
            break;
        case 'defaults-electrical-tbody':
             if (key && range) {
                const costArray = defaultInputs.ELECTRICAL_UPGRADE_COSTS[key];
                updateCostArray(costArray, range, value);
            }
            break;
        case 'defaults-incentives-tbody':
            if (key && source) {
                if (source === 'massSave') defaultInputs.REGIONAL_DATA.MA.massSaveRebates[key] = value;
                if (source === 'federal') defaultInputs.FEDERAL_REBATES[key] = value;
            }
            break;
    }
    
    // After updating the defaults, update the main planner UI and re-render the report
    updateUIFromState();
    render();
}


// --- INITIALIZATION ---
const allInputs = document.querySelectorAll('#inputs input, #inputs select');
allInputs.forEach(input => {
    input.addEventListener('input', updateStateAndRender);
    input.addEventListener('change', updateStateAndRender);
});
facilityTypeEl.addEventListener('change', handleFacilityTypeChange);
useTouRatesEl.addEventListener('change', handleTouToggle);


// Modal listeners
addApplianceBtn.addEventListener('click', openAddApplianceModal);
modalCloseBtn.addEventListener('click', closeAddApplianceModal);
modalOverlayEl.addEventListener('click', closeAddApplianceModal);
modalApplianceGridEl.addEventListener('click', handleModalGridClick);

// Tab Listeners
tabPlannerBtn.addEventListener('click', handleTabClick);
tabReportBtn.addEventListener('click', handleTabClick);
tabMethodologyBtn.addEventListener('click', handleTabClick);
tabInputTablesBtn.addEventListener('click', handleTabClick);
tabSaveLoadBtn.addEventListener('click', handleTabClick);

// Save/Load Listeners
exportPdfBtn.addEventListener('click', handleExportPdf);
saveProjectBtn.addEventListener('click', handleSaveProject);
loadProjectInput.addEventListener('change', handleLoadProject);
saveDefaultsBtn.addEventListener('click', handleSaveDefaults);
loadDefaultsInput.addEventListener('change', handleLoadDefaults);

// Input Tables Listeners (using event delegation)
inputTablesContentEl.addEventListener('change', (event) => {
    if ((event.target as HTMLElement).tagName === 'INPUT') {
        updateDefaultsFromUI(event);
    }
});

// Remove Appliance Listener (using event delegation)
appliancesListEl.addEventListener('click', handleRemoveAppliance);


// Initial render on page load
document.addEventListener('DOMContentLoaded', () => {
    updateUIFromState();
    renderInputTables();
    render(); // Initial render with default state
});