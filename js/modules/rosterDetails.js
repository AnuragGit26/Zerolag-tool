// All roster-related details: sheet names and column mappings

// Sheet names
export const SHEET_SERVICE_CLOUD = 'Service Cloud';
export const SHEET_SALES_CLOUD = 'Sales Cloud';
export const SHEET_INDUSTRY_CLOUD = 'Industry Cloud';
export const SHEET_DATA_CLOUD_AF = 'Data Cloud and AF';

// Signature column mappings
export const COLUMN_MAP = {
    CIC: { APAC: 'B', EMEA: 'J', AMER: 'S' },
    TE: { APAC: 'F', EMEA: 'O', AMER: 'W' },
    SWARM_LEAD: { APAC: 'G', EMEA: 'P', AMER: 'X' }
};

export function getCICColumnForShift(shift) {
    return COLUMN_MAP.CIC[shift] || COLUMN_MAP.CIC.AMER;
}

export function getTEColumnForShift(shift) {
    return COLUMN_MAP.TE[shift] || COLUMN_MAP.TE.AMER;
}

export function getSwarmLeadColumnForShift(shift) {
    return COLUMN_MAP.SWARM_LEAD[shift] || COLUMN_MAP.SWARM_LEAD.AMER;
}

// Premier-specific mappings
export const PREMIER_COLUMN_MAP = {
    SALES_SERVICE: {
        DEV_TE: { APAC: 'D', EMEA: 'M', AMER: 'U' },
        NON_DEV_TE: { APAC: 'C', EMEA: 'K', AMER: 'T' },
        SWARM_LEAD: { APAC: 'E', EMEA: 'N', AMER: 'V' }
    },
    INDUSTRY: {
        TE: { APAC: 'B', EMEA: 'F', AMER: 'H' },
        SWARM_LEAD: { APAC: 'C', EMEA: 'J', AMER: 'I' }
    },
    DATA: {
        DATACLOUD: { APAC: 'B', EMEA: 'C', AMER: 'D' },
        AGENTFORCE: { APAC: 'F', EMEA: 'G', AMER: 'H' }
    }
};

// Premier helpers
export function getPremierSalesDevTEColumn(shift) {
    return PREMIER_COLUMN_MAP.SALES_SERVICE.DEV_TE[shift] || PREMIER_COLUMN_MAP.SALES_SERVICE.DEV_TE.AMER;
}

export function getPremierSalesNonDevTEColumn(shift) {
    return PREMIER_COLUMN_MAP.SALES_SERVICE.NON_DEV_TE[shift] || PREMIER_COLUMN_MAP.SALES_SERVICE.NON_DEV_TE.AMER;
}

export function getPremierSalesSwarmLeadColumn(shift) {
    return PREMIER_COLUMN_MAP.SALES_SERVICE.SWARM_LEAD[shift] || PREMIER_COLUMN_MAP.SALES_SERVICE.SWARM_LEAD.AMER;
}

export function getPremierIndustryTEColumn(shift) {
    return PREMIER_COLUMN_MAP.INDUSTRY.TE[shift] || PREMIER_COLUMN_MAP.INDUSTRY.TE.AMER;
}

export function getPremierIndustrySwarmLeadColumn(shift) {
    return PREMIER_COLUMN_MAP.INDUSTRY.SWARM_LEAD[shift] || PREMIER_COLUMN_MAP.INDUSTRY.SWARM_LEAD.AMER;
}

export function getPremierDataCloudColumn(shift) {
    return PREMIER_COLUMN_MAP.DATA.DATACLOUD[shift] || PREMIER_COLUMN_MAP.DATA.DATACLOUD.AMER;
}

export function getPremierAgentforceColumn(shift) {
    return PREMIER_COLUMN_MAP.DATA.AGENTFORCE[shift] || PREMIER_COLUMN_MAP.DATA.AGENTFORCE.AMER;
}
