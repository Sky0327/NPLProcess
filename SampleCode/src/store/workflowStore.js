import { create } from 'zustand'
import { persist } from 'zustand/middleware'

// Helper to generate unique IDs
const generateId = () => `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`

// Initial Cover Data
const initialCoverData = {
  // 기본 정보
  sellingInstitution: '',     // 매각기관
  businessType: '사후재정산방식 담보부채권 양수도', // 업무구분
  borrowerName: '',           // 차주명
  propertyType: '',           // 물건유형
  address: '',                // 소재지
  reportDate: null,           // 작성일
  preparingOrg: 'MCI대부㈜',  // 작성기관

  // 주요 가정
  discountRate: 0.0628,       // 현가할인율 (6.28%)
  baseRate: 0.056,            // 기본할인율 (5.6%)
  purchaseRate: 0.0088,       // 매입할인율 (0.88%)
  managementCostRate: 0.0095, // 관리비용률 (0.95%)
  discountPeriod: 360,        // 할인기간 (360/420)

  // 사업자 구분
  businessClassification: '개인', // 개인/기업
}

// Create empty debt item
const createEmptyDebt = (priority = '') => ({
  id: generateId(),
  priority,                   // 구분(순위)
  institutionName: '',        // 금융기관명
  accountNumber: '',          // 계좌번호
  maxDebtAmount: 0,           // 채권최고액
  outstandingBalance: 0,      // 대출잔액
  advancePayment: 0,          // 가지급금
  accruedInterest: 0,         // 미수이자
  handlingDate: null,         // 취급일
  maturityDate: null,         // 만기일
  defaultDate: null,          // 기한이익상실일
  mortgageDate: null,         // 근저당설정일
  delinquencyRate: 0,         // 연체이자율
})

// Create empty collateral item
const createEmptyCollateral = (propertyIndex = 1) => ({
  id: generateId(),
  propertyIndex,              // 물건 순번

  // 토지 정보
  landUseZone: '',            // 용도지역
  landArea: 0,                // 면적 (㎡)
  landPrice: 0,               // 개별공시지가 (원/㎡)

  // 건물 정보
  propertyType: '',           // 물건유형
  address: '',                // 소재지
  buildingArea: 0,            // 면적 (㎡)
  buildingScale: '',          // 규모
  buildingStructure: '',      // 구조

  // 감정평가
  appraisedValue: 0,          // 대출당시 감정평가액
  internalAppraisal: 0,       // 자체감정가
  appraiser: '',              // 평가기관
  auctionRate: 0,             // 낙찰가율 (%)
})

// Initial calculated results
const initialCalculatedResults = {
  // Per-property calculations
  propertyResults: [],
  // Aggregated totals
  totalCollateralValue: 0,    // 총 담보평가액
  totalNRV: 0,                // 총 순회수금
  totalNPV: 0,                // 총 매각대금
  totalFV: 0,                 // 총 미래가
  totalSettlement: 0,         // 총 정산차액
  totalFees: 0,               // 총 수수료
  // Debt totals
  totalMaxDebtAmount: 0,      // 총 채권최고액
  totalOutstandingBalance: 0, // 총 대출잔액
  totalRecoveryDebt: 0,       // 총 회수시점채권액
}

const initialTaskResults = {
  // Phase 2 - Parallel Queries
  registry: null,
  landPrice: null,
  auction: null,
  infocareStats: null,
  infocareIntegrated: null,
  realEstatePrice: null,
  valuemapPrice: null,
  // Phase 3 - Intermediate Processing
  kbPrice: null,
  infocareCases: null,
  distanceGookto: null,
  distanceValuemap: null,
  // Phase 4 - Reports
  'report-담보물정보': null,
  'report-감정평가': null,
  'report-경매정보': null,
  'report-낙찰통계': null,
  'report-낙찰사례': null,
  'report-실거래_국토': null,
  'report-실거래_밸류맵': null,
  // Phase 5 - Final
  'report-물건지': null,
  'report-채권현황': null,
  pdfConversion: null,
  xlsxExport: null,
}

const initialPhases = {
  1: { status: 'pending', progress: 0 },
  2: { status: 'pending', progress: 0 },
  3: { status: 'pending', progress: 0 },
  4: { status: 'pending', progress: 0 },
  5: { status: 'pending', progress: 0 },
}

const useWorkflowStore = create(
  persist(
    (set, get) => ({
      // Current active phase
      activePhase: 1,

      // Project Configuration
      projectConfig: {
        reportName: '',
        projectId: '',
        apiId: '',
        apiPassword: '',
        inputFolderPath: '',
      },

      // Cover Sheet Data
      coverData: { ...initialCoverData },

      // Debt Information Array
      debtInfo: [createEmptyDebt('1순위')],

      // Collateral Information Array
      collateralInfo: [createEmptyCollateral(1)],

      // Calculated Results
      calculatedResults: { ...initialCalculatedResults },

      // Phase Status
      phases: { ...initialPhases },

      // Task Results by ID
      taskResults: { ...initialTaskResults },

      // Task Status (loading, error states)
      taskStatus: {},

      // Activity Log
      activityLog: [],

      // Actions
      setActivePhase: (phaseId) => set({ activePhase: phaseId }),

      setProjectConfig: (config) =>
        set((state) => ({
          projectConfig: { ...state.projectConfig, ...config },
        })),

      // Cover Data Actions
      setCoverData: (data) =>
        set((state) => ({
          coverData: { ...state.coverData, ...data },
        })),

      resetCoverData: () => set({ coverData: { ...initialCoverData } }),

      // Debt Info CRUD Actions
      addDebtInfo: (priority = '') =>
        set((state) => ({
          debtInfo: [...state.debtInfo, createEmptyDebt(priority || `${state.debtInfo.length + 1}순위`)],
        })),

      updateDebtInfo: (id, updates) =>
        set((state) => ({
          debtInfo: state.debtInfo.map((debt) =>
            debt.id === id ? { ...debt, ...updates } : debt
          ),
        })),

      removeDebtInfo: (id) =>
        set((state) => ({
          debtInfo: state.debtInfo.filter((debt) => debt.id !== id),
        })),

      setDebtInfo: (debtArray) => set({ debtInfo: debtArray }),

      // Collateral Info CRUD Actions
      addCollateralInfo: () =>
        set((state) => ({
          collateralInfo: [
            ...state.collateralInfo,
            createEmptyCollateral(state.collateralInfo.length + 1),
          ],
        })),

      updateCollateralInfo: (id, updates) =>
        set((state) => ({
          collateralInfo: state.collateralInfo.map((col) =>
            col.id === id ? { ...col, ...updates } : col
          ),
        })),

      removeCollateralInfo: (id) =>
        set((state) => ({
          collateralInfo: state.collateralInfo.filter((col) => col.id !== id),
        })),

      setCollateralInfo: (collateralArray) => set({ collateralInfo: collateralArray }),

      // Calculated Results Actions
      setCalculatedResults: (results) =>
        set((state) => ({
          calculatedResults: { ...state.calculatedResults, ...results },
        })),

      resetCalculatedResults: () => set({ calculatedResults: { ...initialCalculatedResults } }),

      updatePhase: (phaseId, updates) =>
        set((state) => ({
          phases: {
            ...state.phases,
            [phaseId]: { ...state.phases[phaseId], ...updates },
          },
        })),

      setTaskResult: (taskId, result) =>
        set((state) => ({
          taskResults: { ...state.taskResults, [taskId]: result },
        })),

      setTaskStatus: (taskId, status) =>
        set((state) => ({
          taskStatus: { ...state.taskStatus, [taskId]: status },
        })),

      addLogEntry: (entry) =>
        set((state) => ({
          activityLog: [
            {
              id: Date.now().toString(),
              timestamp: new Date().toISOString(),
              ...entry,
            },
            ...state.activityLog,
          ].slice(0, 100), // Keep last 100 entries
        })),

      clearLog: () => set({ activityLog: [] }),

      // Check if Phase 1 task is complete
      isPhase1TaskComplete: (taskId) => {
        const state = get()
        const config = state.projectConfig
        const cover = state.coverData

        if (taskId === 'coverData') {
          return !!(config.reportName && config.projectId && cover.sellingInstitution && cover.borrowerName)
        }
        return false
      },

      // Calculate phase progress
      calculatePhaseProgress: (phaseId) => {
        const state = get()
        const phaseTasks = {
          1: ['coverData'],
          2: [
            'registry',
            'landPrice',
            'auction',
            'infocareStats',
            'infocareIntegrated',
            'realEstatePrice',
            'valuemapPrice',
          ],
          3: ['kbPrice', 'infocareCases', 'distanceGookto', 'distanceValuemap'],
          4: [
            'report-담보물정보',
            'report-감정평가',
            'report-경매정보',
            'report-낙찰통계',
            'report-낙찰사례',
            'report-실거래_국토',
            'report-실거래_밸류맵',
          ],
          5: [
            'report-물건지',
            'report-채권현황',
            'pdfConversion',
            'xlsxExport',
          ],
        }

        const tasks = phaseTasks[phaseId] || []
        if (tasks.length === 0) return 0

        if (phaseId === 1) {
          const completedTasks = tasks.filter(taskId => state.isPhase1TaskComplete(taskId)).length
          return Math.round((completedTasks / tasks.length) * 100)
        }

        const completedTasks = tasks.filter(
          (taskId) => state.taskResults[taskId] !== null
        ).length
        return Math.round((completedTasks / tasks.length) * 100)
      },

      // Calculate overall progress
      calculateOverallProgress: () => {
        const state = get()
        const weights = { 1: 10, 2: 25, 3: 20, 4: 30, 5: 15 }
        let totalProgress = 0

        for (let i = 1; i <= 5; i++) {
          const phaseProgress = state.calculatePhaseProgress(i)
          totalProgress += (phaseProgress / 100) * weights[i]
        }

        return Math.round(totalProgress)
      },

      // Check if task dependencies are met
      canExecuteTask: (taskId) => {
        const state = get()
        const dependencies = {
          kbPrice: ['registry', 'landPrice'],
          infocareCases: ['infocareIntegrated'],
          distanceGookto: ['realEstatePrice'],
          distanceValuemap: ['valuemapPrice'],
          'report-담보물정보': ['registry', 'kbPrice'],
          'report-감정평가': ['kbPrice'],
          'report-경매정보': ['auction'],
          'report-낙찰통계': ['infocareStats'],
          'report-낙찰사례': ['infocareCases'],
          'report-실거래_국토': ['distanceGookto'],
          'report-실거래_밸류맵': ['distanceValuemap'],
        }

        const deps = dependencies[taskId]
        if (!deps) return true

        return deps.every((depId) => state.taskResults[depId] !== null)
      },

      // Reset workflow
      resetWorkflow: () =>
        set({
          activePhase: 1,
          projectConfig: {
            reportName: '',
            projectId: '',
            apiId: '',
            apiPassword: '',
            inputFolderPath: '',
          },
          coverData: { ...initialCoverData },
          debtInfo: [createEmptyDebt('1순위')],
          collateralInfo: [createEmptyCollateral(1)],
          calculatedResults: { ...initialCalculatedResults },
          phases: { ...initialPhases },
          taskResults: { ...initialTaskResults },
          taskStatus: {},
          activityLog: [],
        }),
    }),
    {
      name: 'npl-workflow-storage',
      partialize: (state) => ({
        projectConfig: state.projectConfig,
        coverData: state.coverData,
        debtInfo: state.debtInfo,
        collateralInfo: state.collateralInfo,
        calculatedResults: state.calculatedResults,
        taskResults: state.taskResults,
        activityLog: state.activityLog,
        activePhase: state.activePhase,
      }),
    }
  )
)

export default useWorkflowStore
