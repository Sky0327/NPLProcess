import { create } from 'zustand'
import { persist } from 'zustand/middleware'

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

      // Calculate phase progress
      calculatePhaseProgress: (phaseId) => {
        const state = get()
        const phaseTasks = {
          1: ['projectConfig'],
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
          const config = state.projectConfig
          const filled = [
            config.reportName,
            config.projectId,
            config.apiId,
            config.inputFolderPath,
          ].filter(Boolean).length
          return Math.round((filled / 4) * 100)
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
        taskResults: state.taskResults,
        activityLog: state.activityLog,
        activePhase: state.activePhase,
      }),
    }
  )
)

export default useWorkflowStore
