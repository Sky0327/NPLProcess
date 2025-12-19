import { useCallback } from 'react'
import useWorkflowStore from '../store/workflowStore'
import { PHASES, getTaskById } from '../data/workflowConfig'

export const useWorkflow = () => {
  const store = useWorkflowStore()

  const {
    activePhase,
    setActivePhase,
    taskResults,
    taskStatus,
    calculatePhaseProgress,
    calculateOverallProgress,
    canExecuteTask,
  } = store

  // Get current phase info
  const currentPhase = PHASES[activePhase]

  // Get phase completion status
  const getPhaseStatus = useCallback(
    (phaseId) => {
      const progress = calculatePhaseProgress(phaseId)
      if (progress === 100) return 'completed'
      if (progress > 0) return 'in_progress'
      return 'pending'
    },
    [calculatePhaseProgress]
  )

  // Check if can proceed to next phase
  const canProceedToNext = useCallback(() => {
    const progress = calculatePhaseProgress(activePhase)
    return progress === 100
  }, [activePhase, calculatePhaseProgress])

  // Navigate to next phase
  const goToNextPhase = useCallback(() => {
    if (activePhase < 5 && canProceedToNext()) {
      setActivePhase(activePhase + 1)
      return true
    }
    return false
  }, [activePhase, canProceedToNext, setActivePhase])

  // Navigate to previous phase
  const goToPreviousPhase = useCallback(() => {
    if (activePhase > 1) {
      setActivePhase(activePhase - 1)
      return true
    }
    return false
  }, [activePhase, setActivePhase])

  // Get task status with result
  const getTaskStatus = useCallback(
    (taskId) => {
      const result = taskResults[taskId]
      const status = taskStatus[taskId]

      return {
        hasResult: result !== null && result !== undefined,
        isLoading: status?.loading || false,
        hasError: !!status?.error,
        error: status?.error,
        result,
        canExecute: canExecuteTask(taskId),
      }
    },
    [taskResults, taskStatus, canExecuteTask]
  )

  // Get all tasks for current phase
  const getCurrentPhaseTasks = useCallback(() => {
    const phase = PHASES[activePhase]
    if (!phase) return []

    return phase.tasks.map((task) => ({
      ...task,
      ...getTaskStatus(task.id),
    }))
  }, [activePhase, getTaskStatus])

  // Get phase summary
  const getPhaseSummary = useCallback(
    (phaseId) => {
      const phase = PHASES[phaseId]
      if (!phase) return null

      const tasks = phase.tasks
      const completedTasks = tasks.filter(
        (t) => taskResults[t.id] !== null && taskResults[t.id] !== undefined
      )
      const loadingTasks = tasks.filter((t) => taskStatus[t.id]?.loading)
      const errorTasks = tasks.filter((t) => taskStatus[t.id]?.error)

      return {
        ...phase,
        progress: calculatePhaseProgress(phaseId),
        status: getPhaseStatus(phaseId),
        totalTasks: tasks.length,
        completedCount: completedTasks.length,
        loadingCount: loadingTasks.length,
        errorCount: errorTasks.length,
      }
    },
    [taskResults, taskStatus, calculatePhaseProgress, getPhaseStatus]
  )

  return {
    // Current state
    activePhase,
    currentPhase,
    overallProgress: calculateOverallProgress(),

    // Navigation
    setActivePhase,
    goToNextPhase,
    goToPreviousPhase,
    canProceedToNext,

    // Phase info
    getPhaseStatus,
    getPhaseSummary,

    // Task info
    getTaskStatus,
    getCurrentPhaseTasks,
    canExecuteTask,
  }
}

export default useWorkflow
