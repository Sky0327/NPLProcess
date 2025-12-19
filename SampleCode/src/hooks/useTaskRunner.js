import { useCallback } from 'react'
import useWorkflowStore from '../store/workflowStore'
import { apiService } from '../services/apiService'
import { getTaskById } from '../data/workflowConfig'

export const useTaskRunner = () => {
  const {
    taskResults,
    taskStatus,
    setTaskResult,
    setTaskStatus,
    addLogEntry,
    canExecuteTask,
    projectConfig,
  } = useWorkflowStore()

  // Run a single task
  const runTask = useCallback(
    async (taskId) => {
      const task = getTaskById(taskId)
      if (!task) {
        throw new Error(`Task not found: ${taskId}`)
      }

      // Check dependencies
      if (!canExecuteTask(taskId)) {
        const deps = task.deps || []
        addLogEntry({
          type: 'warning',
          action: `${task.name} 실행 불가`,
          details: `필요한 데이터: ${deps.join(', ')}`,
          phase: task.phaseId,
        })
        return { success: false, error: 'Dependencies not met' }
      }

      // Set loading state
      setTaskStatus(taskId, { loading: true, error: null })

      addLogEntry({
        type: 'info',
        action: `${task.name} 시작`,
        phase: task.phaseId,
      })

      try {
        let result

        // Execute based on task type
        if (task.api && apiService[task.api]) {
          // Gather dependency data if needed
          const depData = (task.deps || []).reduce((acc, dep) => {
            acc[dep] = taskResults[dep]
            return acc
          }, {})

          result = await apiService[task.api](
            task.phaseId === 2 ? projectConfig.inputData || [] : depData
          )
        } else {
          // Simulate task execution for non-API tasks
          await new Promise((resolve) => setTimeout(resolve, 1500))
          result = {
            success: true,
            data: {
              completedAt: new Date().toISOString(),
              status: 'completed',
            },
          }
        }

        if (result.success) {
          setTaskResult(taskId, result.data)
          setTaskStatus(taskId, { loading: false, error: null })

          const count = Array.isArray(result.data) ? result.data.length : 1
          addLogEntry({
            type: 'success',
            action: `${task.name} 완료`,
            details: Array.isArray(result.data) ? `${count}건 처리` : undefined,
            phase: task.phaseId,
          })

          return { success: true, data: result.data }
        } else {
          throw new Error(result.error || '처리 실패')
        }
      } catch (error) {
        setTaskStatus(taskId, { loading: false, error: error.message })

        addLogEntry({
          type: 'error',
          action: `${task.name} 오류`,
          details: error.message,
          phase: task.phaseId,
        })

        return { success: false, error: error.message }
      }
    },
    [
      taskResults,
      setTaskResult,
      setTaskStatus,
      addLogEntry,
      canExecuteTask,
      projectConfig,
    ]
  )

  // Run multiple tasks in parallel
  const runTasksInParallel = useCallback(
    async (taskIds) => {
      addLogEntry({
        type: 'info',
        action: '병렬 작업 시작',
        details: `${taskIds.length}개 작업 실행`,
      })

      const results = await Promise.allSettled(
        taskIds.map((taskId) => runTask(taskId))
      )

      const successCount = results.filter(
        (r) => r.status === 'fulfilled' && r.value?.success
      ).length

      addLogEntry({
        type: successCount === taskIds.length ? 'success' : 'warning',
        action: '병렬 작업 완료',
        details: `${successCount}/${taskIds.length}개 성공`,
      })

      return results
    },
    [runTask, addLogEntry]
  )

  // Run multiple tasks sequentially
  const runTasksSequentially = useCallback(
    async (taskIds) => {
      addLogEntry({
        type: 'info',
        action: '순차 작업 시작',
        details: `${taskIds.length}개 작업 실행`,
      })

      const results = []
      for (const taskId of taskIds) {
        const result = await runTask(taskId)
        results.push({ taskId, ...result })

        // Stop if task failed
        if (!result.success) {
          addLogEntry({
            type: 'warning',
            action: '순차 작업 중단',
            details: `${taskId} 실패로 중단됨`,
          })
          break
        }
      }

      const successCount = results.filter((r) => r.success).length
      if (successCount === taskIds.length) {
        addLogEntry({
          type: 'success',
          action: '순차 작업 완료',
          details: `${successCount}개 모두 성공`,
        })
      }

      return results
    },
    [runTask, addLogEntry]
  )

  // Check if any task is currently loading
  const isAnyTaskLoading = useCallback(
    (taskIds) => {
      return taskIds.some((id) => taskStatus[id]?.loading)
    },
    [taskStatus]
  )

  // Get task execution status
  const getTaskExecutionStatus = useCallback(
    (taskId) => {
      return {
        isLoading: taskStatus[taskId]?.loading || false,
        hasError: !!taskStatus[taskId]?.error,
        error: taskStatus[taskId]?.error,
        hasResult: taskResults[taskId] !== null && taskResults[taskId] !== undefined,
        result: taskResults[taskId],
        canExecute: canExecuteTask(taskId),
      }
    },
    [taskStatus, taskResults, canExecuteTask]
  )

  return {
    runTask,
    runTasksInParallel,
    runTasksSequentially,
    isAnyTaskLoading,
    getTaskExecutionStatus,
  }
}

export default useTaskRunner
