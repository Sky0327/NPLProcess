import { useCallback, useMemo } from 'react'
import useWorkflowStore from '../store/workflowStore'

export const useActivityLog = () => {
  const { activityLog, addLogEntry, clearLog } = useWorkflowStore()

  // Log helper functions
  const logInfo = useCallback(
    (action, details, phase) => {
      addLogEntry({ type: 'info', action, details, phase })
    },
    [addLogEntry]
  )

  const logSuccess = useCallback(
    (action, details, phase) => {
      addLogEntry({ type: 'success', action, details, phase })
    },
    [addLogEntry]
  )

  const logError = useCallback(
    (action, details, phase) => {
      addLogEntry({ type: 'error', action, details, phase })
    },
    [addLogEntry]
  )

  const logWarning = useCallback(
    (action, details, phase) => {
      addLogEntry({ type: 'warning', action, details, phase })
    },
    [addLogEntry]
  )

  // Filter logs by type
  const getLogsByType = useCallback(
    (type) => {
      return activityLog.filter((log) => log.type === type)
    },
    [activityLog]
  )

  // Filter logs by phase
  const getLogsByPhase = useCallback(
    (phase) => {
      return activityLog.filter((log) => log.phase === phase)
    },
    [activityLog]
  )

  // Get recent logs
  const getRecentLogs = useCallback(
    (count = 10) => {
      return activityLog.slice(0, count)
    },
    [activityLog]
  )

  // Get log statistics
  const stats = useMemo(() => {
    return {
      total: activityLog.length,
      success: activityLog.filter((l) => l.type === 'success').length,
      error: activityLog.filter((l) => l.type === 'error').length,
      warning: activityLog.filter((l) => l.type === 'warning').length,
      info: activityLog.filter((l) => l.type === 'info').length,
    }
  }, [activityLog])

  // Export logs as JSON
  const exportLogs = useCallback(() => {
    const data = JSON.stringify(activityLog, null, 2)
    const blob = new Blob([data], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `activity-log-${new Date().toISOString().split('T')[0]}.json`
    a.click()
    URL.revokeObjectURL(url)
  }, [activityLog])

  // Export logs as CSV
  const exportLogsAsCsv = useCallback(() => {
    const headers = ['Timestamp', 'Type', 'Action', 'Details', 'Phase']
    const rows = activityLog.map((log) => [
      log.timestamp,
      log.type,
      log.action,
      log.details || '',
      log.phase || '',
    ])

    const csvContent = [
      headers.join(','),
      ...rows.map((row) =>
        row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(',')
      ),
    ].join('\n')

    const blob = new Blob([csvContent], { type: 'text/csv' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `activity-log-${new Date().toISOString().split('T')[0]}.csv`
    a.click()
    URL.revokeObjectURL(url)
  }, [activityLog])

  return {
    // Log data
    logs: activityLog,
    stats,

    // Log actions
    logInfo,
    logSuccess,
    logError,
    logWarning,
    clearLog,

    // Filtering
    getLogsByType,
    getLogsByPhase,
    getRecentLogs,

    // Export
    exportLogs,
    exportLogsAsCsv,
  }
}

export default useActivityLog
