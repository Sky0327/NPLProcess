import React from 'react'
import {
  Box,
  Typography,
  IconButton,
  Chip,
  Tooltip,
  Button,
} from '@mui/material'
import {
  CheckCircle,
  Error as ErrorIcon,
  Info,
  Warning,
  Delete,
  Download,
} from '@mui/icons-material'
import useWorkflowStore from '../../store/workflowStore'

const LogIcon = ({ type }) => {
  switch (type) {
    case 'success':
      return <CheckCircle sx={{ color: 'success.main', fontSize: 16 }} />
    case 'error':
      return <ErrorIcon sx={{ color: 'error.main', fontSize: 16 }} />
    case 'warning':
      return <Warning sx={{ color: 'warning.main', fontSize: 16 }} />
    default:
      return <Info sx={{ color: 'info.main', fontSize: 16 }} />
  }
}

const formatTime = (timestamp) => {
  const date = new Date(timestamp)
  return date.toLocaleTimeString('ko-KR', {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  })
}

const ActivityLog = ({ open, width }) => {
  const { activityLog, clearLog } = useWorkflowStore()

  const handleExport = () => {
    const data = JSON.stringify(activityLog, null, 2)
    const blob = new Blob([data], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `activity-log-${new Date().toISOString().split('T')[0]}.json`
    a.click()
    URL.revokeObjectURL(url)
  }

  if (!open) return null

  return (
    <Box
      sx={{
        width: width,
        height: '100vh',
        display: 'flex',
        flexDirection: 'column',
        position: 'fixed',
        right: 0,
        top: 0,
        bgcolor: 'background.paper',
      }}
    >
      {/* Header */}
      <Box
        sx={{
          p: 2,
          borderBottom: '1px solid',
          borderColor: 'divider',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
        }}
      >
        <Box>
          <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
            작업 이력
          </Typography>
          <Typography variant="caption" sx={{ color: 'text.secondary' }}>
            {activityLog.length}개의 로그
          </Typography>
        </Box>
        <Box sx={{ display: 'flex', gap: 0.5 }}>
          <Tooltip title="내보내기">
            <IconButton size="small" onClick={handleExport}>
              <Download fontSize="small" />
            </IconButton>
          </Tooltip>
          <Tooltip title="전체 삭제">
            <IconButton size="small" onClick={clearLog}>
              <Delete fontSize="small" />
            </IconButton>
          </Tooltip>
        </Box>
      </Box>

      {/* Log List */}
      <Box
        sx={{
          flexGrow: 1,
          overflow: 'auto',
          '&::-webkit-scrollbar': {
            width: 6,
          },
          '&::-webkit-scrollbar-thumb': {
            bgcolor: 'grey.300',
            borderRadius: 3,
          },
        }}
      >
        {activityLog.length === 0 ? (
          <Box
            sx={{
              p: 4,
              textAlign: 'center',
              color: 'text.secondary',
            }}
          >
            <Info sx={{ fontSize: 40, mb: 1, opacity: 0.3 }} />
            <Typography variant="body2">아직 작업 이력이 없습니다</Typography>
          </Box>
        ) : (
          <Box sx={{ p: 1 }}>
            {activityLog.map((entry) => (
              <Box
                key={entry.id}
                sx={{
                  p: 1.5,
                  mb: 1,
                  borderRadius: 1,
                  bgcolor:
                    entry.type === 'error'
                      ? 'error.50'
                      : entry.type === 'success'
                      ? 'success.50'
                      : 'grey.50',
                  border: '1px solid',
                  borderColor:
                    entry.type === 'error'
                      ? 'error.100'
                      : entry.type === 'success'
                      ? 'success.100'
                      : 'grey.200',
                  '&:hover': {
                    bgcolor:
                      entry.type === 'error'
                        ? 'error.100'
                        : entry.type === 'success'
                        ? 'success.100'
                        : 'grey.100',
                  },
                }}
              >
                <Box
                  sx={{
                    display: 'flex',
                    alignItems: 'flex-start',
                    gap: 1,
                  }}
                >
                  <LogIcon type={entry.type} />
                  <Box sx={{ flexGrow: 1, minWidth: 0 }}>
                    <Box
                      sx={{
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        mb: 0.5,
                      }}
                    >
                      <Typography
                        variant="caption"
                        sx={{
                          fontWeight: 600,
                          color: 'text.primary',
                        }}
                      >
                        {entry.action}
                      </Typography>
                      <Typography
                        variant="caption"
                        sx={{
                          color: 'text.secondary',
                          fontSize: '0.65rem',
                        }}
                      >
                        {formatTime(entry.timestamp)}
                      </Typography>
                    </Box>
                    {entry.details && (
                      <Typography
                        variant="caption"
                        sx={{
                          color: 'text.secondary',
                          display: 'block',
                          fontSize: '0.7rem',
                        }}
                      >
                        {entry.details}
                      </Typography>
                    )}
                    {entry.phase && (
                      <Chip
                        label={`Phase ${entry.phase}`}
                        size="small"
                        sx={{
                          mt: 0.5,
                          height: 16,
                          fontSize: '0.6rem',
                          bgcolor: 'rgba(0,0,0,0.06)',
                        }}
                      />
                    )}
                  </Box>
                </Box>
              </Box>
            ))}
          </Box>
        )}
      </Box>
    </Box>
  )
}

export default ActivityLog
