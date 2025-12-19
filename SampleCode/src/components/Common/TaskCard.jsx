import React from 'react'
import {
  Card,
  CardContent,
  Typography,
  Box,
  Button,
  Chip,
  CircularProgress,
  Tooltip,
} from '@mui/material'
import {
  PlayArrow,
  Refresh,
  CheckCircle,
  Lock,
  Error as ErrorIcon,
} from '@mui/icons-material'
import StatusBadge from './StatusBadge'

const TaskCard = ({
  id,
  name,
  status = 'pending',
  loading = false,
  result = null,
  error = null,
  dependencies = [],
  dependenciesMet = true,
  onExecute,
  compact = false,
}) => {
  const isCompleted = result !== null && !error
  const hasError = !!error

  const getStatusText = () => {
    if (loading) return '실행중...'
    if (hasError) return '오류 발생'
    if (isCompleted) {
      if (Array.isArray(result)) return `${result.length}건 완료`
      return '완료'
    }
    if (!dependenciesMet) return '의존성 대기'
    return '대기중'
  }

  const getBorderColor = () => {
    if (loading) return 'warning.main'
    if (hasError) return 'error.main'
    if (isCompleted) return 'success.main'
    return 'divider'
  }

  if (compact) {
    return (
      <Card
        variant="compact"
        sx={{
          borderLeft: '3px solid',
          borderLeftColor: getBorderColor(),
          height: '100%',
        }}
      >
        <CardContent sx={{ p: 1.5, '&:last-child': { pb: 1.5 } }}>
          <Box
            sx={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              gap: 1,
            }}
          >
            <Box sx={{ minWidth: 0, flexGrow: 1 }}>
              <Typography
                variant="body2"
                sx={{
                  fontWeight: 500,
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap',
                }}
              >
                {name}
              </Typography>
              <Typography
                variant="caption"
                sx={{ color: 'text.secondary', fontSize: '0.7rem' }}
              >
                {getStatusText()}
              </Typography>
            </Box>

            {loading ? (
              <CircularProgress size={24} />
            ) : !dependenciesMet ? (
              <Tooltip title={`필요: ${dependencies.join(', ')}`}>
                <Lock sx={{ color: 'grey.400', fontSize: 20 }} />
              </Tooltip>
            ) : (
              <Button
                size="small"
                variant={isCompleted ? 'outlined' : 'contained'}
                onClick={onExecute}
                disabled={loading}
                sx={{ minWidth: 'auto', px: 1.5 }}
              >
                {isCompleted ? <Refresh fontSize="small" /> : <PlayArrow fontSize="small" />}
              </Button>
            )}
          </Box>

          {hasError && (
            <Typography
              variant="caption"
              sx={{
                color: 'error.main',
                display: 'block',
                mt: 0.5,
                fontSize: '0.65rem',
              }}
            >
              {error}
            </Typography>
          )}
        </CardContent>
      </Card>
    )
  }

  return (
    <Card
      sx={{
        height: '100%',
        borderColor: getBorderColor(),
        borderWidth: 1,
        borderStyle: 'solid',
      }}
    >
      <CardContent>
        <Box
          sx={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'flex-start',
            mb: 2,
          }}
        >
          <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
            {name}
          </Typography>
          {loading ? (
            <CircularProgress size={20} />
          ) : isCompleted ? (
            <CheckCircle sx={{ color: 'success.main', fontSize: 20 }} />
          ) : hasError ? (
            <ErrorIcon sx={{ color: 'error.main', fontSize: 20 }} />
          ) : !dependenciesMet ? (
            <Lock sx={{ color: 'grey.400', fontSize: 20 }} />
          ) : null}
        </Box>

        <Box sx={{ mb: 2 }}>
          <StatusBadge
            status={
              loading
                ? 'in_progress'
                : hasError
                ? 'error'
                : isCompleted
                ? 'completed'
                : 'pending'
            }
            loading={loading}
            customLabel={getStatusText()}
          />
        </Box>

        {dependencies.length > 0 && !dependenciesMet && (
          <Box sx={{ mb: 2 }}>
            <Typography
              variant="caption"
              sx={{ color: 'text.secondary', display: 'block', mb: 0.5 }}
            >
              필요한 데이터:
            </Typography>
            <Box sx={{ display: 'flex', gap: 0.5, flexWrap: 'wrap' }}>
              {dependencies.map((dep) => (
                <Chip
                  key={dep}
                  label={dep}
                  size="small"
                  sx={{ height: 20, fontSize: '0.65rem' }}
                />
              ))}
            </Box>
          </Box>
        )}

        {hasError && (
          <Typography
            variant="caption"
            sx={{
              color: 'error.main',
              display: 'block',
              mb: 2,
              p: 1,
              bgcolor: 'error.50',
              borderRadius: 1,
            }}
          >
            {error}
          </Typography>
        )}

        {isCompleted && Array.isArray(result) && (
          <Chip
            label={`${result.length}건 조회 완료`}
            color="success"
            size="small"
            sx={{ mb: 2 }}
          />
        )}

        <Button
          fullWidth
          variant={isCompleted ? 'outlined' : 'contained'}
          startIcon={
            loading ? (
              <CircularProgress size={16} />
            ) : isCompleted ? (
              <Refresh />
            ) : (
              <PlayArrow />
            )
          }
          onClick={onExecute}
          disabled={loading || !dependenciesMet}
        >
          {loading ? '실행중...' : isCompleted ? '재실행' : '실행'}
        </Button>
      </CardContent>
    </Card>
  )
}

export default TaskCard
