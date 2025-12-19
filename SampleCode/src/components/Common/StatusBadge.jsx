import React from 'react'
import { Chip, CircularProgress, Box } from '@mui/material'
import {
  CheckCircle,
  RadioButtonUnchecked,
  Error as ErrorIcon,
  PlayArrow,
} from '@mui/icons-material'

const statusConfig = {
  pending: {
    label: '대기',
    color: 'default',
    bgcolor: '#e9ecef',
    textColor: '#6c757d',
    icon: RadioButtonUnchecked,
  },
  in_progress: {
    label: '진행중',
    color: 'warning',
    bgcolor: '#fff3e0',
    textColor: '#e65100',
    icon: PlayArrow,
  },
  completed: {
    label: '완료',
    color: 'success',
    bgcolor: '#e8f5e9',
    textColor: '#2e7d32',
    icon: CheckCircle,
  },
  error: {
    label: '오류',
    color: 'error',
    bgcolor: '#ffebee',
    textColor: '#c62828',
    icon: ErrorIcon,
  },
}

const StatusBadge = ({
  status = 'pending',
  loading = false,
  size = 'small',
  showLabel = true,
  customLabel,
}) => {
  const config = statusConfig[status] || statusConfig.pending
  const IconComponent = config.icon

  if (loading) {
    return (
      <Chip
        size={size}
        icon={<CircularProgress size={12} sx={{ color: config.textColor }} />}
        label={showLabel ? (customLabel || '처리중...') : undefined}
        sx={{
          bgcolor: '#fff3e0',
          color: '#e65100',
          '& .MuiChip-icon': {
            color: '#e65100',
          },
        }}
      />
    )
  }

  return (
    <Chip
      size={size}
      icon={
        showLabel ? (
          <IconComponent sx={{ fontSize: 14, color: config.textColor }} />
        ) : undefined
      }
      label={showLabel ? (customLabel || config.label) : undefined}
      sx={{
        bgcolor: config.bgcolor,
        color: config.textColor,
        fontWeight: 500,
        '& .MuiChip-icon': {
          color: config.textColor,
        },
        ...((!showLabel) && {
          width: 24,
          height: 24,
          p: 0,
          '& .MuiChip-label': {
            display: 'none',
          },
        }),
      }}
    />
  )
}

export const StatusDot = ({ status = 'pending', size = 8 }) => {
  const colors = {
    pending: '#9e9e9e',
    in_progress: '#ff9800',
    completed: '#4caf50',
    error: '#f44336',
  }

  return (
    <Box
      sx={{
        width: size,
        height: size,
        borderRadius: '50%',
        bgcolor: colors[status] || colors.pending,
        flexShrink: 0,
      }}
    />
  )
}

export default StatusBadge
