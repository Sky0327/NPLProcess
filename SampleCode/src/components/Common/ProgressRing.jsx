import React from 'react'
import { Box, CircularProgress, Typography } from '@mui/material'

const ProgressRing = ({
  value = 0,
  size = 60,
  thickness = 4,
  color = 'primary',
  showLabel = true,
  label,
}) => {
  return (
    <Box
      sx={{
        position: 'relative',
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      {/* Background circle */}
      <CircularProgress
        variant="determinate"
        value={100}
        size={size}
        thickness={thickness}
        sx={{
          color: 'grey.200',
          position: 'absolute',
        }}
      />
      {/* Progress circle */}
      <CircularProgress
        variant="determinate"
        value={value}
        size={size}
        thickness={thickness}
        color={color}
        sx={{
          '& .MuiCircularProgress-circle': {
            strokeLinecap: 'round',
          },
        }}
      />
      {showLabel && (
        <Box
          sx={{
            position: 'absolute',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
          }}
        >
          <Typography
            variant="caption"
            sx={{
              fontWeight: 700,
              fontSize: size * 0.22,
              lineHeight: 1,
            }}
          >
            {label || `${Math.round(value)}%`}
          </Typography>
        </Box>
      )}
    </Box>
  )
}

export const MiniProgressRing = ({ value = 0, size = 32, color = 'primary' }) => {
  return (
    <Box
      sx={{
        position: 'relative',
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      <CircularProgress
        variant="determinate"
        value={100}
        size={size}
        thickness={3}
        sx={{
          color: 'grey.200',
          position: 'absolute',
        }}
      />
      <CircularProgress
        variant="determinate"
        value={value}
        size={size}
        thickness={3}
        color={color}
        sx={{
          '& .MuiCircularProgress-circle': {
            strokeLinecap: 'round',
          },
        }}
      />
      <Typography
        variant="caption"
        sx={{
          position: 'absolute',
          fontWeight: 600,
          fontSize: '0.6rem',
        }}
      >
        {Math.round(value)}
      </Typography>
    </Box>
  )
}

export default ProgressRing
